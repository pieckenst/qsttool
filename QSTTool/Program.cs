using OpenMcdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace QSTTool
{
    // Type of extracted OLE object or Picture
    public enum OleFileType
    {
        Unknown,
        ImageWmf,       // Windows Metafile
        ImageEmf,       // Enhanced Metafile
        ImagePng,       // PNG image
        ImageJpeg,      // JPEG image
        ImageDib,       // Device Independent Bitmap (often wrapped in WMF/EMF)
        Equation,       // OLE Object identified as Equation
        GenericOle      // Other OLE objects or unknown binary data
    }

    // Represents extracted OLE object or Picture data
    public class OleObject
    {
        public string ClassName { get; set; } // Primarily for OLE Objects
        public byte[] Data { get; set; } // Extracted data (e.g., WMF/EMF/PNG/JPG or raw OLE)
        public OleFileType ExtractedType { get; set; } = OleFileType.Unknown;
        public string SuggestedFileName { get; set; } // Relative filename for export
    }

    // Result of parsing RTF, includes text and any extracted OLE objects
    public class RtfParseResult
    {
        public string Text { get; set; } = string.Empty;
        public List<OleObject> OleObjects { get; set; } = new List<OleObject>();

        // Generates text representation including descriptive placeholders
        public string GetTextWithPlaceholders()
        {
            StringBuilder sb = new StringBuilder(Text);
            foreach (var ole in OleObjects)
            {
                string typeDesc = ole.ExtractedType switch
                {
                    OleFileType.ImageWmf => "Image (WMF)",
                    OleFileType.ImageEmf => "Image (EMF)",
                    OleFileType.ImagePng => "Image (PNG)",
                    OleFileType.ImageJpeg => "Image (JPEG)",
                    OleFileType.ImageDib => "Image (DIB)",
                    OleFileType.Equation => "Equation",
                    OleFileType.GenericOle => "OLE Object",
                    _ => "Unknown Object"
                };
                sb.Append($" [{typeDesc}: {ole.SuggestedFileName ?? ole.ClassName}] ");
            }
            return sb.ToString().Trim(); // Trim at the end
        }

        // Override ToString to use the new placeholder logic
        public override string ToString()
        {
            return GetTextWithPlaceholders();
        }
    }

    class Program
    {
        // Путь к подпапке "export"
        private static string ExportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "export");
        private static readonly Encoding Win1251 = Encoding.GetEncoding(1251);

        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.Unicode;

            while (true)
            {
                ShowMenu();
                var choice = Console.ReadLine();
                switch (choice)
                {
                    case "1": ProcessFiles(); break;
                    case "2": CreateNewFile(); break;
                    case "3": ExportOptions(); break;
                    case "0": return;
                    default: Console.WriteLine("Неверный выбор!"); break;
                }
            }
        }

        static void ShowMenu()
        {
            Console.WriteLine("\n=== Меню ===");
            Console.WriteLine("1 - Обработать файлы (*.qst)");
            Console.WriteLine("2 - Создать новый файл (*.qst)");
            Console.WriteLine("3 - Экспорт данных из *.qst");
            Console.WriteLine("0 - Выход");
            Console.Write("Выберите действие: ");
        }

        static void ProcessFiles()
        {
            Console.Write("Введите путь к файлу *.qst или директории с *.qst файлами: ");
            string path = Console.ReadLine()?.Trim('"'); // Handle quoted paths
            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine("Путь не указан.");
                return;
            }

            var allQuestions = new List<Question>();

            try
            {
                if (Directory.Exists(path))
                {
                    Console.WriteLine($"Обработка директории: {path}");
                    foreach (var file in Directory.GetFiles(path, "*.qst"))
                    {
                        Console.WriteLine($"  Обработка файла: {Path.GetFileName(file)}");
                        allQuestions.AddRange(ProcessFile(file));
                    }
                }
                else if (File.Exists(path))
                {
                    if (path.EndsWith(".qst", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Обработка файла: {Path.GetFileName(path)}");
                        allQuestions = ProcessFile(path);
                    }
                    else
                    {
                        Console.WriteLine("Указанный файл не является *.qst файлом.");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Указанный путь не существует!");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при обработке пути '{path}': {ex.Message}");
                return;
            }


            Console.WriteLine($"\nВсего найдено вопросов: {allQuestions.Count}");
            if (allQuestions.Count > 0)
            {
                ExportData(allQuestions); // Export all collected questions

                // Display parsed content (optional, can be long)
                Console.WriteLine("\nХотите отобразить содержимое вопросов? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLower() == "y")
                {
                    foreach (var q in allQuestions)
                    {
                        Console.WriteLine($"\n--- Вопрос (GUID: {q.Guid}) ---");
                        Console.WriteLine(q.ParsedText.Text); // Display parsed text
                        if (q.ParsedText.OleObjects.Any())
                        {
                            Console.WriteLine("  Обнаружены OLE объекты:");
                            foreach (var ole in q.ParsedText.OleObjects)
                            {
                                Console.WriteLine($"    - Тип: {ole.ClassName}, Размер: {ole.Data?.Length ?? 0} байт");
                            }
                        }
                        Console.WriteLine("  --- Ответы ---");
                        foreach (var a in q.Answers)
                        {
                            Console.WriteLine($"    Ответ (GUID: {a.Guid}, {(a.IsRight ? "Правильный" : "Неправильный")}):");
                            Console.WriteLine($"      {a.ParsedText.Text}"); // Display parsed text
                            if (a.ParsedText.OleObjects.Any())
                            {
                                Console.WriteLine("      Обнаружены OLE объекты в ответе:");
                                foreach (var ole in a.ParsedText.OleObjects)
                                {
                                    Console.WriteLine($"        - Тип: {ole.ClassName}, Размер: {ole.Data?.Length ?? 0} байт");
                                }
                            }
                        }
                        Console.WriteLine(new string('-', 20));
                    }
                }
            }
        }

        static List<Question> ProcessFile(string filePath)
        {
            var questions = new List<Question>();
            string xmlContent = null;
            Stream contentStream = null;
            CompoundFile cf = null;
            ZipArchive archive = null;
            string fileLoadMethod = "Unknown"; // For logging

            try
            {
                // 1) Попытка открыть как ZIP
                try
                {
                    archive = ZipFile.OpenRead(filePath);
                    fileLoadMethod = "ZIP";
                    Console.WriteLine($"    -> Успешно открыт как ZIP.");
                    var entry = archive.GetEntry("Questions.xml") ?? archive.GetEntry("Content/Questions.xml");
                    if (entry != null)
                    {
                        Console.WriteLine($"    -> Найден XML-файл: {entry.FullName}");
                        contentStream = entry.Open();
                    }
                    else
                    {
                        Console.WriteLine("    -> XML-файл ('Questions.xml' или 'Content/Questions.xml') не найден в ZIP.");
                        // Log entries found for debugging
                        var entryNames = archive.Entries.Select(e => e.FullName).ToList();
                        Console.WriteLine($"    -> Найденные файлы в архиве: {string.Join(", ", entryNames)}");
                    }
                }
                catch (InvalidDataException)
                {
                    archive?.Dispose();
                    archive = null;
                    Console.WriteLine("    -> Не является ZIP архивом (InvalidDataException).");
                    fileLoadMethod = "Not ZIP";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    -> Ошибка чтения ZIP {Path.GetFileName(filePath)}: {ex.Message}");
                    archive?.Dispose();
                    return questions;
                }

                // 2) Если не ZIP или XML не найден в ZIP, попытка через Compound File (OLE)
                if (contentStream == null)
                {
                    // Only dispose ZIP archive if it was successfully opened but didn't contain the XML
                    if (fileLoadMethod == "ZIP")
                    {
                        archive?.Dispose();
                        archive = null;
                    }

                    Console.WriteLine("    -> Попытка открыть как Compound File (CFB)...");
                    try
                    {
                        // Check if file exists before attempting to open as CFB
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"    -> Файл не найден для открытия как CFB: {filePath}");
                            return questions;
                        }
                        cf = new CompoundFile(filePath, CFSUpdateMode.ReadOnly, CFSConfiguration.Default);
                        fileLoadMethod = "CFB";
                        Console.WriteLine("    -> Успешно открыт как CFB.");
                        var streamEntry = cf.RootStorage.TryGetStream("Questions") ?? cf.RootStorage.TryGetStream("Content/Questions");
                        if (streamEntry != null)
                        {
                            Console.WriteLine($"    -> Найден поток данных CFB: {streamEntry.Name}");
                            var data = streamEntry.GetData();
                            contentStream = new MemoryStream(data);
                        }
                        else
                        {
                            Console.WriteLine("    -> Поток данных ('Questions' или 'Content/Questions') не найден в CFB.");
                        }
                    }
                    catch (Exception exWhenOpeningCfb) // Catch exceptions during CFB processing
                    {
                        cf?.Dispose();
                        cf = null;
                        Console.WriteLine($"    -> Ошибка при попытке чтения как CFB: {exWhenOpeningCfb.Message}");
                        fileLoadMethod = "Not CFB";
                    }
                }

                // 3) Если не ZIP/CFB или данные не найдены, попытка прямого чтения как XML (текст)
                if (contentStream == null)
                {
                    // Only dispose CFB if it was successfully opened but didn't contain the stream
                    if (fileLoadMethod == "CFB")
                    {
                        cf?.Dispose();
                        cf = null;
                    }
                    Console.WriteLine("    -> Попытка чтения файла как простого текста (XML)...");
                    fileLoadMethod = "Text";
                    try
                    {
                        // ReadAllText might lock the file, use FileStream for safer read sharing
                        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            contentStream = new MemoryStream();
                            fs.CopyTo(contentStream);
                            contentStream.Position = 0;
                            Console.WriteLine("    -> Файл успешно прочитан как текст.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"    -> Не удалось прочитать файл {Path.GetFileName(filePath)} как текст: {ex.Message}");
                        // Don't return yet, final check for xmlContent below
                        fileLoadMethod = "Not Text";
                    }
                }

                // Если поток получен (из ZIP, CFB или файла), читаем XML
                if (contentStream != null)
                {
                    using (var reader = new StreamReader(contentStream, Win1251, detectEncodingFromByteOrderMarks: true)) // Use Win1251 by default
                    {
                        xmlContent = reader.ReadToEnd();
                        Console.WriteLine($"    -> XML содержимое успешно прочитано из потока (метод: {fileLoadMethod}).");
                    }
                }
            }
            finally
            {
                // Ensure disposal happens correctly
                contentStream?.Dispose(); // Dispose MemoryStream if created
                cf?.Dispose();
                archive?.Dispose();
            }

            // Если xmlContent все еще null, файл не удалось прочитать ни одним способом
            if (xmlContent == null)
            {
                Console.WriteLine($"    -> Не удалось извлечь XML из файла {Path.GetFileName(filePath)} ни одним из методов ({fileLoadMethod}).");
                return questions;
            }

            // Парсим XML
            try
            {
                var doc = new XmlDocument();
                // Prevent XML External Entity (XXE) attacks
                doc.XmlResolver = null;
                doc.LoadXml(xmlContent);

                var questionNodes = doc.SelectNodes("/Questions/Question");
                if (questionNodes == null || questionNodes.Count == 0)
                {
                    // Try finding nodes anywhere if not directly under root
                    questionNodes = doc.SelectNodes("//Question");
                }

                if (questionNodes == null || questionNodes.Count == 0)
                {
                    Console.WriteLine($"    -> В XML-содержимом файла {Path.GetFileName(filePath)} не найдены узлы 'Question'.");
                    return questions;
                }

                int questionIndex = 0;
                foreach (XmlNode questionNode in questionNodes)
                {
                    string questionGuid = questionNode.Attributes?["Guid"]?.Value ?? Guid.NewGuid().ToString();
                    string questionRtf = questionNode.SelectSingleNode("Text")?.InnerText ?? string.Empty;

                    var q = new Question
                    {
                        Guid = questionGuid,
                        OriginalRtf = questionRtf,
                        // Pass context for filename generation
                        ParsedText = ParseRtf(questionRtf, $"q{questionIndex}")
                    };

                    var answerNodes = questionNode.SelectNodes("Answer");
                    if (answerNodes != null)
                    {
                        int answerIndex = 0;
                        foreach (XmlNode answerNode in answerNodes)
                        {
                            string answerGuid = answerNode.Attributes?["Guid"]?.Value ?? Guid.NewGuid().ToString();
                            string answerRtf = answerNode.SelectSingleNode("Text")?.InnerText ?? string.Empty;
                            bool isRight = string.Equals(answerNode.Attributes?["IsRight"]?.Value, "True", StringComparison.OrdinalIgnoreCase);

                            var a = new Answer
                            {
                                Guid = answerGuid,
                                IsRight = isRight,
                                OriginalRtf = answerRtf,
                                // Pass context for filename generation
                                ParsedText = ParseRtf(answerRtf, $"q{questionIndex}_a{answerIndex}")
                            };
                            q.Answers.Add(a);
                            answerIndex++;
                        }
                    }

                    // Examiner logic: only add questions with at least one correct answer
                    if (q.Answers.Exists(a => a.IsRight))
                    {
                        questions.Add(q);
                    }
                    else
                    {
                        Console.WriteLine($"    -> Предупреждение: Вопрос GUID {q.Guid} в файле {Path.GetFileName(filePath)} пропущен (нет правильных ответов).");
                    }
                    questionIndex++;
                }
                Console.WriteLine($"    -> Успешно разобрано {questions.Count} вопросов из XML.");
            }
            catch (XmlException ex)
            {
                Console.WriteLine($"    -> Ошибка парсинга XML в файле {Path.GetFileName(filePath)}: {ex.Message} (Строка: {ex.LineNumber}, Позиция: {ex.LinePosition})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    -> Непредвиденная ошибка парсинга XML {Path.GetFileName(filePath)}: {ex.Message}");
            }

            return questions;
        }

        // Enhanced RTF Parser with explicit OLE and Picture data handling
        static RtfParseResult ParseRtf(string rtf, string objectNamePrefix)
        {
            var result = new RtfParseResult();
            if (string.IsNullOrWhiteSpace(rtf) || !rtf.TrimStart().StartsWith("{\\rtf"))
            {
                result.Text = rtf ?? string.Empty;
                return result;
            }

            var textBuilder = new StringBuilder();
            var groupStack = new Stack<RtfGroupState>();
            groupStack.Push(new RtfGroupState { IsRoot = true }); // Root state
            int objectIndex = 0; // Combined index for OLE objects and pictures

            int i = 0;
            while (i < rtf.Length)
            {
                char c = rtf[i];
                var currentState = groupStack.Peek();

                // 1. Handle ignored content first
                if (currentState.IgnoreContent)
                {
                    if (c == '{')
                    {
                        // Push a new ignored state
                        groupStack.Push(new RtfGroupState(currentState) { IgnoreLevel = currentState.IgnoreLevel + 1 });
                    }
                    else if (c == '}')
                    {
                        if (groupStack.Count > 1) // Don't pop the root
                        {
                            var poppedState = groupStack.Pop();
                            // If we are leaving the top-level ignored group, reset ignore on parent
                            if (poppedState.IgnoreLevel == 1 && groupStack.Count > 0)
                            {
                                groupStack.Peek().IgnoreContent = false;
                            }
                        }
                    }
                    i++;
                    continue; // Skip normal processing
                }

                // 2. Normal processing switch
                switch (c)
                {
                    case '{':
                        // Flush pending text before starting new group
                        textBuilder.Append(currentState.PendingText);
                        currentState.PendingText.Clear();

                        // Create and push the new state, inheriting relevant flags
                        var newState = new RtfGroupState(currentState);

                        // Check for ignorable group flag set by previous \\*
                        if (currentState.NextGroupIsIgnorable)
                        {
                            newState.IgnoreContent = true;
                            newState.IgnoreLevel = 1;
                            currentState.NextGroupIsIgnorable = false; // Reset flag on parent
                        }

                        // Handle objclass text consumption inheritance
                        if (currentState.ConsumeNextTextAs == RtfConsumeTarget.ObjectClass)
                        {
                            newState.ConsumeNextTextAs = RtfConsumeTarget.ObjectClass;
                            currentState.ConsumeNextTextAs = RtfConsumeTarget.None;
                        }

                        groupStack.Push(newState);
                        i++;
                        break;

                    case '}':
                        // Flush pending text from the closing group
                        textBuilder.Append(currentState.PendingText);
                        currentState.PendingText.Clear();

                        if (groupStack.Count > 1) // Don't pop the root state
                        {
                            var finishedState = groupStack.Pop();

                            // --- Process collected data if any ---

                            // A. Process OLE Object Data
                            if (finishedState.IsObject && finishedState.ObjectDataHex.Length > 0)
                            {
                                var oleObject = new OleObject { ClassName = finishedState.ObjectClass };
                                try
                                {
                                    var rawOleData = HexStringToByteArray(finishedState.ObjectDataHex.ToString());
                                    AnalyzeAndExtractOleData(rawOleData, oleObject, objectNamePrefix, objectIndex++);
                                    result.OleObjects.Add(oleObject);
                                    // Placeholder is generated later by RtfParseResult.GetTextWithPlaceholders()
                                }
                                catch (FormatException hexEx)
                                {
                                    Console.WriteLine($"    -> Warning: Invalid hex in OLE Data for {oleObject.ClassName}: {hexEx.Message}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"    -> Warning: Error decoding/analyzing OLE Data for {oleObject.ClassName}: {ex.Message}");
                                    // Fallback: Store raw hex if conversion failed but we want to keep track? No, Analyze handles fallback.
                                }
                            }
                            // B. Process Picture Data
                            else if (finishedState.IsPicture && finishedState.PictureDataHex.Length > 0)
                            {
                                var picObject = new OleObject { ExtractedType = finishedState.DetectedPictureType };
                                try
                                {
                                    picObject.Data = HexStringToByteArray(finishedState.PictureDataHex.ToString());
                                    // Determine filename based on detected type
                                    string ext = picObject.ExtractedType switch
                                    {
                                        OleFileType.ImageWmf => ".wmf",
                                        OleFileType.ImageEmf => ".emf",
                                        OleFileType.ImagePng => ".png",
                                        OleFileType.ImageJpeg => ".jpeg",
                                        OleFileType.ImageDib => ".dib", // Or .bmp? DIB is complex.
                                        _ => ".bin" // Fallback
                                    };
                                    picObject.SuggestedFileName = $"{objectNamePrefix}_pic{objectIndex++}{ext}";
                                    result.OleObjects.Add(picObject); // Add picture as an OleObject
                                    Console.WriteLine($"    -> Extracted Picture data ({picObject.ExtractedType}, {picObject.Data?.Length ?? 0} bytes) as {picObject.SuggestedFileName}");
                                }
                                catch (FormatException hexEx)
                                {
                                    Console.WriteLine($"    -> Warning: Invalid hex in Picture Data: {hexEx.Message}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"    -> Warning: Error processing Picture Data: {ex.Message}");
                                }
                            }

                            // --- Handle text consumption across group boundaries (e.g., for objclass) ---
                            if (groupStack.Count > 0) // Check if stack is not empty after pop
                            {
                                var parentState = groupStack.Peek();
                                if (parentState.ConsumeNextTextAs != RtfConsumeTarget.None)
                                {
                                    // Text might be in the finished group's PENDING text (which we just flushed)
                                    // Or it might be plain text immediately following the group. The spec is ambiguous.
                                    // Let's assume for now that objclass text must be WITHIN a {} group after the tag.
                                    // If finishedState.PendingText was non-empty, we should use that. (This was flushed to textBuilder though)
                                    // Revisit this if \objclass {ClassName} doesn't work.
                                    // Alternative: Check plain text immediately following '}'

                                    // Correction: Let's handle objclass text if it was inside the {}
                                    string groupText = finishedState.InternalText.ToString().Trim(); // Use InternalText collected within the group
                                    if (!string.IsNullOrEmpty(groupText))
                                    {
                                        if (parentState.ConsumeNextTextAs == RtfConsumeTarget.ObjectClass)
                                        {
                                            parentState.ObjectClass = groupText; // Assign class name to PARENT
                                            parentState.ConsumeNextTextAs = RtfConsumeTarget.None; // Consumed
                                        }
                                    }
                                }
                            }
                        }
                        i++;
                        break;

                    case '\\':
                        i++; // Move past backslash
                        if (i >= rtf.Length) break;

                        char nextChar = rtf[i];
                        if (nextChar == '\\' || nextChar == '{' || nextChar == '}')
                        {
                            // Escaped character, treat as literal text ONLY IF not reading binary data
                            if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                            {
                                currentState.PendingText.Append(nextChar);
                                currentState.InternalText.Append(nextChar); // Also add to internal text
                            }
                            i++;
                        }
                        else if (nextChar == '\'') // Hex escape character \'xx
                        {
                            i++; // Move past quote
                            if (i + 1 < rtf.Length)
                            {
                                string hex = rtf.Substring(i, 2);
                                try
                                {
                                    byte decodedByte = Convert.ToByte(hex, 16);
                                    // Append decoded char ONLY IF not reading binary data
                                    if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                                    {
                                        char decodedChar = Win1251.GetString(new[] { decodedByte })[0];
                                        currentState.PendingText.Append(decodedChar);
                                        currentState.InternalText.Append(decodedChar); // Also add to internal text
                                    }
                                    i += 2;
                                }
                                catch (FormatException) { i++; /* Skip malformed hex byte */ }
                                catch (ArgumentOutOfRangeException) { i++; /* Skip malformed hex byte */ }
                            }
                            else { i++; } // Skip quote if at end
                        }
                        else if (nextChar == '*') // Ignorable destination group marker
                        {
                            i++;
                            currentState.NextGroupIsIgnorable = true;
                        }
                        else if (char.IsLetter(nextChar)) // Control Word
                        {
                            // Flush pending text before processing control word
                            // Important: Do NOT flush if we are actively reading binary data
                            if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                            {
                                textBuilder.Append(currentState.PendingText);
                                currentState.PendingText.Clear();
                            }

                            int start = i;
                            while (i < rtf.Length && char.IsLetter(rtf[i])) i++;
                            string controlWord = rtf.Substring(start, i - start);

                            // Read optional parameter
                            bool negative = false;
                            int paramStart = i;
                            if (i < rtf.Length && rtf[i] == '-') { negative = true; i++; paramStart++; }
                            while (i < rtf.Length && char.IsDigit(rtf[i])) i++;
                            int? parameter = null;
                            if (i > paramStart)
                            {
                                if (int.TryParse(rtf.Substring(paramStart, i - paramStart), out int pVal))
                                    parameter = negative ? -pVal : pVal;
                            }

                            // Optional space delimiter after word/parameter
                            if (i < rtf.Length && rtf[i] == ' ') i++;

                            // Process known control words
                            ProcessControlWord(controlWord, parameter, textBuilder, groupStack);
                        }
                        else if (nextChar == '~') // Non-breaking space
                        {
                            if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                            {
                                currentState.PendingText.Append('\u00A0');
                                currentState.InternalText.Append('\u00A0');
                            }
                            i++;
                            if (i < rtf.Length && rtf[i] == ' ') i++; // Skip optional space
                        }
                        else if (nextChar == '-') // Optional hyphen (treated as text)
                        {
                            if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                            {
                                currentState.PendingText.Append('-');
                                currentState.InternalText.Append('-');
                            }
                            i++;
                            if (i < rtf.Length && rtf[i] == ' ') i++; // Skip optional space
                        }
                        else if (nextChar == '_') // Non-breaking hyphen (treated as text)
                        {
                            if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                            {
                                currentState.PendingText.Append('-');
                                currentState.InternalText.Append('-');
                            }
                            i++;
                            if (i < rtf.Length && rtf[i] == ' ') i++; // Skip optional space
                        }
                        else
                        {
                            // Unknown symbol after backslash, just skip it and the symbol
                            i++;
                        }
                        break;

                    // Treat line breaks as spaces or ignore, depending on context?
                    // RTF spec usually ignores them unless preceded by specific control words.
                    case '\r':
                    case '\n':
                        i++; // Skip CR/LF
                        break;
                    // Treat tab as space for plain text extraction
                    case '\t':
                        if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                        {
                            currentState.PendingText.Append('\t');
                            currentState.InternalText.Append('\t');
                        }
                        i++;
                        break;

                    default: // Plain character
                        // Append ONLY if not inside binary data sections
                        if (currentState.IsReadingObjectData)
                        {
                            if (IsHexDigit(c)) currentState.ObjectDataHex.Append(c);
                        }
                        else if (currentState.IsReadingPictureData)
                        {
                            if (IsHexDigit(c)) currentState.PictureDataHex.Append(c);
                        }
                        else if (currentState.ConsumeNextTextAs != RtfConsumeTarget.None)
                        {
                            // Text meant for a specific purpose (like objclass) within a group
                            currentState.PendingText.Append(c); // Append to current group's pending text
                            currentState.InternalText.Append(c); // Append to internal text as well
                        }
                        else
                        {
                            // Regular text character
                            currentState.PendingText.Append(c);
                            currentState.InternalText.Append(c); // Append to internal text
                        }
                        i++;
                        break;
                }
            }

            // Append any remaining text from the root state
            textBuilder.Append(groupStack.Peek().PendingText);

            result.Text = textBuilder.ToString().Trim(); // Trim final result only
            return result;
        }

        // Updated ProcessControlWord - no longer passes result directly
        private static void ProcessControlWord(string word, int? param, StringBuilder textBuilder, Stack<RtfGroupState> groupStack)
        {
            var currentState = groupStack.Peek();

            // Control words should generally not affect output if reading binary data,
            // EXCEPT for the words that TERMINATE binary reading (handled by '}' processing).
            // Special handling for words STARTING binary reading is needed here.

            switch (word)
            {
                // --- State Changers ---
                case "object":
                    currentState.IsObject = true;
                    currentState.IsReadingObjectData = false; // Reset flag
                    currentState.ObjectDataHex.Clear();
                    currentState.ObjectClass = "Unknown"; // Default class name
                    break;
                case "objclass":
                    // Expect next text block (likely within {}) to be the class name
                    currentState.ConsumeNextTextAs = RtfConsumeTarget.ObjectClass;
                    break;
                case "objdata":
                    if (currentState.IsObject) // Only makes sense within an object group
                    {
                        currentState.IsReadingObjectData = true; // Start reading hex data
                        currentState.ObjectDataHex.Clear();
                    }
                    break;
                case "pict":
                    currentState.IsPicture = true;
                    currentState.IsReadingPictureData = true; // Assume picture data follows immediately
                    currentState.PictureDataHex.Clear();
                    currentState.DetectedPictureType = OleFileType.Unknown; // Reset detected type
                    break;
                // Picture type specifiers (often appear within \\pict group before data)
                case "wmetafile": currentState.DetectedPictureType = OleFileType.ImageWmf; break;
                case "emfblip": currentState.DetectedPictureType = OleFileType.ImageEmf; break;
                case "pngblip": currentState.DetectedPictureType = OleFileType.ImagePng; break;
                case "jpegblip": currentState.DetectedPictureType = OleFileType.ImageJpeg; break;
                case "dibitmap": // Or just "dib"? Check spec. Assuming DIB data follows.
                    currentState.DetectedPictureType = OleFileType.ImageDib;
                    break;

                // --- Text Outputters (Only if NOT reading binary) ---
                case "par":
                case "line":
                    if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                    {
                        textBuilder.Append(Environment.NewLine);
                        currentState.PendingText.Clear(); // New line flushes pending buffer? Yes.
                        currentState.InternalText.Append(Environment.NewLine); // Keep track internally too
                    }
                    break;
                case "tab":
                    if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData)
                    {
                        currentState.PendingText.Append('\t');
                        currentState.InternalText.Append('\t');
                    }
                    break;
                case "u": // Unicode character \\uN (N is signed decimal)
                    if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData && param.HasValue)
                    {
                        try
                        {
                            // N is unicode value, but might be negative for fallback?
                            // Standard says N is signed 16-bit.
                            // Need to handle potential negative N for fallback with \\uc
                            int unicodeValue = param.Value;
                            if (unicodeValue < 0) unicodeValue += 65536; // Handle signed wrap-around if necessary?

                            char unicodeChar = Convert.ToChar(unicodeValue);
                            currentState.PendingText.Append(unicodeChar);
                            currentState.InternalText.Append(unicodeChar);

                            // Skip specified number of ASCII fallback chars following \\uN
                            // TODO: Implement skipping based on currentState.AsciiFallbackChars
                            // Need to advance 'i' in the main loop. This is complex.
                        }
                        catch { currentState.PendingText.Append('?'); currentState.InternalText.Append('?'); } // Fallback char
                    }
                    break;
                case "uc": // Specifies number of chars following \\uN to skip
                    currentState.AsciiFallbackChars = param ?? 1;
                    break;

                // Named escape characters (append only if not reading binary)
                case "emdash": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('—'); currentState.InternalText.Append('—'); } break; // Unicode U+2014
                case "endash": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('–'); currentState.InternalText.Append('–'); } break; // Unicode U+2013
                case "bullet": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('•'); currentState.InternalText.Append('•'); } break; // Unicode U+2022
                case "lquote": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('‘'); currentState.InternalText.Append('‘'); } break; // Unicode U+2018
                case "rquote": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('’'); currentState.InternalText.Append('’'); } break; // Unicode U+2019
                case "ldblquote": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('"'); currentState.InternalText.Append('"'); } break; // Unicode U+201C
                case "rdblquote": if (!currentState.IsReadingObjectData && !currentState.IsReadingPictureData) { currentState.PendingText.Append('"'); currentState.InternalText.Append('"'); } break; // Unicode U+201D

                // --- Ignored Control Words ---
                case "result": // Often contains cached rendering of object/picture, ignore its group content
                    currentState.NextGroupIsIgnorable = true; // Ignore the group immediately following \\result
                    break;
                // Add other known formatting/metadata words that should always be ignored for text extraction
                case "rtf":
                case "ansi":
                case "ansicpg":
                case "deff":
                case "deflang":
                case "fonttbl":
                case "colortbl":
                case "stylesheet":
                case "info":
                case "author":
                case "creatim":
                case "revtbl":
                case "operator": // More metadata
                case "b":
                case "i":
                case "ul":
                case "strike":
                case "fs":
                case "f":
                case "cf": // Basic formatting
                case "pard":
                case "sectd":
                case "plain": // Paragraph/section formatting reset
                case "qc":
                case "qj":
                case "ql":
                case "qr": // Alignment
                case "li":
                case "fi": // Indentation
                case "objw":
                case "objh":
                case "picw":
                case "pich": // Size specifiers
                case "objscalex":
                case "objscaley":
                case "picscalex":
                case "picscaley": // Scaling
                case "shppict":
                case "nonshppict": // Shape related
                case "NeXTGraphic": // Specific picture format wrapper
                case "*": // Destination marker often used with ignorable groups like \\*\\themedata
                          // Just ignore these known words, don't affect state or output
                    break;

                default:
                    // Unknown control word, standard says to ignore it
                    // Console.WriteLine($"    -> Unknown RTF control word: \\{word}"); // Optional debug output
                    break;
            }
        }

        // Analyzes raw OLE data, extracts presentation/package, sets OleObject properties
        private static void AnalyzeAndExtractOleData(byte[] rawOleData, OleObject oleObject, string namePrefix, int index)
        {
            string baseName = $"{namePrefix}_obj{index}_{oleObject.ClassName ?? "Unknown"}";
            oleObject.Data = rawOleData; // Default to raw data
            oleObject.ExtractedType = OleFileType.GenericOle; // Default type
            oleObject.SuggestedFileName = baseName + ".ole";   // Default extension

            // Tentatively mark as Equation based on class name early on
            bool isLikelyEquation = oleObject.ClassName?.StartsWith("Equation.", StringComparison.OrdinalIgnoreCase) == true;
            if (isLikelyEquation)
            {
                oleObject.ExtractedType = OleFileType.Equation;
                // Keep .ole extension for raw data initially, may change if presentation found
            }

            if (rawOleData == null || rawOleData.Length < 8) return; // Need at least header size

            try
            {
                // Avoid nested using if cfOle creation fails
                CompoundFile cfOle = null;
                try
                {
                    // Pass the MemoryStream directly, use Default config
                    cfOle = new CompoundFile(new MemoryStream(rawOleData), CFSUpdateMode.ReadOnly, CFSConfiguration.Default);
                }
                catch (Exception ex)
                {
                    // Log the error, but KEEP the Equation type if it was set.
                    Console.WriteLine($"    -> Warning: Failed to open inner OLE structure as CFB for {oleObject.ClassName}: {ex.Message}. Storing raw data.");
                    // Function returns here, oleObject retains raw data and potentially Equation type.
                    return;
                }

                // If cfOle is null here (shouldn't happen if exception wasn't caught, but defensive check)
                if (cfOle == null) return;

                using (cfOle) // Now safe to use using
                {
                    // 1. Try to get Presentation Stream (WMF/EMF)
                    var presStream = cfOle.RootStorage.TryGetStream("\\x01OlePres000");
                    if (presStream != null)
                    {
                        var presData = presStream.GetData();
                        oleObject.Data = presData; // Store presentation data instead of raw
                        // Basic check for WMF/EMF headers (can be more robust)
                        if (presData.Length > 20 && presData[0] == 0xD7 && presData[1] == 0xCD && presData[2] == 0xC6 && presData[3] == 0x9A) // WMF Placeable header
                        {
                            oleObject.ExtractedType = OleFileType.ImageWmf;
                            oleObject.SuggestedFileName = baseName + ".wmf";
                        }
                        else if (presData.Length > 40 && presData[0] == 0x01 && presData[1] == 0x00 && presData[2] == 0x00 && presData[3] == 0x00) // EMF header start
                        {
                            oleObject.ExtractedType = OleFileType.ImageEmf;
                            oleObject.SuggestedFileName = baseName + ".emf";
                        }
                        else
                        {
                            // Unknown presentation format, save as .bin?
                            oleObject.ExtractedType = OleFileType.GenericOle; // Revert to generic if unknown presentation
                            oleObject.SuggestedFileName = baseName + "_pres.bin";
                        }

                        // ***Important: Override type back to Equation if classname matches***
                        // Keep the extracted presentation data and corresponding extension (.wmf/.emf)
                        if (isLikelyEquation)
                        {
                            oleObject.ExtractedType = OleFileType.Equation;
                        }
                        Console.WriteLine($"    -> Extracted OLE Presentation stream ({oleObject.ExtractedType}, {presData.Length} bytes) for {oleObject.ClassName} as {oleObject.SuggestedFileName}");
                        return; // Prioritize presentation stream
                    }

                    // 2. Try to get Package Stream (might contain original file)
                    var packageStream = cfOle.RootStorage.TryGetStream("Package");
                    if (packageStream != null)
                    {
                        var packageData = packageStream.GetData();
                        oleObject.Data = packageData;
                        // TODO: Could try to detect file type based on magic bytes in packageData
                        string ext = DetermineExtensionFromPackage(packageData, oleObject.ClassName);
                        oleObject.ExtractedType = OleFileType.GenericOle; // Mark as generic package for now
                        oleObject.SuggestedFileName = baseName + ext;

                        // Override type if it's an equation package
                        if (isLikelyEquation)
                        {
                            oleObject.ExtractedType = OleFileType.Equation;
                            // Keep the guessed package extension
                        }
                        Console.WriteLine($"    -> Extracted OLE Package stream ({oleObject.ExtractedType}, {packageData.Length} bytes) for {oleObject.ClassName} as {oleObject.SuggestedFileName}");
                        return;
                    }

                    // 3. Add checks for other known streams if necessary (e.g., CONTENTS)

                    // No specific stream found. If it was likely an equation, keep that type.
                    if (isLikelyEquation)
                    {
                        oleObject.ExtractedType = OleFileType.Equation;
                        oleObject.SuggestedFileName = baseName + ".ole"; // Fallback to raw OLE for equations with no stream
                    }
                    else
                    {
                        // Revert to GenericOle if not an equation and no stream found
                        oleObject.ExtractedType = OleFileType.GenericOle;
                        oleObject.SuggestedFileName = baseName + ".ole";
                    }
                    Console.WriteLine($"    -> No specific stream (Presentation/Package) found in OLE for {oleObject.ClassName}. Storing raw OLE data ({oleObject.ExtractedType}).");
                }
            }
            catch (Exception ex)
            {
                // Log error but fallback to storing raw data
                Console.WriteLine($"    -> Warning: Error during OLE structure processing for {oleObject.ClassName}: {ex.Message}. Storing raw data.");
                oleObject.Data = rawOleData;
                // Preserve Equation type if it was initially identified, otherwise generic.
                if (isLikelyEquation)
                {
                    oleObject.ExtractedType = OleFileType.Equation;
                    oleObject.SuggestedFileName = baseName + ".ole"; // Keep .ole for raw fallback
                }
                else
                {
                    oleObject.ExtractedType = OleFileType.GenericOle;
                    oleObject.SuggestedFileName = baseName + ".ole";
                }
            }
        }

        // Helper to guess extension (very basic)
        private static string DetermineExtensionFromPackage(byte[] data, string className)
        {
            if (className?.Contains("Excel") == true) return ".xlsx"; // Guess based on class
            if (className?.Contains("Word") == true) return ".docx";
            if (className?.Contains("PowerPoint") == true) return ".pptx";
            // Add magic byte checks here if needed
            return ".bin"; // Default fallback
        }

        private static byte[] HexStringToByteArray(string hex)
        {
            if (hex == null) throw new ArgumentNullException(nameof(hex));
            if (hex.Length % 2 != 0) throw new ArgumentException("Hex string must have an even number of digits.", nameof(hex));

            var numberChars = hex.Length;
            var bytes = new byte[numberChars / 2];
            for (var i = 0; i < numberChars; i += 2)
            {
                try
                {
                    bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
                }
                catch (FormatException ex)
                {
                    throw new FormatException($"Invalid hex character found at position {i} in hex string.", ex);
                }
            }
            return bytes;
        }

        private static bool IsHexDigit(char c)
        {
            return (c >= '0' && c <= '9') ||
                   (c >= 'a' && c <= 'f') ||
                   (c >= 'A' && c <= 'F');
        }

        static void ExportData(List<Question> questions)
        {
            if (questions == null || questions.Count == 0)
            {
                Console.WriteLine("Нет вопросов для экспорта.");
                return;
            }
            if (!Directory.Exists(ExportPath))
            {
                try { Directory.CreateDirectory(ExportPath); }
                catch (Exception ex)
                {
                    Console.WriteLine($"Не удалось создать директорию экспорта '{ExportPath}': {ex.Message}");
                    return;
                }
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string baseName = $"exported_questions_{timestamp}";
            string jsonPath = Path.Combine(ExportPath, baseName + ".json");
            string xmlPath = Path.Combine(ExportPath, baseName + ".xml"); // XML with original RTF
            string txtPath = Path.Combine(ExportPath, baseName + ".txt");
            string qstPath = Path.Combine(ExportPath, baseName + "_regen.qst"); // Regenerated QST
            string filesDir = Path.Combine(ExportPath, baseName + "_files"); // Directory for OLE files

            try
            {
                // Create directory for binary files BEFORE saving them
                if (questions.Any(q => q.ParsedText.OleObjects.Any() || q.Answers.Any(a => a.ParsedText.OleObjects.Any())))
                {
                    if (!Directory.Exists(filesDir))
                    {
                        Directory.CreateDirectory(filesDir);
                        Console.WriteLine($"Создана директория для OLE файлов: {filesDir}");
                    }
                }

                // Save OLE Objects from all questions/answers
                foreach (var q in questions)
                {
                    foreach (var ole in q.ParsedText.OleObjects)
                    {
                        if (ole.Data != null && ole.SuggestedFileName != null)
                            File.WriteAllBytes(Path.Combine(filesDir, ole.SuggestedFileName), ole.Data);
                    }
                    foreach (var a in q.Answers)
                    {
                        foreach (var ole in a.ParsedText.OleObjects)
                        {
                            if (ole.Data != null && ole.SuggestedFileName != null)
                                File.WriteAllBytes(Path.Combine(filesDir, ole.SuggestedFileName), ole.Data);
                        }
                    }
                }
                Console.WriteLine($"Экспортированы OLE файлы (если были) в: {filesDir}");

                // JSON (uses parsed text with placeholders)
                File.WriteAllText(jsonPath, SerializeToJson(questions), Encoding.UTF8);
                Console.WriteLine($"Экспортировано в JSON: {jsonPath}");

                // XML (uses original RTF for fidelity)
                File.WriteAllText(xmlPath, SerializeToXml(questions, useOriginalRtf: true), Win1251); // Use Win1251 for XML content like original
                Console.WriteLine($"Экспортировано в XML (с RTF): {xmlPath}");

                // TXT (uses parsed text with placeholders)
                File.WriteAllText(txtPath, SerializeToTxt(questions), Encoding.UTF8);
                Console.WriteLine($"Экспортировано в TXT: {txtPath}");

                // Regenerated QST (uses original RTF internally)
                SaveQstFile(qstPath, questions, useOriginalRtf: true);
                Console.WriteLine($"Сгенерирован QST: {qstPath}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка при экспорте данных: {ex.Message}");
            }
        }

        // Improved JSON Serialization (still manual, but better escaping)
        static string SerializeToJson(List<Question> questions)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < questions.Count; i++)
            {
                var q = questions[i];
                sb.AppendLine("  {");
                sb.AppendFormat("    \"Guid\": \"{0}\",\n", q.Guid);
                // Use the ToString() which now includes descriptive placeholders
                sb.AppendFormat("    \"Text\": \"{0}\",\n", JsonEscape(q.ParsedText.ToString()));
                sb.AppendLine("    \"Answers\": [");
                for (int j = 0; j < q.Answers.Count; j++)
                {
                    var a = q.Answers[j];
                    sb.AppendLine("      {");
                    sb.AppendFormat("        \"Guid\": \"{0}\",\n", a.Guid);
                    sb.AppendFormat("        \"IsRight\": {0},\n", a.IsRight.ToString().ToLowerInvariant());
                    // Use the ToString() which now includes descriptive placeholders
                    sb.AppendFormat("        \"Text\": \"{0}\"\n", JsonEscape(a.ParsedText.ToString()));
                    sb.Append("      }");
                    if (j < q.Answers.Count - 1) sb.Append(",");
                    sb.AppendLine();
                }
                sb.AppendLine("    ]");
                sb.Append("  }");
                if (i < questions.Count - 1) sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        // Helper for basic JSON string escaping
        static string JsonEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var sb = new StringBuilder();
            foreach (char c in value)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ')
                        {
                            // Control characters as Unicode escapes
                            sb.AppendFormat("\\u{0:x4}", (int)c);
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }

        // Corrected XML Serialization using XmlWriter and CDATA
        static string SerializeToXml(List<Question> questions, bool useOriginalRtf = true)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                Encoding = Win1251, // Match Examiner's likely internal encoding for the XML *inside* QST
                OmitXmlDeclaration = false, // Keep declaration
                ConformanceLevel = ConformanceLevel.Document // Ensure it's a full document
            };

            // Use StringWriter with specific encoding for the internal XML string
            using (var stringWriter = new StringWriterWithEncoding(Win1251))
            using (var writer = XmlWriter.Create(stringWriter, settings))
            {
                // Add standalone="yes" to match original decompiled code intention
                writer.WriteStartDocument(true);
                writer.WriteStartElement("Questions");

                foreach (var q in questions)
                {
                    writer.WriteStartElement("Question");
                    writer.WriteAttributeString("Guid", q.Guid);

                    writer.WriteStartElement("Text");
                    // Use Original RTF for saving back to QST format by default
                    // Use Parsed Text if specifically requested (e.g., for plain XML export)
                    string questionText = useOriginalRtf ? q.OriginalRtf : q.ParsedText.ToString();
                    writer.WriteCData(questionText ?? string.Empty);
                    writer.WriteEndElement(); // End Text

                    foreach (var a in q.Answers)
                    {
                        writer.WriteStartElement("Answer");
                        writer.WriteAttributeString("Guid", a.Guid);
                        writer.WriteAttributeString("IsRight", a.IsRight.ToString()); // Default C# bool.ToString() is "True"/"False"

                        writer.WriteStartElement("Text");
                        // Use Original RTF for saving back to QST format by default
                        string answerText = useOriginalRtf ? a.OriginalRtf : a.ParsedText.ToString();
                        writer.WriteCData(answerText ?? string.Empty);
                        writer.WriteEndElement(); // End Text

                        writer.WriteEndElement(); // End Answer
                    }
                    writer.WriteEndElement(); // End Question
                }
                writer.WriteEndElement(); // End Questions
                writer.WriteEndDocument();
                writer.Flush();
                return stringWriter.ToString();
            }
        }

        // TXT Serialization using parsed text
        static string SerializeToTxt(List<Question> questions)
        {
            var sb = new StringBuilder();
            int qNum = 1;
            foreach (var q in questions)
            {
                sb.AppendLine($"ВОПРОС {qNum++}:");
                sb.AppendLine(q.ParsedText.ToString()); // Uses GetTextWithPlaceholders implicitly
                sb.AppendLine("ОТВЕТЫ:");
                int aNum = 1;
                foreach (var a in q.Answers)
                    sb.AppendLine($"  {aNum++}) ({(a.IsRight ? "Правильный" : "Неправильный")}): {a.ParsedText.ToString()}");
                sb.AppendLine(new string('-', 40));
            }
            return sb.ToString();
        }

        static void CreateNewFile()
        {
            Console.Write("Введите имя нового файла (без .qst): ");
            string baseFileName = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(baseFileName))
            {
                Console.WriteLine("Имя файла не может быть пустым.");
                return;
            }
            string fileName = Path.ChangeExtension(baseFileName, ".qst");

            var questions = new List<Question>();
            bool addMoreQuestions = true;

            while (addMoreQuestions)
            {
                Console.WriteLine("\n--- Добавление нового вопроса ---");
                var q = new Question { Guid = Guid.NewGuid().ToString() };

                Console.Write("Введите текст вопроса (RTF OLE объекты НЕ поддерживаются при создании): ");
                string questionInput = Console.ReadLine();
                q.OriginalRtf = questionInput;
                q.ParsedText = ParseRtf(q.OriginalRtf, $"newQ{questions.Count}"); // Parse basic RTF if any

                bool addMoreAnswers = true;
                while (addMoreAnswers)
                {
                    Console.WriteLine("  --- Добавление ответа ---");
                    var a = new Answer { Guid = Guid.NewGuid().ToString() };
                    Console.Write("  Введите текст ответа: ");
                    string answerInput = Console.ReadLine();
                    a.OriginalRtf = answerInput;
                    a.ParsedText = ParseRtf(a.OriginalRtf, $"newQ{questions.Count}_A{q.Answers.Count}");

                    Console.Write("  Это правильный ответ? (y/n): ");
                    a.IsRight = Console.ReadLine()?.Trim().ToLowerInvariant() == "y";
                    q.Answers.Add(a);

                    Console.Write("  Добавить еще ответ к этому вопросу? (y/n): ");
                    if (Console.ReadLine()?.Trim().ToLowerInvariant() != "y") addMoreAnswers = false;
                }

                if (q.Answers.Exists(ans => ans.IsRight))
                {
                    questions.Add(q);
                }
                else
                {
                    Console.WriteLine("Предупреждение: Вопрос не добавлен, так как у него нет правильных ответов.");
                }

                Console.Write("\nДобавить еще вопрос в файл? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLowerInvariant() != "y") addMoreQuestions = false;
            }

            if (questions.Count > 0)
            {
                try
                {
                    SaveQstFile(fileName, questions, useOriginalRtf: true);
                    Console.WriteLine($"Файл '{fileName}' успешно сохранен!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка сохранения файла '{fileName}': {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Нет вопросов для сохранения.");
            }
        }

        // Updated SaveQstFile to accept original RTF flag
        static void SaveQstFile(string fileName, List<Question> questions, bool useOriginalRtf)
        {
            string tempXmlPath = Path.GetTempFileName();
            string xmlContent = SerializeToXml(questions, useOriginalRtf);

            try
            {
                // Write XML content with specific encoding first (Win1251)
                File.WriteAllText(tempXmlPath, xmlContent, Win1251);

                // Create ZIP archive
                using (var fs = new FileStream(fileName, FileMode.Create))
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    // Add the XML file to the archive
                    archive.CreateEntryFromFile(tempXmlPath, "Questions.xml", CompressionLevel.Optimal);
                }
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempXmlPath))
                {
                    try { File.Delete(tempXmlPath); } catch { /* Ignore cleanup error */ }
                }
            }
        }

        static void ExportOptions()
        {
            Console.Write("Введите путь к файлу *.qst или директории для экспорта: ");
            string path = Console.ReadLine()?.Trim('"');
            if (string.IsNullOrEmpty(path))
            {
                Console.WriteLine("Путь не указан.");
                return;
            }

            List<Question> questions = new List<Question>();

            try
            {
                if (Directory.Exists(path))
                {
                    Console.WriteLine($"Обработка директории для экспорта: {path}");
                    foreach (var file in Directory.GetFiles(path, "*.qst"))
                    {
                        Console.WriteLine($"  Чтение файла: {Path.GetFileName(file)}");
                        questions.AddRange(ProcessFile(file));
                    }
                }
                else if (File.Exists(path))
                {
                    if (path.EndsWith(".qst", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Чтение файла для экспорта: {Path.GetFileName(path)}");
                        questions = ProcessFile(path);
                    }
                    else
                    {
                        Console.WriteLine("Указанный файл не является *.qst файлом.");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Указанный путь не существует!");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при чтении файлов для экспорта из '{path}': {ex.Message}");
                return;
            }

            if (questions.Count == 0)
            {
                Console.WriteLine("Не найдено вопросов для экспорта.");
                return;
            }

            Console.WriteLine("\nВыберите формат экспорта:");
            Console.WriteLine("1 - JSON (текст с OLE плейсхолдерами)");
            Console.WriteLine("2 - XML (с исходным RTF, как в .qst)");
            Console.WriteLine("3 - TXT (текст с OLE плейсхолдерами)");
            Console.WriteLine("4 - Regenerated QST (пересобранный .qst файл)");
            Console.Write("Ваш выбор: ");
            string format = Console.ReadLine();

            Console.Write("Введите базовое имя выходного файла (без расширения, например, 'my_export'): ");
            string outputBase = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(outputBase))
            {
                Console.WriteLine("Имя выходного файла не может быть пустым.");
                return;
            }

            // Construct paths within the designated ExportPath
            string baseOutputDir = Path.Combine(ExportPath, outputBase + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            string filesDir = baseOutputDir + "_files"; // Files dir related to this export operation
            Directory.CreateDirectory(baseOutputDir); // Create main output dir

            try
            {
                // Create directory for binary files if needed
                if (questions.Any(q => q.ParsedText.OleObjects.Any() || q.Answers.Any(a => a.ParsedText.OleObjects.Any())))
                {
                    Directory.CreateDirectory(filesDir);
                    Console.WriteLine($"Создана директория для OLE файлов: {filesDir}");
                }

                // Save OLE Objects first
                foreach (var q in questions)
                {
                    foreach (var ole in q.ParsedText.OleObjects)
                    {
                        if (ole.Data != null && ole.SuggestedFileName != null)
                            File.WriteAllBytes(Path.Combine(filesDir, ole.SuggestedFileName), ole.Data);
                    }
                    foreach (var a in q.Answers)
                    {
                        foreach (var ole in a.ParsedText.OleObjects)
                        {
                            if (ole.Data != null && ole.SuggestedFileName != null)
                                File.WriteAllBytes(Path.Combine(filesDir, ole.SuggestedFileName), ole.Data);
                        }
                    }
                }
                if (Directory.Exists(filesDir) && Directory.GetFiles(filesDir).Length > 0)
                {
                    Console.WriteLine($"Экспортированы OLE файлы в: {filesDir}");
                }

                string outputPath;
                switch (format)
                {
                    case "1": // JSON
                        outputPath = Path.Combine(baseOutputDir, outputBase + ".json");
                        File.WriteAllText(outputPath, SerializeToJson(questions), Encoding.UTF8);
                        Console.WriteLine($"JSON экспортирован в: {outputPath}");
                        break;
                    case "2": // XML with RTF
                        outputPath = Path.Combine(baseOutputDir, outputBase + ".xml");
                        File.WriteAllText(outputPath, SerializeToXml(questions, useOriginalRtf: true), Win1251);
                        Console.WriteLine($"XML (с RTF) экспортирован в: {outputPath}");
                        break;
                    case "3": // TXT
                        outputPath = Path.Combine(baseOutputDir, outputBase + ".txt");
                        File.WriteAllText(outputPath, SerializeToTxt(questions), Encoding.UTF8);
                        Console.WriteLine($"TXT экспортирован в: {outputPath}");
                        break;
                    case "4": // Regenerated QST
                        outputPath = Path.Combine(baseOutputDir, outputBase + ".qst");
                        SaveQstFile(outputPath, questions, useOriginalRtf: true);
                        Console.WriteLine($"QST файл пересобран в: {outputPath}");
                        break;
                    default:
                        Console.WriteLine("Неверный формат!");
                        // Clean up empty base directory if nothing was exported
                        if (!Directory.EnumerateFileSystemEntries(baseOutputDir).Any())
                            Directory.Delete(baseOutputDir);
                        if (Directory.Exists(filesDir) && !Directory.EnumerateFileSystemEntries(filesDir).Any())
                            Directory.Delete(filesDir);
                        return;
                }
                Console.WriteLine("Экспорт завершен!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка во время экспорта в формат '{format}': {ex.Message}");
                // Attempt cleanup on error
                try
                {
                    if (Directory.Exists(filesDir)) Directory.Delete(filesDir, true);
                    if (Directory.Exists(baseOutputDir)) Directory.Delete(baseOutputDir, true);
                }
                catch { }
            }
        }
    }

    // Helper class for StringWriter with specific encoding
    public sealed class StringWriterWithEncoding : StringWriter
    {
        private readonly Encoding encoding;

        public StringWriterWithEncoding(Encoding encoding)
        {
            this.encoding = encoding ?? throw new ArgumentNullException(nameof(encoding));
        }

        public override Encoding Encoding => encoding;
    }

    // Internal state for the RTF parser group stack
    internal class RtfGroupState
    {
        public bool IsRoot { get; set; } = false; // Identify the root group
        public StringBuilder PendingText { get; set; } = new StringBuilder();
        public StringBuilder InternalText { get; set; } = new StringBuilder(); // Track text within this specific group for objclass etc.

        // OLE Object state
        public bool IsObject { get; set; } = false;
        public string ObjectClass { get; set; } = "Unknown";
        public StringBuilder ObjectDataHex { get; set; } = new StringBuilder();
        public bool IsReadingObjectData { get; set; } = false;

        // Picture state
        public bool IsPicture { get; set; } = false;
        public StringBuilder PictureDataHex { get; set; } = new StringBuilder();
        public bool IsReadingPictureData { get; set; } = false; // Often true immediately inside \pict
        public OleFileType DetectedPictureType { get; set; } = OleFileType.Unknown; // Set by \wmetafile, etc.

        public RtfConsumeTarget ConsumeNextTextAs { get; set; } = RtfConsumeTarget.None;
        public bool NextGroupIsIgnorable { get; set; } = false; // Flag for \*

        // Unicode state
        public int AsciiFallbackChars { get; set; } = 1; // For \uN

        // Ignore state (for \* groups or \result)
        public bool IgnoreContent { get; set; } = false;
        public int IgnoreLevel { get; set; } = 0; // Track nesting level of ignored groups

        // Default constructor for root
        public RtfGroupState() { }

        // Constructor for nested groups, inheriting state
        public RtfGroupState(RtfGroupState parent)
        {
            // Inherit flags that persist across nested groups unless explicitly changed
            IsObject = parent.IsObject;
            IsPicture = parent.IsPicture;
            ObjectClass = parent.ObjectClass;
            AsciiFallbackChars = parent.AsciiFallbackChars;
            IgnoreContent = parent.IgnoreContent;
            IgnoreLevel = parent.IgnoreLevel;
            IsReadingObjectData = parent.IsReadingObjectData; // Inherit reading state? Maybe reset? Let's inherit for now.
            IsReadingPictureData = parent.IsReadingPictureData; // Inherit reading state? Maybe reset? Let's inherit.
            DetectedPictureType = parent.DetectedPictureType; // Inherit detected type

            // Reset specific flags/buffers for the new group
            PendingText = new StringBuilder();
            InternalText = new StringBuilder(); // Reset internal text for new group
            ObjectDataHex = new StringBuilder(); // Reset data buffers for new group
            PictureDataHex = new StringBuilder();
            ConsumeNextTextAs = RtfConsumeTarget.None; // Reset consumption target
            NextGroupIsIgnorable = false; // Reset ignorable flag trigger
        }
    }
    internal enum RtfConsumeTarget
    {
        None,
        ObjectClass,
        // Add other targets if needed (e.g., Font name in fonttbl)
    }

    // Updated Question class to hold both original RTF and parsed result
    public class Question
    {
        public string Guid { get; set; }
        public string OriginalRtf { get; set; }
        public RtfParseResult ParsedText { get; set; }
        public List<Answer> Answers { get; set; } = new List<Answer>();
    }

    // Updated Answer class similar to Question
    public class Answer
    {
        public string Guid { get; set; }
        public bool IsRight { get; set; }
        public string OriginalRtf { get; set; }
        public RtfParseResult ParsedText { get; set; }
    }
}