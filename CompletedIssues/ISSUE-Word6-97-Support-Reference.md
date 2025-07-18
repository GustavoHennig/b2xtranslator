

This parser has good support for Word 97 files but currently ignores the Word 95 format, even though it is less complex than Word 97.
Please add support for Word 95 files using the code below as a reference. Try to keep the existing class structure intact and avoid excessive refactoring.

It must work with the file `samples/a-idade-media.word95.doc`, it has nFib = 104, and the document is valid.
To test, run from solution path:
```bat
dotnet run --project .\Shell\doc2text\doc2text.csproj -- "samples/a-idade-media.word95.doc"
```

Below is a code example of another project I am working on that fully interprets Word95, and Word6 files, please use this code as a reference.

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using WvWareNet.Parsers;
using WvWareNet.Utilities;

namespace WvWareNet.Parsers
{
    public class WordDocumentParser
    {
        private readonly CompoundFileBinaryFormatParser _cfbfParser;
        private readonly ILogger _logger;

        public WordDocumentParser(CompoundFileBinaryFormatParser cfbfParser, ILogger logger)
        {
            _cfbfParser = cfbfParser;
            _logger = logger;
            _documentModel = new Core.DocumentModel();
        }

        private Core.DocumentModel _documentModel;

        private static bool IsAlphaNumeric(byte b)
        {
            char c = (char)b;
            return char.IsLetterOrDigit(c);
        }

        public void ParseDocument(string? password = null)
        {
            // Ensure header is parsed before reading directory entries
            _cfbfParser.ParseHeader();

            // Parse directory entries
            var entries = _cfbfParser.ParseDirectoryEntries();

            // Locate WordDocument stream
            var wordDocEntry = entries.Find(e =>
                e.Name.Contains("WordDocument", StringComparison.OrdinalIgnoreCase));

            if (wordDocEntry == null)
            {
                _logger.LogError("Required WordDocument stream not found in CFBF file. Available directory entries:");
                foreach (var entry in entries)
                {
                    _logger.LogInfo($"- {entry.Name} (Type: {entry.EntryType}, Size: {entry.StreamSize})");
                }
                throw new InvalidDataException("Required WordDocument stream not found in CFBF file.");
            }
            _logger.LogInfo($"[DEBUG] WordDocument Entry: Name='{wordDocEntry.Name}', StartingSectorLocation={wordDocEntry.StartingSectorLocation}, StreamSize={wordDocEntry.StreamSize}");

            // Read stream data
            var wordDocStream = _cfbfParser.ReadStream(wordDocEntry);

            // Log the first 32 bytes of the WordDocument stream for debugging
            _logger.LogInfo($"[DEBUG] WordDocument stream length: {wordDocStream.Length}");
            if (wordDocStream.Length >= 32)
            {
                var headerBytes = new byte[32];
                Array.Copy(wordDocStream, 0, headerBytes, 0, 32);
                _logger.LogInfo($"[DEBUG] WordDocument stream header: {BitConverter.ToString(headerBytes).Replace("-", " ")}");
            }


            // Parse FIB early to detect Word95 before Table stream check
            var fib = WvWareNet.Core.FileInformationBlock.Parse(wordDocStream);
            _logger.LogInfo($"[DEBUG] Parsed FIB. nFib: {fib.NFib}, fComplex: {fib.FComplex}, fEncrypted: {fib.FEncrypted}");


            // Decrypt Word95 documents if necessary
            if (fib.NFib == 0x0065 && fib.FEncrypted)
            {
                if (string.IsNullOrEmpty(password))
                    throw new InvalidDataException("Password required for Word95 decryption.");

                wordDocStream = WvWareNet.Core.Decryptor95.Decrypt(wordDocStream, password, fib.LKey);
                fib = WvWareNet.Core.FileInformationBlock.Parse(wordDocStream);
            }

            var streamName = fib.FWhichTblStm
                ? "1Table"
                : "0Table";

            var tableEntry = entries
                .FirstOrDefault(e =>
                    e.Name.Equals(streamName, StringComparison.OrdinalIgnoreCase));


            // Word95 files: 100=Word6, 101=Word95, 104=Word97 but some Word95 files use 104
            bool isWord95 = fib.NFib == 100 || fib.NFib == 101 || fib.NFib == 104;

            if (tableEntry == null)
            {
                if (isWord95)
                {
                    // For Word95 files, try to proceed without Table stream if not found
                    _logger.LogWarning($"Table stream not found in Word95/Word6 document (NFib={fib.NFib}), attempting to parse with reduced functionality");
                }
                else if (fib.NFib == 53200) // Special case for Word95 test file
                {
                    _logger.LogWarning($"Table stream not found in Word95 document (NFib={fib.NFib}), attempting to parse with reduced functionality");
                }
                else
                {
                    _logger.LogError($"Table stream not found, and not a recognized Word95/Word6 document (NFib={fib.NFib})");
                    throw new InvalidDataException("Required Table stream not found in CFBF file.");
                }
            }

            _logger.LogInfo($"[DEBUG] Found WordDocument stream: {wordDocEntry.Name}");
            byte[]? tableStream = null;

            if (tableEntry != null)
            {
                _logger.LogInfo($"[DEBUG] Found Table stream: {tableEntry.Name}");
                tableStream = _cfbfParser.ReadStream(tableEntry);
            }

            _documentModel.FileInfo = fib;

            if ((fib.FEncrypted || fib.FCrypto)) // && !isWord95)
                throw new NotSupportedException("Encrypted Word documents are not supported.");

            if (fib.FibVersion == null)
                _logger.LogWarning($"Unknown Word version NFib={fib.NFib}");
            else
                _logger.LogInfo($"Detected Word version: {fib.FibVersion}");

            var pieceTable = WvWareNet.Core.PieceTable.CreateFromStreams(_logger, fib, tableStream, wordDocStream);

            // Extract CHPX (character formatting) data from Table stream using PLCFCHPX
            // PLCFCHPX location is in FIB: FcPlcfbteChpx, LcbPlcfbteChpx
            var chpxList = new List<byte[]>();
            if (tableStream != null && fib.FcPlcfbteChpx > 0 && fib.LcbPlcfbteChpx > 0 && fib.FcPlcfbteChpx + fib.LcbPlcfbteChpx <= tableStream.Length)
            {
                // WORD95 do no use this block
            }
            else
            {
                // Fallback: assign null CHPX to all pieces
                for (int i = 0; i < pieceTable.Pieces.Count; i++)
                    chpxList.Add(null);
            }
            pieceTable.AssignChpxToPieces(chpxList);

            // Parse stylesheet
            WvWareNet.Core.Stylesheet stylesheet = new WvWareNet.Core.Stylesheet();
            _logger.LogInfo($"[DEBUG] Stylesheet info: FcStshf={fib.FcStshf}, LcbStshf={fib.LcbStshf}, tableStream.Length={tableStream?.Length ?? 0}");

            if (tableStream != null && fib.FcStshf > 0 && fib.LcbStshf > 0 && fib.FcStshf + fib.LcbStshf <= tableStream.Length)
            {
                // WORD95 do no use this block
            }
            else
            {
                _logger.LogWarning($"No stylesheet found via FIB - FcStshf={fib.FcStshf}, LcbStshf={fib.LcbStshf}");

                // Create basic default stylesheet with standard built-in styles
                stylesheet = new WvWareNet.Core.Stylesheet();
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 0, Name = "Normal" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 1, Name = "heading 1" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 2, Name = "heading 2" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 3, Name = "heading 3" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 4, Name = "heading 4" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 5, Name = "heading 5" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 6, Name = "heading 6" });
                _logger.LogInfo("[DEBUG] Using default built-in styles as fallback");
            }

            // Extract text using PieceTable and populate DocumentModel with sections and paragraphs
            _documentModel = new WvWareNet.Core.DocumentModel();
            _documentModel.FileInfo = fib; // Store FIB in the document model
            _documentModel.Stylesheet = stylesheet; // Store stylesheet in the document model

            // For now, create a single default section.
            // In a more complete implementation, sections would be parsed from the document structure.
            var defaultSection = new WvWareNet.Core.Section();
            _documentModel.Sections.Add(defaultSection);

            using var wordDocMs = new System.IO.MemoryStream(wordDocStream);

            // Log character counts for debugging
            _logger.LogInfo($"[DEBUG] Character counts - Text: {fib.CcpText}, Footnotes: {fib.CcpFtn}, Headers: {fib.CcpHdr}");

            // --- Paragraph boundary detection using [MS-DOC] 2.4.2 algorithm ---
            // Use new FindParagraphStartCp and FindParagraphEndCp methods
            if (tableStream != null && fib.FcPlcfbtePapx > 0 && fib.LcbPlcfbtePapx > 0 && fib.FcPlcfbtePapx + fib.LcbPlcfbtePapx <= tableStream.Length)
            {
                // WORD95 do no use this block
            }
            else
            {
                // Fallback: old logic using newlines
                var currentParagraph = new WvWareNet.Core.Paragraph();
                defaultSection.Paragraphs.Add(currentParagraph);

                for (int i = 0; i < pieceTable.Pieces.Count; i++)
                {
                    string text = pieceTable.GetTextForPiece(i, wordDocMs);
                    _logger.LogInfo($"[DEBUG] Extracted text from piece {i}: \"{text.Replace("\r", "\\r").Replace("\n", "\\n")}\"");

                    var chpx = pieceTable.Pieces[i].Chpx;
                    var charProps = new WvWareNet.Core.CharacterProperties();
                    if (chpx != null && chpx.Length > 0)
                    {
                        for (int b = 0; b < chpx.Length - 2; b++)
                        {
                            byte sprm = chpx[b];
                            byte val = chpx[b + 2];
                            switch (sprm)
                            {
                                case 0x08: charProps.IsBold = val != 0; break;
                                case 0x09: charProps.IsItalic = val != 0; break;
                                case 0x0A: charProps.IsStrikeThrough = val != 0; break;
                                case 0x0D: charProps.IsSmallCaps = val != 0; break;
                                case 0x0E: charProps.IsAllCaps = val != 0; break;
                                case 0x0F: charProps.IsHidden = val != 0; break;
                                case 0x18: charProps.IsUnderlined = val != 0; break;
                                case 0x2A: charProps.FontSize = val; break;
                                case 0x2B: break;
                                case 0x2D: charProps.IsUnderlined = val != 0; break;
                                case 0x2F: charProps.FtcaAscii = val; break;
                                case 0x30: charProps.LanguageId = val; break;
                                case 0x31: charProps.CharacterSpacing = (short)val; break;
                                case 0x32: charProps.CharacterScaling = val; break;
                            }
                        }
                    }

                    var run = new WvWareNet.Core.Run { Text = text, Properties = charProps };
                    currentParagraph.Runs.Add(run);

                    if (text.EndsWith("\r\n") || text.EndsWith("\n") || text.EndsWith("\r"))
                    {
                        currentParagraph = new WvWareNet.Core.Paragraph();
                        defaultSection.Paragraphs.Add(currentParagraph);
                    }
                }

                if (defaultSection.Paragraphs.Count > 0 && defaultSection.Paragraphs[^1].Runs.Count == 0)
                {
                    defaultSection.Paragraphs.RemoveAt(defaultSection.Paragraphs.Count - 1);
                }
            }

            // --- Header/Footer/Footnote Extraction ---
            // Parse PLCF for headers/footers if available
            if (tableStream != null && fib.FcPlcfhdd > 0 && fib.LcbPlcfhdd > 0 && fib.FcPlcfhdd + fib.LcbPlcfhdd <= tableStream.Length)
            {
                // WORD95 do no use this block
            }
            else
            {
                _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            }

       

            // Parse PLCF for footnotes if available
            if (tableStream != null && fib.FcPlcffldFtn > 0 && fib.LcbPlcffldFtn > 0 && fib.FcPlcffldFtn + fib.LcbPlcffldFtn <= tableStream.Length)
            {
                // WORD95 do no use this block
            }
            else
            {
                _documentModel.Footnotes.Add(new WvWareNet.Core.Footnote { ReferenceId = 1 });
            }

            // Add a default footer if not already added
            if (_documentModel.Footers.Count == 0)
            {
                _documentModel.Footers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            }

            // Log document structure for debugging
            _logger.LogInfo($"[STRUCTURE] Document has {_documentModel.Sections.Count} section(s)");
            for (int i = 0; i < _documentModel.Sections.Count; i++)
            {
                var section = _documentModel.Sections[i];
                _logger.LogInfo($"[STRUCTURE] Section {i + 1}: {section.Paragraphs.Count} paragraph(s)");
                for (int j = 0; j < section.Paragraphs.Count; j++)
                {
                    var paragraph = section.Paragraphs[j];
                    _logger.LogInfo($"[STRUCTURE] Paragraph {j + 1}: {paragraph.Runs.Count} run(s)");
                    for (int k = 0; k < paragraph.Runs.Count; k++)
                    {
                        var run = paragraph.Runs[k];
                        _logger.LogInfo($"[STRUCTURE] Run {k + 1}: Length={run.Text?.Length ?? 0}, TextPreview='{(run.Text != null ? run.Text.Substring(0, Math.Min(20, run.Text.Length)) : "null")}'");
                    }
                }
            }

            _logger.LogInfo("[PARSING] Document parsing completed");
        }


        public string ExtractText()
        {
            if (_documentModel == null)
                throw new InvalidOperationException("Document not parsed. Call ParseDocument() first.");

            var textBuilder = new System.Text.StringBuilder();
            int runCount = 0;
            int charCount = 0;

            // Extract text from document body only (ignore headers/footers/notes)
            foreach (var section in _documentModel.Sections)
            {
                foreach (var paragraph in section.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        if (run.Text != null)
                        {
                            textBuilder.Append(run.Text);
                            runCount++;
                            charCount += run.Text.Length;
                        }
                    }
                }
            }

            _logger.LogInfo($"[EXTRACTION] Extracted {runCount} runs with {charCount} characters");
            return textBuilder.ToString();
        }
    }
}


using System;
using System.Collections.Generic;
using System.IO;
using WvWareNet.Utilities;

namespace WvWareNet.Core;

public class PieceTable
{
    private readonly ILogger _logger;
    private readonly List<PieceDescriptor> _pieces = new();

    public IReadOnlyList<PieceDescriptor> Pieces => _pieces;

    public PieceTable(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Factory method to create and parse a PieceTable from the given streams and FIB.
    /// Handles CLX extraction and all fallbacks.
    /// </summary>
    public static PieceTable CreateFromStreams(
        ILogger logger,
        FileInformationBlock fib,
        byte[]? tableStream,
        byte[] wordDocStream)
    {
        var pieceTable = new PieceTable(logger);

        byte[]? clxData = null;
        if (tableStream != null)
        {
            // WORD95 do no use this block
        }

        if (clxData != null && clxData.Length > 0)
        {
            // WORD95 do no use this block
        }
        else
        {
            logger.LogWarning("No CLX data found or invalid, treating document as single piece");
            pieceTable.SetSinglePiece(fib);
            logger.LogInfo($"[DEBUG] Fallback single piece created. IsUnicode: {pieceTable.Pieces[0].IsUnicode}");
        }

        return pieceTable;
    }

    /// <summary>
    /// Assigns CHPX (character formatting) data to each piece descriptor.
    /// </summary>
    public void AssignChpxToPieces(List<byte[]> chpxList)
    {
        // Assumes chpxList.Count == _pieces.Count or chpxList.Count == _pieces.Count - 1
        int count = Math.Min(_pieces.Count, chpxList.Count);
        for (int i = 0; i < count; i++)
        {
            _pieces[i].Chpx = chpxList[i];
        }
    }


    public string GetTextForPiece(int index, Stream documentStream, int? codePage = null)
    {
        if (index < 0 || index >= _pieces.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        var piece = _pieces[index];
        return GetTextForRange(piece.FcStart, piece.FcEnd, documentStream, piece.IsUnicode, codePage);
    }

    /// <summary>
    /// Retrieve text for an arbitrary file position range. The range may span
    /// multiple pieces and does not need to align to piece boundaries.
    /// </summary>
    public string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, int? codePage = null)
    {
        if (fcEnd <= fcStart)
            return string.Empty;

        var sb = new System.Text.StringBuilder();

        foreach (var piece in _pieces)
        {
            int start = Math.Max(fcStart, piece.FcStart);
            int end = Math.Min(fcEnd, piece.FcEnd);
            if (start >= end)
                continue;

            sb.Append(GetTextForRange(start, end, documentStream, piece.IsUnicode, codePage));
        }

        return sb.ToString();
    }

    private string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, bool isUnicode, int? codePage = null)
    {
        _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart={fcStart}, fcEnd={fcEnd}, isUnicode={isUnicode}, streamLength={documentStream.Length}, codePage={codePage}");
        int length = fcEnd - fcStart;
        if (length <= 0)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: Invalid length {length}, returning empty string");
            return string.Empty;
        }

        if (fcStart >= documentStream.Length)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart {fcStart} >= streamLength {documentStream.Length}, returning empty string");
            return string.Empty;
        }

        // Clamp to stream bounds
        if (fcEnd > documentStream.Length)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: Clamping fcEnd from {fcEnd} to {documentStream.Length}");
            fcEnd = (int)documentStream.Length;
            length = fcEnd - fcStart;
        }

        using var reader = new BinaryReader(documentStream, System.Text.Encoding.UTF8, leaveOpen: true);
        documentStream.Seek(fcStart, SeekOrigin.Begin);
        byte[] bytes = reader.ReadBytes(length);
        
        _logger.LogInfo($"[DEBUG] GetTextForRange: Read {bytes.Length} bytes from offset {fcStart}");
        if (bytes.Length > 0)
        {
            string hexPreview = BitConverter.ToString(bytes, 0, Math.Min(32, bytes.Length));
            _logger.LogInfo($"[DEBUG] GetTextForRange: First bytes: {hexPreview}");
        }

        string text;
        if (isUnicode)
        {
            text = System.Text.Encoding.Unicode.GetString(bytes);
            _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as UTF-16LE");
        }
        else
        {
            // Use code page if provided, otherwise default to Windows-1252
            var encoding = codePage.HasValue ? System.Text.Encoding.GetEncoding(codePage.Value) : System.Text.Encoding.GetEncoding(1252);
            text = encoding.GetString(bytes);
            _logger.LogInfo($"[DEBUG] GetTextForRange: Decoded as {(codePage.HasValue ? $"code page {codePage.Value}" : "Windows-1252")}");
        }

        string processedText = ProcessFieldCodes(text);
        _logger.LogInfo($"[DEBUG] GetTextForRange raw text length: {text.Length}, processed length: {processedText.Length}");
        _logger.LogInfo($"[DEBUG] GetTextForRange returning: '{processedText.Substring(0, Math.Min(50, processedText.Length))}'");
        return CleanText(processedText);
    }

    /// <summary>
    /// Process Word field codes and extract only the field results.
    /// Field structure: [0x13][field code][0x14][field result][0x15]
    /// We want to keep only the field result part.
    /// For HYPERLINK fields without field result, extract the URL from the field code.
    /// </summary>
    private static string ProcessFieldCodes(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        int i = 0;
        
        while (i < input.Length)
        {
            char c = input[i];
            
            if (c == (char)0x13) // Field begin (0x13)
            {
                // Find the field separator (0x14) and field end (0x15)
                int separatorPos = -1;
                int endPos = -1;
                int depth = 1; // Track nested fields
                
                for (int j = i + 1; j < input.Length; j++)
                {
                    if (input[j] == (char)0x13) // Nested field begin
                    {
                        depth++;
                    }
                    else if (input[j] == (char)0x15) // Field end
                    {
                        depth--;
                        if (depth == 0)
                        {
                            endPos = j;
                            break;
                        }
                    }
                    else if (input[j] == (char)0x14 && depth == 1 && separatorPos == -1) // Field separator at our level
                    {
                        separatorPos = j;
                    }
                }
                
                if (endPos != -1)
                {
                    if (separatorPos != -1)
                    {
                        // Extract field result (between separator and end)
                        string fieldResult = input.Substring(separatorPos + 1, endPos - separatorPos - 1);
                        sb.Append(fieldResult);
                    }
                    else
                    {
                        // No field separator found, extract field code and check if it's a HYPERLINK
                        string fieldCode = input.Substring(i + 1, endPos - i - 1);
                        
                        // Handle HYPERLINK fields
                        if (fieldCode.StartsWith(" HYPERLINK ", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract URL from HYPERLINK field
                            string url = fieldCode.Substring(11).Trim(); // Remove " HYPERLINK "
                            
                            // Remove any additional parameters or quotes
                            int spaceIndex = url.IndexOf(' ');
                            if (spaceIndex > 0)
                            {
                                url = url.Substring(0, spaceIndex);
                            }
                            
                            // Remove quotes if present
                            url = url.Trim('"');
                            
                            sb.Append(url);
                        }
                        else
                        {
                            // For other field types without separator, skip the field
                        }
                    }
                    // Skip to after the field end
                    i = endPos + 1;
                }
                else
                {
                    // Malformed field, just append the character and continue
                    sb.Append(c);
                    i++;
                }
            }
            else if (c == (char)0x14 || c == (char)0x15) // Standalone field separators/ends (shouldn't happen in well-formed text)
            {
                // Skip these control characters
                i++;
            }
            else
            {
                // Regular character, append it
                sb.Append(c);
                i++;
            }
        }
        
        return sb.ToString();
    }

    private static string CleanText(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        foreach (char c in input)
        {
            // Handle special Word control characters
            if (c == (char)0x07) // Word tab character (0x07) -> convert to standard tab
            {
                sb.Append('\t');
            }
            else if (c == (char)0x0B) // Word line break (0x0B) -> convert to newline
            {
                sb.Append('\n');
            }
            // Allow a much wider range of characters - be more permissive
            else if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
                char.IsSymbol(c) ||  
                // TODO: This seems redundany, review it
                c == '\r' || c == '\n' || c == '\t' || c == '\v' || c == '~' ||
                (c >= 'А' && c <= 'я') || // Cyrillic range
                (c >= 'À' && c <= 'ÿ') || // Latin-1 Supplement
                c == '=' || c == '<' || c == '>' || c == '^' || c == '|' || c == '+' || c == '-') // Explicitly allow missing chars
            {
                sb.Append(c);
            }
            else if (c == '\0' || c < ' ') // Replace null chars and control chars
            {
                // Skip these characters
            }
            else
            {
                // For debugging: log any character that's being filtered out
                Console.WriteLine($"[DEBUG] CleanText filtering out character: '{c}' (U+{(int)c:X4})");
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Attempt to determine if the pieces in a Word95 document contain
    /// 16-bit text. This mirrors the heuristic used by wvGuess16bit in the
    /// original wvWare project.
    /// </summary>
    /// <param name="fcValues">File positions for each piece.</param>
    /// <param name="cpArray">Character position array from the piece table.</param>
    private static bool GuessPiecesAre16Bit(uint[] fcValues, int[] cpArray)
    {
        var tuples = new List<(uint Fc, uint Offset)>();
        for (int i = 0; i < fcValues.Length; i++)
        {
            uint offset = (uint)(cpArray[i + 1] - cpArray[i]) * 2u;
            tuples.Add((fcValues[i], offset));
        }

        tuples.Sort((a, b) => a.Fc.CompareTo(b.Fc));

        for (int i = 0; i < tuples.Count - 1; i++)
        {
            if (tuples[i].Fc + tuples[i].Offset > tuples[i + 1].Fc)
                return false; // Overlap means 8-bit text
        }

        return true; // No overlap detected -> assume 16-bit
    }

    /// <summary>
    /// Replace the current table with a single Unicode piece using the
    /// supplied file positions. Used as a fallback when the piece table is
    /// corrupt or not present.
    /// </summary>
    public void SetSinglePiece(FileInformationBlock fib)
    {
        _pieces.Clear();
        // Always use 8-bit encoding for fallback, as in the working version.
        bool isUnicode = false;
        _pieces.Add(new PieceDescriptor
        {
            FilePosition = fib.FcMin,
            IsUnicode = isUnicode,
            HasFormatting = false,
            CpStart = 0,
            CpEnd = (int)(fib.FcMac - fib.FcMin),
            FcStart = (int)fib.FcMin,
            FcEnd = (int)fib.FcMac
        });
    }
}


namespace WvWareNet.Core;

public class FileInformationBlock
{
    /* properties suppressed for simplicity */

    public static FileInformationBlock Parse(byte[] wordDocumentStream)
    {
        var fib = new FileInformationBlock();
        if (wordDocumentStream.Length < 512) // A minimal FIB requires at least the header size
        {
            return fib; // Return a default FIB if stream is too small
        }

        using var ms = new System.IO.MemoryStream(wordDocumentStream);
        using var reader = new System.IO.BinaryReader(ms);

        // Read basic FIB fields from the beginning
        fib.WIdent = reader.ReadUInt16();
        fib.NFib = reader.ReadUInt16();
        fib.NProduct = reader.ReadUInt16();
        fib.Lid = reader.ReadUInt16();
        fib.PnNext = reader.ReadInt16();

        // Parse flag bits
        ushort flags = reader.ReadUInt16();
        fib.FDot = (flags & 0x0001) != 0;
        fib.FGlsy = (flags & 0x0002) != 0;
        fib.FComplex = (flags & 0x0004) != 0;
        fib.FHasPic = (flags & 0x0008) != 0;
        fib.CQuickSaves = (byte)((flags >> 4) & 0x000F);
        fib.FEncrypted = (flags & 0x0100) != 0;
        fib.FWhichTblStm = (flags & 0x0200) != 0;
        fib.FReadOnlyRecommended = (flags & 0x0400) != 0;
        fib.FWriteReservation = (flags & 0x0800) != 0;
        fib.FExtChar = (flags & 0x1000) != 0;
        fib.FLoadOverride = (flags & 0x2000) != 0;
        fib.FFarEast = (flags & 0x4000) != 0;
        fib.FCrypto = (flags & 0x8000) != 0;

        fib.NFibBack = reader.ReadUInt16();
        fib.LKey = reader.ReadUInt32();

        fib.Envr = reader.ReadByte();
        byte envrFlags = reader.ReadByte();
        fib.FMac = (envrFlags & 0x01) != 0;
        fib.FEmptySpecial = (envrFlags & 0x02) != 0;
        fib.FLoadOverridePage = (envrFlags & 0x04) != 0;
        fib.FFutureSavedUndo = (envrFlags & 0x08) != 0;
        fib.FWord97Saved = (envrFlags & 0x10) != 0;
        fib.FSpare0 = (byte)(envrFlags >> 5);

        fib.Chse = reader.ReadUInt16();
        fib.ChsTables = reader.ReadUInt16();

        // Read Cfclcb to determine the size of the FC/LCB array
        ms.Position = 0x005C;
        if (ms.Length > 0x005C + 2)
            fib.Cfclcb = reader.ReadUInt16();

        // Read FcMin and FcMac at standard offset
        ms.Position = 0x18;
        fib.FcMin = reader.ReadUInt32();
        fib.FcMac = reader.ReadUInt32();

        // Determine version-specific offsets
        int offsetPlcfbteChpx = 0x00FA; // Default for Word 6/95
        int offsetPlcfbtePapx = 0x0102;
        int offsetClx = 0x00A4;
        int offsetPlcfhdd = 0x00F2;
        int offsetPlcfftn = 0x012A;
        int offsetPlcftxbxTxt = 0x01F6;
        int offsetStshf = 0x00A0; // Stylesheet offset for Word 6/95

        if (fib.NFib >= 104) // Word 97 and later
        {
            offsetPlcfbteChpx = 0x014E;
            offsetPlcfbtePapx = 0x0156;
            offsetClx = 0x01A2;
            offsetPlcfhdd = 0x0142;
            offsetPlcfftn = 0x015E;
            offsetPlcftxbxTxt = 0x01C6;
            offsetStshf = 0x00A2; // Corrected stylesheet offset for Word 97+
        }

        // Read version-specific properties
        ms.Position = offsetPlcftxbxTxt;
        fib.FcPlcftxbxTxt = reader.ReadInt32();
        ms.Position = offsetPlcftxbxTxt + 4;
        fib.LcbPlcftxbxTxt = reader.ReadUInt32();

        ms.Position = offsetPlcfbteChpx;
        fib.FcPlcfbteChpx = reader.ReadInt32();
        fib.LcbPlcfbteChpx = reader.ReadUInt32();

        ms.Position = offsetPlcfbtePapx;
        fib.FcPlcfbtePapx = reader.ReadInt32();
        fib.LcbPlcfbtePapx = reader.ReadUInt32();

        // Read CLX information
            ms.Position = offsetClx;
            fib.FcClx = reader.ReadInt32();
            fib.LcbClx = reader.ReadUInt32();

        // Debugging: Output CLX values for fast saved documents
        if (fib.FDot || fib.CQuickSaves > 1)
        {
            System.Console.WriteLine($"[DEBUG] Fast Save detected - CLX Offset: {fib.FcClx}, Length: {fib.LcbClx}");
        }

        ms.Position = offsetPlcfhdd;
        fib.FcPlcfhdd = reader.ReadInt32();
        fib.LcbPlcfhdd = reader.ReadUInt32();

        ms.Position = offsetPlcfftn;
        fib.FcPlcffldFtn = reader.ReadInt32();
        fib.LcbPlcffldFtn = reader.ReadUInt32();

        // Read stylesheet information
        ms.Position = offsetStshf;
        fib.FcStshf = reader.ReadInt32();
        fib.LcbStshf = reader.ReadUInt32();

        // Read character counts (Word 97+)
        if (fib.NFib >= 104) // Word 97 and later
        {
            ms.Position = 0x00A4; // Character count section
            fib.CcpText = reader.ReadUInt32();
            fib.CcpFtn = reader.ReadInt32();
            fib.CcpHdr = reader.ReadInt32();
            fib.CcpMcr = reader.ReadInt32();
            fib.CcpAtn = reader.ReadInt32();
            fib.CcpEdn = reader.ReadInt32();
            fib.CcpTxbx = reader.ReadInt32();
            fib.CcpHdrTxbx = reader.ReadInt32();
        }
        else
        {
            // For older versions, try to determine text length from FcMin/FcMac
            fib.CcpText = fib.FcMac - fib.FcMin;
            fib.CcpFtn = 0;
            fib.CcpHdr = 0;
            fib.CcpMcr = 0;
            fib.CcpAtn = 0;
            fib.CcpEdn = 0;
            fib.CcpTxbx = 0;
            fib.CcpHdrTxbx = 0;
        }

        return fib;
    }
}



```
