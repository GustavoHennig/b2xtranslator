using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;
using System.Collections.Generic;

namespace b2xtranslator.txt.TextMapping
{
    public class MainDocumentMapping : DocumentMapping
    {
        public MainDocumentMapping(ConversionContext ctx, IWriter writer)
            : base(ctx, writer)
        {
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            //start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "document", OpenXmlNamespaces.WordprocessingML);

            //write namespaces
            _writer.WriteAttributeString("xmlns", "v", null, OpenXmlNamespaces.VectorML);
            _writer.WriteAttributeString("xmlns", "o", null, OpenXmlNamespaces.Office);
            _writer.WriteAttributeString("xmlns", "w10", null, OpenXmlNamespaces.OfficeWord);
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            _writer.WriteStartElement("w", "body", OpenXmlNamespaces.WordprocessingML);

            //convert the document
            // Handle Word95 files which may not have AllPapxFkps
            if (_doc.AllPapxFkps != null && _doc.AllPapxFkps.Count > 0 && 
                _doc.AllPapxFkps[0].grppapx != null && _doc.AllPapxFkps[0].grppapx.Length > 0)
            {
                _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];
            }
            else
            {
                // For Word95 files, create a default PAPX
                _lastValidPapx = new ParagraphPropertyExceptions();
            }

            // Enhanced debugging: Track character position coverage
            var coverageMap = new bool[doc.FIB.ccpText];
            int paragraphCount = 0;
            int tableCount = 0;
            int skippedPositions = 0;

            // Debug for fast-saved documents only when needed
            System.Console.WriteLine($"[DEBUG] Document has {doc.FIB.ccpText} characters to process");
            System.Console.WriteLine($"[DEBUG] Fast save count: {doc.FIB.cQuickSaves}");
            System.Console.WriteLine($"[DEBUG] Fast save flag: {doc.FIB.fFastSaved}");
            
            // FAST-SAVE RECONSTRUCTION: Piece table analysis for debugging
            System.Console.WriteLine($"[PIECE] Document has {doc.FIB.cQuickSaves} quick saves");
            System.Console.WriteLine($"[PIECE] Piece table has {_doc.PieceTable.Pieces.Count} pieces");
            System.Console.WriteLine($"[PIECE] FC range: {doc.FIB.fcMin} - {doc.FIB.fcMac}");
            
            // Verify piece table coverage
            int totalPieceCoverage = 0;
            for (int i = 0; i < _doc.PieceTable.Pieces.Count; i++)
            {
                var piece = _doc.PieceTable.Pieces[i];
                int pieceLength = piece.cpEnd - piece.cpStart;
                totalPieceCoverage += pieceLength;
                System.Console.WriteLine($"[PIECE] Piece {i}: CP {piece.cpStart}-{piece.cpEnd} (length: {pieceLength}), FC: {piece.fc}");
            }
            System.Console.WriteLine($"[PIECE] Total piece coverage: {totalPieceCoverage} vs expected: {doc.FIB.ccpText}");

            int cp = 0;
            while (cp < doc.FIB.ccpText)
            {
                int lastCp = cp;
                
                // Check if we have a valid piece table entry
                if (!_doc.PieceTable.FileCharacterPositions.ContainsKey(cp))
                {
                    System.Console.WriteLine($"[WARNING] Missing piece table entry at CP {cp}, skipping to next");
                    skippedPositions++;
                    cp++;
                    continue;
                }

                int fc = _doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                System.Console.WriteLine($"[DEBUG] Processing CP {cp}, FC {fc}, InTable: {tai.fInTable}");

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    System.Console.WriteLine($"[DEBUG] Processing table at CP {cp}");
                    cp = writeTable(cp, tai.iTap);
                    tableCount++;
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    System.Console.WriteLine($"[DEBUG] Processing paragraph at CP {cp}");
                    int newCp = writeParagraph(cp);
                    System.Console.WriteLine($"[DEBUG] Paragraph processing returned CP {newCp}");
                    cp = newCp;
                    paragraphCount++;
                }

                // Mark coverage
                for (int i = lastCp; i < cp && i < coverageMap.Length; i++)
                {
                    coverageMap[i] = true;
                }

                // Safety check to prevent infinite loops
                if (cp == lastCp)
                {
                    System.Console.WriteLine($"[ERROR] No progress made at CP {cp}, advancing by 1 to prevent infinite loop");
                    cp++;
                }

                System.Console.WriteLine($"[DEBUG] Advanced from CP {lastCp} to {cp}");
            }

            // Report coverage gaps
            var uncoveredRanges = new List<(int start, int end)>();
            int rangeStart = -1;
            for (int i = 0; i < coverageMap.Length; i++)
            {
                if (!coverageMap[i] && rangeStart == -1)
                {
                    rangeStart = i;
                }
                else if (coverageMap[i] && rangeStart != -1)
                {
                    uncoveredRanges.Add((rangeStart, i - 1));
                    rangeStart = -1;
                }
            }
            if (rangeStart != -1)
            {
                uncoveredRanges.Add((rangeStart, coverageMap.Length - 1));
            }

            System.Console.WriteLine($"[DEBUG] Processing complete: {paragraphCount} paragraphs, {tableCount} tables, {skippedPositions} skipped positions");
            if (uncoveredRanges.Count > 0)
            {
                System.Console.WriteLine($"[WARNING] Found {uncoveredRanges.Count} uncovered character ranges:");
                foreach (var range in uncoveredRanges)
                {
                    int rangeSize = range.end - range.start + 1;
                    System.Console.WriteLine($"  CP {range.start}-{range.end} ({rangeSize} characters)");
                }
            }
            else
            {
                System.Console.WriteLine("[DEBUG] All character positions were processed successfully");
            }

            //write the section properties of the body with the last SEPX
            // Handle Word95 files which may not have AllSepx
            if (_doc.AllSepx != null && _doc.AllSepx.Count > 0)
            {
                int lastSepxCp = 0;
                foreach (int sepxCp in _doc.AllSepx.Keys)
                    lastSepxCp = sepxCp;
                
                var lastSepx = _doc.AllSepx[lastSepxCp];
                lastSepx.Convert(new SectionPropertiesMapping(_writer, _ctx, _sectionNr));
            }

            //end the document
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
