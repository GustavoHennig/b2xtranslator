using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.txt.TextModel;
using b2xtranslator.Tools;
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
            TraceLogger.Debug("[DEBUG] Document has {0} characters to process", doc.FIB.ccpText);
            TraceLogger.Debug("[DEBUG] Fast save count: {0}", doc.FIB.cQuickSaves);
            TraceLogger.Debug("[DEBUG] Fast save flag: {0}", doc.FIB.fFastSaved);
            
            // FAST-SAVE RECONSTRUCTION: Piece table analysis for debugging
            TraceLogger.Debug("[PIECE] Document has {0} quick saves", doc.FIB.cQuickSaves);
            TraceLogger.Debug("[PIECE] Piece table has {0} pieces", _doc.PieceTable.Pieces.Count);
            TraceLogger.Debug("[PIECE] FC range: {0} - {1}", doc.FIB.fcMin, doc.FIB.fcMac);
            
            // Verify piece table coverage
            int totalPieceCoverage = 0;
            for (int i = 0; i < _doc.PieceTable.Pieces.Count; i++)
            {
                var piece = _doc.PieceTable.Pieces[i];
                int pieceLength = piece.cpEnd - piece.cpStart;
                totalPieceCoverage += pieceLength;
                TraceLogger.Debug("[PIECE] Piece {0}: CP {1}-{2} (length: {3}), FC: {4}", i, piece.cpStart, piece.cpEnd, pieceLength, piece.fc);
            }
            TraceLogger.Debug("[PIECE] Total piece coverage: {0} vs expected: {1}", totalPieceCoverage, doc.FIB.ccpText);

            int cp = 0;
            while (cp < doc.FIB.ccpText)
            {
                int lastCp = cp;
                
                // Check if we have a valid piece table entry
                if (!_doc.PieceTable.FileCharacterPositions.ContainsKey(cp))
                {
                    TraceLogger.Warning("Missing piece table entry at CP {0}, skipping to next", cp);
                    skippedPositions++;
                    cp++;
                    continue;
                }

                int fc = _doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                TraceLogger.Debug("[DEBUG] Processing CP {0}, FC {1}, InTable: {2}", cp, fc, tai.fInTable);

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    TraceLogger.Debug("[DEBUG] Processing table at CP {0}", cp);
                    cp = writeTable(cp, tai.iTap);
                    tableCount++;
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    TraceLogger.Debug("[DEBUG] Processing paragraph at CP {0}", cp);
                    int newCp = writeParagraph(cp);
                    TraceLogger.Debug("[DEBUG] Paragraph processing returned CP {0}", newCp);
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
                    TraceLogger.Error("No progress made at CP {0}, advancing by 1 to prevent infinite loop", cp);
                    cp++;
                }

                TraceLogger.Debug("[DEBUG] Advanced from CP {0} to {1}", lastCp, cp);
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

            TraceLogger.Debug("[DEBUG] Processing complete: {0} paragraphs, {1} tables, {2} skipped positions", paragraphCount, tableCount, skippedPositions);
            if (uncoveredRanges.Count > 0)
            {
                TraceLogger.Warning("Found {0} uncovered character ranges:", uncoveredRanges.Count);
                foreach (var range in uncoveredRanges)
                {
                    int rangeSize = range.end - range.start + 1;
                    TraceLogger.Warning("  CP {0}-{1} ({2} characters)", range.start, range.end, rangeSize);
                }
            }
            else
            {
                TraceLogger.Debug("[DEBUG] All character positions were processed successfully");
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
