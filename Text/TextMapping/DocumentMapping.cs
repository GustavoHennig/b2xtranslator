using System;
using System.Collections.Generic;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.WordprocessingMLMapping;

namespace b2xtranslator.txt.TextMapping
{
    public abstract class DocumentMapping :
        AbstractMapping,
        IMapping<WordDocument>
    {


        protected WordDocument _doc;
        protected ConversionContext _ctx;
        protected ParagraphPropertyExceptions _lastValidPapx;
        protected SectionPropertyExceptions _lastValidSepx;
        protected int _skipRuns = 0;
        protected int _sectionNr = 0;
        protected int _footnoteNr = 0;
        protected int _endnoteNr = 0;
        protected int _commentNr = 0;
        protected bool _writeInstrText = false;
        //protected ContentPart _targetPart;

        private class Symbol
        {
            public string FontName;
            public string HexValue;
        }

        /// <summary>
        /// Creates a new DocumentMapping that writes to the given IWriter
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="targetPart"></param>
        public DocumentMapping(ConversionContext ctx, IWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        public abstract void Apply(WordDocument doc);

        #region TableConversion

        /// <summary>
        /// Writes the table starts at the given cp value
        /// </summary>
        /// <param name="cp">The cp at where the table begins</param>
        /// <returns>The character pointer to the first character after this table</returns>
        protected int writeTable(int initialCp, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = _doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            //build the table grid
            var grid = buildTableGrid(cp, nestingLevel);

            //find first row end
            int fcRowEnd = findRowEndFc(cp, nestingLevel);
            var row1Tapx = new TablePropertyExceptions(findValidPapx(fcRowEnd), _doc.DataStream);

            //start table
            _writer.WriteStartElement("w", "tbl", OpenXmlNamespaces.WordprocessingML);

            //Convert it
            row1Tapx.Convert(new TablePropertiesMapping(_writer, _doc.Styles, grid));

            //convert all rows
            if (nestingLevel > 1)
            {
                //It's an inner table
                //only convert the cells with the given nesting level
                while (tai.iTap == nestingLevel)
                {
                    cp = writeTableRow(cp, grid, nestingLevel);
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }
            else
            {
                //It's a outer table (nesting level 1)
                //convert until the end of table is reached
                while (tai.fInTable)
                {
                    cp = writeTableRow(cp, grid, nestingLevel);
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }

            //close w:tbl
            _writer.WriteEndElement();

            return cp;
        }

        /// <summary>
        /// Writes the table row that starts at the given cp value and ends at the next row end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the row begins</param>
        /// <returns>The character pointer to the first character after this row</returns>
        protected int writeTableRow(int initialCp, List<short> grid, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = _doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            //start w:tr
            _writer.WriteStartElement("w", "tr", OpenXmlNamespaces.WordprocessingML);

            //convert the properties
            int fcRowEnd = findRowEndFc(cp, nestingLevel);
            var rowEndPapx = findValidPapx(fcRowEnd);
            var tapx = new TablePropertyExceptions(rowEndPapx, _doc.DataStream);
            var chpxs = _doc.GetCharacterPropertyExceptions(fcRowEnd, fcRowEnd + 1);

            if (tapx != null)
            {
                tapx.Convert(new TableRowPropertiesMapping(_writer, chpxs[0]));
            }
            int gridIndex = 0;
            int cellIndex = 0;

            if (nestingLevel > 1)
            {
                //It's an inner table.
                //Write until the first "inner trailer paragraph" is reached
                while (!(_doc.Text[cp] == TextMark.ParagraphEnd && tai.fInnerTtp) && tai.fInTable)
                {
                    cp = writeTableCell(cp, tapx, grid, ref gridIndex, cellIndex, nestingLevel);
                    cellIndex++;

                    //each cell has it's own PAPX
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }
            else
            {
                //It's a outer table
                //Write until the first "row end trailer paragraph" is reached
                while (!(_doc.Text[cp] == TextMark.CellOrRowMark && tai.fTtp) && tai.fInTable)
                {
                    cp = writeTableCell(cp, tapx, grid, ref gridIndex, cellIndex, nestingLevel);
                    cellIndex++;

                    //each cell has it's own PAPX
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
            }


            //end w:tr
            _writer.WriteEndElement();

            //skip the row end mark
            cp++;

            return cp;
        }


        /// <summary>
        /// Writes the table cell that starts at the given cp value and ends at the next cell end mark
        /// </summary>
        /// <param name="initialCp">The cp at where the cell begins</param>
        /// <param name="tapx">The TAPX that formats the row to which the cell belongs</param>
        /// <param name="gridIndex">The index of this cell in the grid</param>
        /// <param name="gridIndex">The grid</param>
        /// <returns>The character pointer to the first character after this cell</returns>
        protected int writeTableCell(int initialCp, TablePropertyExceptions tapx, List<short> grid, ref int gridIndex, int cellIndex, uint nestingLevel)
        {
            int cp = initialCp;

            //start w:tc
            _writer.WriteStartElement("w", "tc", OpenXmlNamespaces.WordprocessingML);

            //find cell end
            int cpCellEnd = findCellEndCp(initialCp, nestingLevel);
            
            //convert the properties
            var mapping = new TableCellPropertiesMapping(_writer, grid, gridIndex, cellIndex);
            if (tapx != null)
            {
                tapx.Convert(mapping);
            }
            gridIndex = gridIndex + mapping.GridSpan;
            

            //write the paragraphs of the cell
            while (cp < cpCellEnd)
            {
                //cp = writeParagraph(cp);
                int fc = _doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                //cp = writeParagraph(cp);

                if (tai.iTap > nestingLevel)
                {
                    //write the inner table if this is not a inner table (endless loop)
                    cp = writeTable(cp, tai.iTap);

                    //after a inner table must be at least one paragraph
                    //if (cp >= cpCellEnd)
                    //{
                    //    _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);
                    //    _writer.WriteEndElement();
                    //}
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }


            //end w:tc
            _writer.WriteEndElement();

            return cp;
        }


        /// <summary>
        /// Builds a list that contains the width of the several columns of the table.
        /// </summary>
        /// <param name="initialCp"></param>
        /// <returns></returns>
        protected List<short> buildTableGrid(int initialCp, uint nestingLevel)
        {
            var backup = _lastValidPapx;

            var boundaries = new List<short>();
            var grid = new List<short>();
            int cp = initialCp;
            int fc = _doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            int fcRowEnd = findRowEndFc(cp, out cp, nestingLevel);

            while (tai.fInTable)
            {
                //TODO: Eventual infinite loop

                //check all SPRMs of this TAPX
                foreach (var sprm in papx.grpprl)
                {
                    //find the tDef SPRM
                    if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmTDefTable)
                    {
                        byte itcMac = sprm.Arguments[0];
                        for (int i = 0; i < itcMac; i++)
                        {
                            short boundary1 = BitConverter.ToInt16(sprm.Arguments, 1 + i * 2);
                            if (!boundaries.Contains(boundary1))
                                boundaries.Add(boundary1);

                            short boundary2 = BitConverter.ToInt16(sprm.Arguments, 1 + (i + 1) * 2);
                            if (!boundaries.Contains(boundary2))
                                boundaries.Add(boundary2);
                        }
                    }
                }

                //get the next papx
                papx = findValidPapx(fcRowEnd);
                tai = new TableInfo(papx);
                fcRowEnd = findRowEndFc(cp, out cp, nestingLevel);
            }

            //build the grid based on the boundaries
            boundaries.Sort();
            for (int i = 0; i < boundaries.Count - 1; i++)
            {
                grid.Add((short)(boundaries[i + 1] - boundaries[i]));
            }

            _lastValidPapx = backup;
            return grid;
        }

        /// <summary>
        /// Finds the FC of the next row end mark.
        /// </summary>
        /// <param name="initialCp">Some CP before the row end</param>
        /// <param name="rowEndCp">The CP of the next row end mark</param>
        /// <returns>The FC of the next row end mark</returns>
        protected int findRowEndFc(int initialCp, out int rowEndCp, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = _doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            if (nestingLevel > 1)
            {
                //Its an inner table.
                //Search the "inner table trailer paragraph"
                while (tai.fInnerTtp == false && tai.fInTable == true)
                {
                    while (_doc.Text[cp] != TextMark.ParagraphEnd)
                    {
                        cp++;
                    }
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }
            else 
            {
                //Its an outer table.
                //Search the "table trailer paragraph"
                while (tai.fTtp == false && tai.fInTable == true)
                {
                    while (_doc.Text[cp] != TextMark.CellOrRowMark)
                    {
                        cp++;
                    }
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }

            rowEndCp = cp;
            return fc;
        }

        /// <summary>
        /// Finds the FC of the next row end mark.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        protected int findRowEndFc(int initialCp, uint nestingLevel)
        {
            int cp = initialCp;
            int fc = _doc.PieceTable.FileCharacterPositions[cp];
            var papx = findValidPapx(fc);
            var tai = new TableInfo(papx);

            if (nestingLevel > 1)
            {
                //Its an inner table.
                //Search the "inner table trailer paragraph"
                while (tai.fInnerTtp == false && tai.fInTable == true)
                {
                    while (_doc.Text[cp] != TextMark.ParagraphEnd)
                    {
                        cp++;
                    }
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }
            else
            {
                //Its an outer table.
                //Search the "table trailer paragraph"
                while (tai.fTtp == false && tai.fInTable == true)
                {
                    while (_doc.Text[cp] != TextMark.CellOrRowMark)
                    {
                        cp++;
                    }
                    fc = _doc.PieceTable.FileCharacterPositions[cp];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                    cp++;
                }
            }

            return fc;
        }


        protected int findCellEndCp(int initialCp, uint nestingLevel)
        {
            int cpCellEnd = initialCp;

            if (nestingLevel > 1)
            {
                int fc = _doc.PieceTable.FileCharacterPositions[initialCp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                while (!tai.fInnerTableCell)
                {
                    cpCellEnd++;

                    fc = _doc.PieceTable.FileCharacterPositions[cpCellEnd];
                    papx = findValidPapx(fc);
                    tai = new TableInfo(papx);
                }
                cpCellEnd++;
            }
            else 
            {
                while (_doc.Text[cpCellEnd] != TextMark.CellOrRowMark)
                {
                    cpCellEnd++;
                }
                cpCellEnd++;
            }

            return cpCellEnd;
        }


        #endregion

        #region ParagraphRunConversion

        /// <summary>
        /// Writes a Paragraph that starts at the given cp and 
        /// ends at the next paragraph end mark or section end mark
        /// </summary>
        /// <param name="cp"></param>
        protected int writeParagraph(int cp) 
        {
            //search the paragraph end
            int cpParaEnd = cp;
            while (_doc.Text[cpParaEnd] != TextMark.ParagraphEnd &&
                _doc.Text[cpParaEnd] != TextMark.CellOrRowMark &&
                !(_doc.Text[cpParaEnd] == TextMark.PageBreakOrSectionMark && isSectionEnd(cpParaEnd)))
            {
                cpParaEnd++;
            }

            if (_doc.Text[cpParaEnd] == TextMark.PageBreakOrSectionMark)
            {
                //there is a page break OR section mark,
                //write the section only if it's a section mark
                bool sectionEnd = isSectionEnd(cpParaEnd);
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, sectionEnd);
            }
            else
            {
                cpParaEnd++;
                return writeParagraph(cp, cpParaEnd, false);
            }
        }

        /// <summary>
        /// Writes a Paragraph that starts at the given cpStart and 
        /// ends at the given cpEnd
        /// </summary>
        /// <param name="cpStart"></param>
        /// <param name="cpEnd"></param>
        /// <param name="sectionEnd">Set if this paragraph is the last paragraph of a section</param>
        /// <returns></returns>
        protected int writeParagraph(int initialCp, int cpEnd, bool sectionEnd)
        {
            int cp = initialCp;
            int fc = _doc.PieceTable.FileCharacterPositions[cp];
            int fcEnd = _doc.PieceTable.FileCharacterPositions[cpEnd];
            var papx = findValidPapx(fc);

            //get all CHPX between these boundaries to determine the count of runs
            var chpxs = _doc.GetCharacterPropertyExceptions(fc, fcEnd);
            var chpxFcs = _doc.GetFileCharacterPositions(fc, fcEnd);
            chpxFcs.Add(fcEnd);

            //the last of these CHPX formats the paragraph end mark
            var paraEndChpx = chpxs[chpxs.Count-1];

            //start paragraph
            _writer.WriteStartElement("w", "p", OpenXmlNamespaces.WordprocessingML);

            //check for section properties
            if (sectionEnd)
            {
                //this is the last paragraph of this section
                //write properties with section properties
                if (papx != null)
                {
                    papx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, _doc, paraEndChpx, findValidSepx(cpEnd), _sectionNr));
                }
                _sectionNr++;
            }
            else
            {
                //write properties
                if (papx != null)
                {
                    papx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, _doc, paraEndChpx));
                }
            }

            //write a runs for each CHPX
            for (int i = 0; i < chpxs.Count; i++)
            {
                //get the FC range for this run
                int fcChpxStart = chpxFcs[i];
                int fcChpxEnd = chpxFcs[i + 1];

                //it's the first chpx and it starts before the paragraph
                if (i == 0 && fcChpxStart < fc)
                {
                    //so use the FC of the paragraph
                    fcChpxStart = fc;
                }

                //it's the last chpx and it exceeds the paragraph
                if (i == chpxs.Count - 1 && fcChpxEnd > fcEnd)
                {
                    //so use the FC of the paragraph
                    fcChpxEnd = fcEnd;
                }

                //read the chars that are formatted via this CHPX
                var chpxChars = _doc.PieceTable.GetChars(fcChpxStart, fcChpxEnd, _doc.WordDocumentStream);

                //search for bookmarks in the chars
                var bookmarks = searchBookmarks(chpxChars, cp);

                //if there are bookmarks in this run, split the run into several runs
                if (bookmarks.Count > 0)
                {
                    var runs = splitCharList(chpxChars, bookmarks);
                    for (int s = 0; s < runs.Count; s++)
                    {
                        if (_doc.BookmarkStartPlex.CharacterPositions.Contains(cp) &&
                            _doc.BookmarkEndPlex.CharacterPositions.Contains(cp))
                        {
                            //there start and end bookmarks here

                            //so get all bookmarks that end here
                            for (int b = 0; b < _doc.BookmarkEndPlex.CharacterPositions.Count; b++)
                            {
                                if (_doc.BookmarkEndPlex.CharacterPositions[b] == cp)
                                {
                                    //and check if the matching start bookmark also starts here
                                    if (_doc.BookmarkStartPlex.CharacterPositions[b] == cp)
                                    {
                                        //then write a start and a end
                                        if (_doc.BookmarkStartPlex.Elements.Count > b)
                                        {
                                            writeBookmarkStart(_doc.BookmarkStartPlex.Elements[b]);
                                            writeBookmarkEnd(_doc.BookmarkStartPlex.Elements[b]);
                                        }
                                    }
                                    else
                                    {
                                        //write a end
                                        writeBookmarkEnd(_doc.BookmarkStartPlex.Elements[b]);
                                    }
                                }
                            }

                            writeBookmarkStarts(cp);
                        }
                        else if (_doc.BookmarkStartPlex.CharacterPositions.Contains(cp))
                        {
                            writeBookmarkStarts(cp);
                        }
                        else if (_doc.BookmarkEndPlex.CharacterPositions.Contains(cp))
                        {
                            writeBookmarkEnds(cp);
                        }

                        cp = writeRun(runs[s], chpxs[i], cp);
                    }
                }
                else
                {
                    cp = writeRun(chpxChars, chpxs[i], cp);
                }
            }

            //end paragraph
            _writer.WriteEndElement();

            return cpEnd++;
        }

        /// <summary>
        /// Writes a run with the given characters and CHPX
        /// </summary>
        protected int writeRun(List<char> chars, CharacterPropertyExceptions chpx, int initialCp)
        {
            int cp = initialCp;

            if (_skipRuns <= 0 && chars.Count > 0)
            {
                var rev = new RevisionData(chpx);

                if (rev.Type == RevisionData.RevisionType.Deleted)
                {
                    //If it's a deleted run
                    _writer.WriteStartElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, "[b2x: could not retrieve author]");
                    _writer.WriteAttributeString("w", "date", OpenXmlNamespaces.WordprocessingML, "[b2x: could not retrieve date]");
                }
                else if (rev.Type == RevisionData.RevisionType.Inserted)
                {
                    //if it's a inserted run
                    _writer.WriteStartElement("w", "ins", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, _doc.RevisionAuthorTable.Strings[rev.Isbt]);
                    rev.Dttm?.Convert(new DateMapping(_writer));
                }

                //start run
                _writer.WriteStartElement("w", "r", OpenXmlNamespaces.WordprocessingML);

                //append rsids
                if (rev.Rsid != 0)
                {
                    string rsid = string.Format("{0:x8}", rev.Rsid);
                    _writer.WriteAttributeString("w", "rsidR", OpenXmlNamespaces.WordprocessingML, rsid);
                    _ctx.AddRsid(rsid);
                }
                if (rev.RsidDel != 0)
                {
                    string rsidDel = string.Format("{0:x8}", rev.RsidDel);
                    _writer.WriteAttributeString("w", "rsidDel", OpenXmlNamespaces.WordprocessingML, rsidDel);
                    _ctx.AddRsid(rsidDel);
                }
                if (rev.RsidProp != 0)
                {
                    string rsidProp = string.Format("{0:x8}", rev.RsidProp);
                    _writer.WriteAttributeString("w", "rsidRPr", OpenXmlNamespaces.WordprocessingML, rsidProp);
                    _ctx.AddRsid(rsidProp);
                }

                //convert properties
                chpx.Convert(new CharacterPropertiesMapping(_writer, _doc, rev, _lastValidPapx, false));

                if (rev.Type == RevisionData.RevisionType.Deleted)
                    writeText(chars, cp, chpx, true);
                else
                    writeText(chars, cp, chpx, false);

                //end run
                _writer.WriteEndElement();

                if (rev.Type == RevisionData.RevisionType.Deleted || rev.Type == RevisionData.RevisionType.Inserted)
                {
                    _writer.WriteEndElement();
                }
            }
            else
            {
                _skipRuns--;
            }

            return cp + chars.Count;
        }


        /// <summary>
        /// Writes the given text to the document
        /// </summary>
        /// <param name="chars"></param>
        protected void writeText(List<char> chars, int initialCp, CharacterPropertyExceptions chpx, bool writeDeletedText)
        {
            int cp = initialCp;
            bool fSpec = isSpecial(chpx);

            //detect text type
            string textType = "t";
            if(writeDeletedText)
                textType = "delText";
            else if(_writeInstrText)
                textType = "instrText";
 
            //open a new w:t element
            writeTextStart(textType);

            //write text
            for (int i = 0; i < chars.Count; i++)
            {
                char c = chars[i];

                if (c == TextMark.Tab)
                {
                    _writer.WriteEndElement();
                    _writer.WriteElementString("w", "tab", OpenXmlNamespaces.WordprocessingML, "");
                    writeTextStart(textType);
                }
                else if (c == TextMark.HardLineBreak)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();
                    _writer.WriteElementString("w", "br", OpenXmlNamespaces.WordprocessingML, "");

                    writeTextStart(textType);
                }
                else if (c == TextMark.ParagraphEnd)
                {
                    _writer.WriteChar(c);
                    //do nothing
                }
                else if (c == TextMark.PageBreakOrSectionMark)
                {
                    //write page break, section breaks are written by writeParagraph() method
                    if (!isSectionEnd(cp))
                    {
                        //close previous w:t ...
                        _writer.WriteEndElement();

                        _writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "page");
                        _writer.WriteEndElement();

                        writeTextStart(textType);
                    }
                }
                else if (c == TextMark.ColumnBreak)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("w", "br", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "column");
                    _writer.WriteEndElement();

                    writeTextStart(textType);
                }
                else if (c == TextMark.FieldBeginMark)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    int cpFieldStart = initialCp + i;
                    int cpFieldEnd = searchNextTextMark(_doc.Text, cpFieldStart, TextMark.FieldEndMark);
                    var f = new Field(_doc.Text.GetRange(cpFieldStart, cpFieldEnd - cpFieldStart + 1));

                    if(f.FieldCode.StartsWith(" FORM"))
                    {
                        _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "begin");

                        int cpPic = searchNextTextMark(_doc.Text, cpFieldStart, TextMark.Picture);
                        if (cpPic < cpFieldEnd)
                        {
                            int fcPic = _doc.PieceTable.FileCharacterPositions[cpPic];
                            var chpxPic = _doc.GetCharacterPropertyExceptions(fcPic, fcPic + 1)[0];
                            var npbd = new NilPicfAndBinData(chpxPic, _doc.DataStream);
                            var ffdata = new FormFieldData(npbd.binData);
                            ffdata.Convert(new FormFieldDataMapping(_writer));
                        }

                        _writer.WriteEndElement();
                    }
                    else if (f.FieldCode.StartsWith(" EMBED") || f.FieldCode.StartsWith(" LINK"))
                    {
                        _writer.WriteStartElement("w", "object", OpenXmlNamespaces.WordprocessingML);

                        int cpPic = searchNextTextMark(_doc.Text, cpFieldStart, TextMark.Picture);
                        int cpFieldSep = searchNextTextMark(_doc.Text, cpFieldStart, TextMark.FieldSeperator);

                        if (cpPic < cpFieldEnd)
                        {
                            int fcPic = _doc.PieceTable.FileCharacterPositions[cpPic];
                            var chpxPic = _doc.GetCharacterPropertyExceptions(fcPic, fcPic + 1)[0];
                            var pic = new PictureDescriptor(chpxPic, _doc.DataStream);

                            //append the origin attributes
                            _writer.WriteAttributeString("w", "dxaOrig", OpenXmlNamespaces.WordprocessingML, (pic.dxaGoal + pic.dxaOrigin).ToString());
                            _writer.WriteAttributeString("w", "dyaOrig", OpenXmlNamespaces.WordprocessingML, (pic.dyaGoal + pic.dyaOrigin).ToString());

                            pic.Convert(new VMLPictureMapping(_writer, true));

                            if (cpFieldSep < cpFieldEnd)
                            {
                                int fcFieldSep = _doc.PieceTable.FileCharacterPositions[cpFieldSep];
                                var chpxSep = _doc.GetCharacterPropertyExceptions(fcFieldSep, fcFieldSep + 1)[0];
                                var ole = new OleObject(chpxSep, _doc.Storage);
                                ole.Convert(new OleObjectMapping(_writer, _doc, pic));
                            }
                        }

                        _writer.WriteEndElement();

                        _skipRuns = 4;
                    }
                    else
                    {
                        _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "begin");
                        _writer.WriteEndElement();
                    }

                    _writeInstrText = true;

                    writeTextStart("instrText");
                }
                else if (c == TextMark.FieldSeperator)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "separate");
                    _writer.WriteEndElement();

                    writeTextStart(textType);
                }
                else if (c == TextMark.FieldEndMark)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    _writer.WriteStartElement("w", "fldChar", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "fldCharType", OpenXmlNamespaces.WordprocessingML, "end");
                    _writer.WriteEndElement();

                    _writeInstrText = false;

                    writeTextStart("t");
                }
                else if (c == TextMark.Symbol && fSpec)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    var s = getSymbol(chpx);
                    _writer.WriteStartElement("w", "sym", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "font", OpenXmlNamespaces.WordprocessingML, s.FontName);
                    _writer.WriteAttributeString("w", "char", OpenXmlNamespaces.WordprocessingML, s.HexValue);
                    _writer.WriteEndElement();

                    writeTextStart(textType);
                }
                else if (c == TextMark.DrawnObject && fSpec)
                {
                    FileShapeAddress fspa = null;
                    if (GetType() == typeof(MainDocumentMapping))
                    {
                        fspa = _doc.OfficeDrawingPlex.GetStruct(cp);
                    }
                    else if (GetType() == typeof(HeaderMapping) || GetType() == typeof(FooterMapping))
                    {
                        int headerCp = cp - _doc.FIB.ccpText - _doc.FIB.ccpFtn;
                        fspa = _doc.OfficeDrawingPlexHeader.GetStruct(headerCp);
                    }
                    if (fspa != null)
                    {
                        var shape = _doc.OfficeArtContent.GetShapeContainer(fspa.spid);
                        if (shape != null)
                        {
                            //close previous w:t ...
                            _writer.WriteEndElement();
                            _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

                            shape.Convert(new VMLShapeMapping(_writer, fspa, null, _ctx));

                            _writer.WriteEndElement();
                            writeTextStart(textType);
                        }
                    }
                }
                else if (c == TextMark.Picture && fSpec)
                {
                    var pict = new PictureDescriptor(chpx, _doc.DataStream);
                    if (pict.mfp.mm > 98 && pict.ShapeContainer != null)
                    {
                        //close previous w:t ...
                        _writer.WriteEndElement();
                        _writer.WriteStartElement("w", "pict", OpenXmlNamespaces.WordprocessingML);

                        if (isWordArtShape(pict.ShapeContainer))
                        {
                            // a PICT without a BSE can stand for a WordArt Shape
                            pict.ShapeContainer.Convert(new VMLShapeMapping(_writer, null, pict, _ctx));
                        }
                        else
                        {
                            // it's a normal picture
                            pict.Convert(new VMLPictureMapping(_writer, false));
                        }

                        _writer.WriteEndElement();
                        writeTextStart(textType);
                    }                   
                }
                else if (c == TextMark.AutoNumberedFootnoteReference && fSpec)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    if (GetType() != typeof(FootnotesMapping) && GetType() != typeof(EndnotesMapping))
                    {
                        //it's in the document
                        if (_doc.FootnoteReferencePlex.CharacterPositions.Contains(cp))
                        {
                            _writer.WriteStartElement("w", "footnoteReference", OpenXmlNamespaces.WordprocessingML);
                            _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, _footnoteNr.ToString());
                            _writer.WriteEndElement();
                            _footnoteNr++;
                        }
                        else if (_doc.EndnoteReferencePlex.CharacterPositions.Contains(cp))
                        {
                            _writer.WriteStartElement("w", "endnoteReference", OpenXmlNamespaces.WordprocessingML);
                            _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, _endnoteNr.ToString());
                            _writer.WriteEndElement();
                            _endnoteNr++;
                        }
                    }
                    else
                    {
                        // it's not the document, write the short ref
                        if(GetType() != typeof(FootnotesMapping))
                        {
                            _writer.WriteElementString("w", "footnoteRef", OpenXmlNamespaces.WordprocessingML, "");
                        }
                        if (GetType() != typeof(EndnotesMapping))
                        {
                            _writer.WriteElementString("w", "endnoteRef", OpenXmlNamespaces.WordprocessingML, "");
                        }
                    }

                    writeTextStart(textType);
                }
                else if (c == TextMark.AnnotationReference)
                {
                    //close previous w:t ...
                    _writer.WriteEndElement();

                    if (GetType() != typeof(CommentsMapping))
                    {
                        _writer.WriteStartElement("w", "commentReference", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, _commentNr.ToString());
                        _writer.WriteEndElement();
                    }
                    else
                    {
                        _writer.WriteElementString("w", "annotationRef", OpenXmlNamespaces.WordprocessingML, "");
                    }

                    _commentNr++;

                    writeTextStart(textType);
                }
                else if (c > 31 && c != 0xFFFF)
                {
                    // Convert Windows-1252 control characters to proper Unicode
                    // TODO: Review this, seems a poor workaround
                    char convertedChar = ConvertWindows1252ToUnicode(c);
                    _writer.WriteChars(new char[] { convertedChar }, 0, 1);
                }

                cp++;
            }

            //close w:t
            _writer.WriteEndElement();
        }


        protected void writeTextStart(string textType)
        {
            _writer.WriteStartElement("w", textType, OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("xml", "space", "", "preserve");
        }

        /// <summary>
        /// Writes a bookmark start element at the given position
        /// </summary>
        /// <param name="cp"></param>
        protected void writeBookmarkStarts(int cp)
        {
            if (_doc.BookmarkStartPlex.CharacterPositions.Count > 1)
            {
                for (int b = 0; b < _doc.BookmarkStartPlex.CharacterPositions.Count; b++)
                {
                    if (_doc.BookmarkStartPlex.CharacterPositions[b] == cp)
                    {
                        if (_doc.BookmarkStartPlex.Elements.Count > b)
                        {
                            writeBookmarkStart(_doc.BookmarkStartPlex.Elements[b]);
                        }
                    }
                }
            }
        }

        protected void writeBookmarkStart(BookmarkFirst bookmark)
        {
            //write bookmark start
            _writer.WriteStartElement("w", "bookmarkStart", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, bookmark.ibkl.ToString());
            _writer.WriteAttributeString("w", "name", OpenXmlNamespaces.WordprocessingML, _doc.BookmarkNames.Strings[bookmark.ibkl]);
            _writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a bookmark end element at the given position
        /// </summary>
        /// <param name="cp"></param>
        protected void writeBookmarkEnds(int cp)
        {
            if (_doc.BookmarkEndPlex.CharacterPositions.Count > 1)
            {
                //write all bookmark ends
                for (int b = 0; b < _doc.BookmarkEndPlex.CharacterPositions.Count; b++)
                {
                    if (_doc.BookmarkEndPlex.CharacterPositions[b] == cp)
                    {
                        writeBookmarkEnd(_doc.BookmarkStartPlex.Elements[b]);
                    }
                }     
            }
        }

        protected void writeBookmarkEnd(BookmarkFirst bookmark)
        {
            //write bookmark end
            _writer.WriteStartElement("w", "bookmarkEnd", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, bookmark.ibkl.ToString());
            _writer.WriteEndElement();
        }

        #endregion

        #region CharacterConversion

        /// <summary>
        /// Converts Windows-1252 control characters to proper Unicode equivalents
        /// </summary>
        /// <param name="c">The character to convert</param>
        /// <returns>The converted Unicode character</returns>
        protected char ConvertWindows1252ToUnicode(char c)
        {
            // Map Windows-1252 control characters (0x80-0x9F) to Unicode
            switch ((int)c)
            {
                case 0x91: return '\u2018'; // LEFT SINGLE QUOTATION MARK
                case 0x92: return '\u2019'; // RIGHT SINGLE QUOTATION MARK
                case 0x93: return '\u201C'; // LEFT DOUBLE QUOTATION MARK
                case 0x94: return '\u201D'; // RIGHT DOUBLE QUOTATION MARK
                case 0x95: return '\u2022'; // BULLET
                case 0x96: return '\u2013'; // EN DASH
                case 0x97: return '\u2014'; // EM DASH
                case 0x98: return '\u02DC'; // SMALL TILDE
                case 0x99: return '\u2122'; // TRADE MARK SIGN
                case 0x9A: return '\u0161'; // LATIN SMALL LETTER S WITH CARON
                case 0x9B: return '\u203A'; // SINGLE RIGHT-POINTING ANGLE QUOTATION MARK
                case 0x9C: return '\u0153'; // LATIN SMALL LIGATURE OE
                case 0x9E: return '\u017E'; // LATIN SMALL LETTER Z WITH CARON
                case 0x9F: return '\u0178'; // LATIN CAPITAL LETTER Y WITH DIAERESIS
                default: return c; // Return the character unchanged if no mapping exists
            }
        }

        #endregion

        #region HelpFunctions

        protected bool isWordArtShape(ShapeContainer shape)
        {
            bool result = false;
            var options = shape.ExtractOptions();
            foreach (var entry in options)
            {
                if (entry.pid == ShapeOptions.PropertyId.gtextUNICODE)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// Splits a list of characters into several lists
        /// </summary>
        /// <returns></returns>
        protected List<List<char>> splitCharList(List<char> chars, List<int> splitIndices)
        {
            var ret = new List<List<char>>();

            int startIndex = 0;

            //add the parts
            for (int i = 0; i < splitIndices.Count; i++)
            {
                int cch = splitIndices[i] - startIndex;
                if (cch > 0)
                {
                    ret.Add(chars.GetRange(startIndex, cch));
                }
                startIndex += cch;
            }

            //add the last part
            ret.Add(chars.GetRange(startIndex, chars.Count-startIndex));

            return ret;
        }

        /// <summary>
        /// Searches for bookmarks in the list of characters.
        /// </summary>
        /// <param name="chars"></param>
        /// <returns>A List with all bookmarks indices in the given character list</returns>
        protected List<int> searchBookmarks(List<char> chars, int initialCp)
        {
            var ret = new List<int>();
            int cp = initialCp;
            for (int i = 0; i < chars.Count; i++)
            {
                if (_doc.BookmarkStartPlex.CharacterPositions.Contains(cp) ||
                    _doc.BookmarkEndPlex.CharacterPositions.Contains(cp))
                {
                    ret.Add(i);
                }
                cp++;
            }
            return ret;
        }

        /// <summary>
        /// Searches the given List for the next FieldEnd character.
        /// </summary>
        /// <param name="chars">The List of chars</param>
        /// <param name="initialCp">The position where the search should start</param>
        /// <param name="mark">The TextMark</param>
        /// <returns>The position of the next FieldEnd mark</returns>
        protected int searchNextTextMark(List<char> chars, int initialCp, char mark)
        {
            int ret = initialCp;
            for (int i = initialCp; i < chars.Count; i++)
            {
                if (chars[i] == mark)
                {
                    ret = i;
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Checks if the PAPX is old
        /// </summary>
        /// <param name="chpx">The PAPX</param>
        /// <returns></returns>
        protected bool isOld(ParagraphPropertyExceptions papx)
        {
            bool ret = false;
            foreach (var sprm in papx.grpprl)
            {
                if(sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPWall)
                {
                    //sHasOldProps
                    ret = true;
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Checks if the CHPX is special
        /// </summary>
        /// <param name="chpx">The CHPX</param>
        /// <returns></returns>
        protected bool isSpecial(CharacterPropertyExceptions chpx)
        {
            bool ret = false;
            foreach (var sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCPicLocation ||
                    sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCHsp)
                {
                    //special picture
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCSymbol)
                {
                    //special symbol
                    ret = true;
                    break;
                }
                else if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCFSpec)
                {
                    //special value
                    ret = Utils.ByteToBool(sprm.Arguments[0]);
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chpx"></param>
        /// <returns></returns>
        private Symbol getSymbol(CharacterPropertyExceptions chpx)
        {
            Symbol ret = null;
            foreach (var sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCSymbol)
                {
                    //special symbol
                    ret = new Symbol();
                    short fontIndex = BitConverter.ToInt16(sprm.Arguments, 0);
                    short code = BitConverter.ToInt16(sprm.Arguments, 2);

                    var ffn = (FontFamilyName)_doc.FontTable.Data[fontIndex];
                    ret.FontName = ffn.xszFtn;
                    ret.HexValue = string.Format("{0:x4}", code);
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Looks into the section table to find out if this CP is the end of a section
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        protected bool isSectionEnd(int cp)
        {
            bool result = false;

            //if cp is the last char of a section, the next section will start at cp +1
            int search = cp + 1;

            for (int i = 0; i < _doc.SectionPlex.CharacterPositions.Count; i++)
            {
                if (_doc.SectionPlex.CharacterPositions[i] == search)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Finds the PAPX that is valid for the given FC.
        /// </summary>
        /// <param name="fc"></param>
        /// <returns></returns>
        protected ParagraphPropertyExceptions findValidPapx(int fc)
        {
            ParagraphPropertyExceptions ret = null;

            if(_doc.AllPapx.ContainsKey(fc))
            {
                ret = _doc.AllPapx[fc];
                _lastValidPapx = ret;
            }
            else
            {
                ret = _lastValidPapx;
            }

            return ret;
        }

        /// <summary>
        /// Finds the SEPX that is valid for the given CP.
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        protected SectionPropertyExceptions findValidSepx(int cp)
        {
            SectionPropertyExceptions ret = null;

            try
            {
                ret = _doc.AllSepx[cp];
                _lastValidSepx = ret;
            }
            catch (KeyNotFoundException)
            {
                //there is no SEPX at this position, 
                //so the previous SEPX is valid for this cp

                int lastKey = _doc.SectionPlex.CharacterPositions[1];
                foreach (int key in _doc.AllSepx.Keys)
                {
                    if (cp > lastKey && cp < key)
                    {
                        ret = _doc.AllSepx[lastKey];
                        break;
                    }
                    else
                    {
                        lastKey = key;
                    }
                }
            }

            return ret;
        }

        #endregion
    }
}
