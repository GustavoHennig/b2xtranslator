using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;

namespace b2xtranslator.txt.TextMapping
{
    public class TextboxMapping : DocumentMapping
    {
        public static int TextboxCount = 0;
        private int _textboxIndex;

        public TextboxMapping(ConversionContext ctx, int textboxIndex, IWriter writer)
            : base(ctx, writer)
        {
            TextboxCount++;
            _textboxIndex = textboxIndex;
        }

        public TextboxMapping(ConversionContext ctx, IWriter writer)
            : base(ctx, writer)
        {
            TextboxCount++;
            _textboxIndex = TextboxCount - 1;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartElement("v", "textbox", OpenXmlNamespaces.VectorML);
            _writer.WriteStartElement("w", "txbxContent", OpenXmlNamespaces.WordprocessingML);

            int cp = 0;
            int cpEnd = 0;
            BreakDescriptor bkd = null;
            int txtbxSubdocStart = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn + doc.FIB.ccpEdn;

            if(_writer.GetType() == typeof(MainDocumentPart)) // FIXME (Types not used)
            {
                cp = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[_textboxIndex];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlex.CharacterPositions[_textboxIndex + 1];
                bkd = doc.TextboxBreakPlex.Elements[_textboxIndex];
            }
            if (_writer.GetType() == typeof(HeaderPart) || _writer.GetType() == typeof(FooterPart)) // FIXME (Types not used)
            {
                txtbxSubdocStart += doc.FIB.ccpTxbx;
                cp = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[_textboxIndex];
                cpEnd = txtbxSubdocStart + doc.TextboxBreakPlexHeader.CharacterPositions[_textboxIndex + 1];
                bkd = doc.TextboxBreakPlexHeader.Elements[_textboxIndex];
            }

            //convert the textbox text
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];

            while (cp < cpEnd)
            {
                int fc = _doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                var tai = new TableInfo(papx);

                if (tai.fInTable)
                {
                    //this PAPX is for a table
                    cp = writeTable(cp, tai.iTap);
                }
                else
                {
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
