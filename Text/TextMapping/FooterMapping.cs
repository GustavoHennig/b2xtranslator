using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.txt.TextModel;

namespace b2xtranslator.txt.TextMapping
{
    public class FooterMapping : DocumentMapping
    {
        private CharacterRange _ftr;

        public FooterMapping(ConversionContext ctx, IWriter writer, CharacterRange ftr)
            : base(ctx, writer)
        {
            _ftr = ftr;
        }
        
        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "ftr", OpenXmlNamespaces.WordprocessingML);

            //convert the footer text
            if(_doc.AllPapxFkps[0].grppapx.Length == 0)
            {
                return;
            }
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];
            int cp = _ftr.CharacterPosition;
            int cpMax = _ftr.CharacterPosition + _ftr.CharacterCount;

            //the CharacterCount of the footers also counts the guard paragraph mark.
            //this additional paragraph mark shall not be converted.
            cpMax--;

            while (cp < cpMax)
            {
                int fc = _doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                if (papx == null)
                {
                    // If no valid PAPX is found, skip to the next character position
                    cp++;
                    continue;
                }
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
            _writer.WriteEndDocument();

            _writer.Flush();
        }
    }
}
