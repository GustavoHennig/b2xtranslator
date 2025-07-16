using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.WordprocessingML;

namespace b2xtranslator.txt.TextMapping
{
    public class HeaderMapping : DocumentMapping
    {
        private CharacterRange _hdr;

        public HeaderMapping(ConversionContext ctx, IWriter writer, CharacterRange hdr)
            : base(ctx, writer)
        {
            _hdr = hdr;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            if(_doc.AllPapxFkps[0].grppapx.Length == 0)
                {
                //if there are no PAPX, then there is nothing to convert
                return;
            }

            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "hdr", OpenXmlNamespaces.WordprocessingML);



            //convert the header text
            _lastValidPapx = _doc.AllPapxFkps[0].grppapx[0];
            int cp = _hdr.CharacterPosition;
            int cpMax = _hdr.CharacterPosition + _hdr.CharacterCount;

            //the CharacterCount of the headers also counts the guard paragraph mark.
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
