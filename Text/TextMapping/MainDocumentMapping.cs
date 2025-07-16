using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

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
            int cp = 0;
            while (cp < doc.FIB.ccpText)
            {
                int fc = _doc.PieceTable.FileCharacterPositions[cp];
                var papx = findValidPapx(fc);
                if(papx == null)
                {
                    // If no valid PAPX is found, skip to the next character position
                    cp++;
                    continue;
                }
                var tai = new TableInfo(papx);

                if (tai.fInTable)
                    //this PAPX is for a table
                    cp = writeTable(cp, tai.iTap);
                else
                    //this PAPX is for a normal paragraph
                    cp = writeParagraph(cp);
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
