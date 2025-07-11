using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.txt.TextMapping
{
    public class FootnotesMapping : DocumentMapping
    {
        public FootnotesMapping(ConversionContext ctx)
            : base(ctx, ctx.TextDoc.FootnotesWriter)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;
            int id = 0;

            _writer.WriteStartElement("w", "footnotes", OpenXmlNamespaces.WordprocessingML);

            int cp = doc.FIB.ccpText;
            while (cp < doc.FIB.ccpText + doc.FIB.ccpFtn - 2)
            {
                _writer.WriteStartElement("w", "footnote", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, id.ToString());
                cp = writeParagraph(cp);
                _writer.WriteEndElement();
                id++;
            }

            _writer.WriteEndElement();

            _writer.Flush();
        }
    }


}
