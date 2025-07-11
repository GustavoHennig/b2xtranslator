using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.txt.TextMapping
{
    public class EndnotesMapping : DocumentMapping
    {
        public EndnotesMapping(ConversionContext ctx)
            : base(ctx, ctx.TextDoc.EndnotesWriter)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;
            int id = 0;

            _writer.WriteStartElement("w", "endnotes", OpenXmlNamespaces.WordprocessingML);

            int cp = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn;
            int cpEnd = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr + doc.FIB.ccpAtn + doc.FIB.ccpEdn - 2;
            while (cp < cpEnd)
            {
                _writer.WriteStartElement("w", "endnote", OpenXmlNamespaces.WordprocessingML);
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
