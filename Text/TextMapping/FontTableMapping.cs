using System;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.txt.TextMapping
{
    public class FontTableMapping : AbstractMapping,
        IMapping<StringTable>
    {
        protected enum FontFamily
        {
            auto,
            decorative,
            modern,
            roman,
            script,
            swiss
        }

        public FontTableMapping(ConversionContext ctx, IWriter writer)
            : base(writer)
        {
        }

        public void Apply(StringTable table)
        {
            _writer.WriteStartElement("w", "fonts", OpenXmlNamespaces.WordprocessingML);

            foreach (FontFamilyName font in table.Data)
            {
                _writer.WriteStartElement("w", "font", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "name", OpenXmlNamespaces.WordprocessingML, font.xszFtn);

                //alternative name
                if (font.xszAlt!= null && font.xszAlt.Length > 0)
                {
                    _writer.WriteStartElement("w", "altName", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, font.xszAlt);
                    _writer.WriteEndElement();
                }

                //charset
                _writer.WriteStartElement("w", "charset", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x2}", font.chs));
                _writer.WriteEndElement();

                //font family
                _writer.WriteStartElement("w", "family", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((FontFamily)font.ff).ToString());
                _writer.WriteEndElement();

                //panose
                _writer.WriteStartElement("w", "panose1", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteStartAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                foreach (byte b in font.panose)
                {
                    _writer.WriteString(string.Format("{0:x2}", b));
                }
                _writer.WriteEndAttribute();
                _writer.WriteEndElement();

                //pitch
                _writer.WriteStartElement("w", "pitch", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, font.prq.ToString());
                _writer.WriteEndElement();

                //truetype
                if (!font.fTrueType)
                {
                    _writer.WriteStartElement("w", "notTrueType", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "true");
                    _writer.WriteEndElement();
                }

                //font signature
                _writer.WriteStartElement("w", "sig", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "usb0", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield0));
                _writer.WriteAttributeString("w", "usb1", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield1));
                _writer.WriteAttributeString("w", "usb2", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield2));
                _writer.WriteAttributeString("w", "usb3", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.UnicodeSubsetBitfield3));
                _writer.WriteAttributeString("w", "csb0", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.CodePageBitfield0));
                _writer.WriteAttributeString("w", "csb1", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x8}", font.fs.CodePageBitfield1));
                _writer.WriteEndElement();

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();

            _writer.Flush();
        }
    }
}
