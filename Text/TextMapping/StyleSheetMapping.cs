using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.txt.TextMapping
{
    public class StyleSheetMapping 
        : AbstractMapping,
          IMapping<StyleSheet>
    {
        private ConversionContext _ctx;
        private WordDocument _parentDoc;

        public StyleSheetMapping(ConversionContext ctx, WordDocument parentDoc, IWriter writer)
            : base(writer)
        {
            _ctx = ctx;
            _parentDoc = parentDoc;
        }

        public void Apply(StyleSheet sheet)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "styles", OpenXmlNamespaces.WordprocessingML);

            //document defaults
            _writer.WriteStartElement("w", "docDefaults", OpenXmlNamespaces.WordprocessingML);
            writeRunDefaults(sheet);
            writeParagraphDefaults(sheet);
            _writer.WriteEndElement();

            //write the default styles
            if (sheet.Styles[11] == null)
            {
                //NormalTable
                writeNormalTableStyle();
            }

            foreach (var style in sheet.Styles)
            {
                if (style != null)
                {
                    _writer.WriteStartElement("w", "style", OpenXmlNamespaces.WordprocessingML);

                    _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, style.stk.ToString());

                    if (style.sti != StyleSheetDescription.StyleIdentifier.Null && style.sti != StyleSheetDescription.StyleIdentifier.User)
                    {
                        //it's a default style
                        _writer.WriteAttributeString("w", "default", OpenXmlNamespaces.WordprocessingML, "1");
                    }

                    _writer.WriteAttributeString("w", "styleId", OpenXmlNamespaces.WordprocessingML, MakeStyleId(style));

                    // <w:name val="" />
                    _writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, getStyleName(style));
                    _writer.WriteEndElement();

                    // <w:basedOn val="" />
                    if (style.istdBase != 4095 && style.istdBase < sheet.Styles.Count)
                    {
                        _writer.WriteStartElement("w", "basedOn", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, MakeStyleId(sheet.Styles[(int)style.istdBase]));
                        _writer.WriteEndElement();
                    }

                    // <w:next val="" />
                    if (style.istdNext < sheet.Styles.Count)
                    {
                        _writer.WriteStartElement("w", "next", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, MakeStyleId(sheet.Styles[(int)style.istdNext]));
                        _writer.WriteEndElement();
                    }

                    // <w:link val="" />
                    if (style.istdLink < sheet.Styles.Count)
                    {
                        _writer.WriteStartElement("w", "link", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, MakeStyleId(sheet.Styles[(int)style.istdLink]));
                        _writer.WriteEndElement();
                    }

                    // <w:locked/>
                    if (style.fLocked)
                    {
                        _writer.WriteElementString("w", "locked", OpenXmlNamespaces.WordprocessingML, null);
                    }

                    // <w:hidden/>
                    if (style.fHidden)
                    {
                        _writer.WriteElementString("w", "hidden", OpenXmlNamespaces.WordprocessingML, null);
                    }

                    // <w:semiHidden/>
                    if (style.fSemiHidden)
                    {
                        _writer.WriteElementString("w", "semiHidden", OpenXmlNamespaces.WordprocessingML, null);
                    }

                    //write paragraph properties
                    if (style.papx != null)
                    {
                        style.papx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, _parentDoc, null));
                    }
                    
                    //write character properties
                    if (style.chpx != null)
                    {
                        var rev = new RevisionData();
                        rev.Type = RevisionData.RevisionType.NoRevision;
                        style.chpx.Convert(new CharacterPropertiesMapping(_writer, _parentDoc, rev, style.papx, true));
                    }

                    //write table properties
                    if (style.tapx != null)
                    {
                        style.tapx.Convert(new TablePropertiesMapping(_writer, sheet, new List<short>()));
                    }

                    _writer.WriteEndElement();
                }
            }

            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        private void writeRunDefaults(StyleSheet sheet)
        {
            _writer.WriteStartElement("w", "rPrDefault", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "rPr", OpenXmlNamespaces.WordprocessingML);

            //write default fonts
            _writer.WriteStartElement("w", "rFonts", OpenXmlNamespaces.WordprocessingML);

            var ffnAscii = (FontFamilyName)_ctx.Doc.FontTable.Data[sheet.stshi.rgftcStandardChpStsh[0]];
            _writer.WriteAttributeString("w", "ascii", OpenXmlNamespaces.WordprocessingML, ffnAscii.xszFtn);

            var ffnAsia = (FontFamilyName)_ctx.Doc.FontTable.Data[sheet.stshi.rgftcStandardChpStsh[1]];
            _writer.WriteAttributeString("w", "eastAsia", OpenXmlNamespaces.WordprocessingML, ffnAsia.xszFtn);

            var ffnAnsi = (FontFamilyName)_ctx.Doc.FontTable.Data[sheet.stshi.rgftcStandardChpStsh[2]];
            _writer.WriteAttributeString("w", "hAnsi", OpenXmlNamespaces.WordprocessingML, ffnAsia.xszFtn);

            var ffnComplex = (FontFamilyName)_ctx.Doc.FontTable.Data[sheet.stshi.rgftcStandardChpStsh[3]];
            _writer.WriteAttributeString("w", "cs", OpenXmlNamespaces.WordprocessingML, ffnComplex.xszFtn);

            _writer.WriteEndElement();

            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }

        private void writeParagraphDefaults(StyleSheet sheet)
        {
            //if there is no pPrDefault, Word will not used the default paragraph settings.
            //writing an empty pPrDefault will cause Word to load the default paragraph settings.
            _writer.WriteStartElement("w", "pPrDefault", OpenXmlNamespaces.WordprocessingML);

            _writer.WriteEndElement();
        }

        /// <summary>
        /// Generates a style id for custom style names or returns the build-in identifier for build-in styles.
        /// </summary>
        /// <param name="std">The StyleSheetDescription</param>
        /// <returns></returns>
        public static string MakeStyleId(StyleSheetDescription std)
        {
            if (std is null) return "";
            if (std.sti != StyleSheetDescription.StyleIdentifier.User &&
                std.sti != StyleSheetDescription.StyleIdentifier.Null)
            {
                //use the identifier
                return std.sti.ToString();
            }
            else
            {
                //if no identifier is set, use the name
                string ret = std.xstzName;
                ret = ret.Replace(" ", "");
                ret = ret.Replace("(", "");
                ret = ret.Replace(")", "");
                ret = ret.Replace("'", "");
                ret = ret.Replace("\"", "");
                return ret;
            }
        }

        /// <summary>
        /// Chooses the correct style name.
        /// Word 2007 needs the identifier instead of the stylename for translating it into the UI language.
        /// </summary>
        /// <param name="std">The StyleSheetDescription</param>
        /// <returns></returns>
        private string getStyleName(StyleSheetDescription std)
        {
            if (std.sti != StyleSheetDescription.StyleIdentifier.User &&
                std.sti != StyleSheetDescription.StyleIdentifier.Null)
            {
                //use the identifier
                return std.sti.ToString();
            }
            else
            {
                //if no identifier is set, use the name
                return std.xstzName;
            }
        }

        /// <summary>
        /// Writes the "NormalTable" default style
        /// </summary>
        private void writeNormalTableStyle()
        {
            _writer.WriteStartElement("w", "style", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "table");
            _writer.WriteAttributeString("w", "default", OpenXmlNamespaces.WordprocessingML, "1");
            _writer.WriteAttributeString("w", "styleId", OpenXmlNamespaces.WordprocessingML, "TableNormal");
            _writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "Normal Table");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "uiPriority", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "99");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "semiHidden", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "unhideWhenUsed", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "qFormat", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "tblPr", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "tblInd", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "w", OpenXmlNamespaces.WordprocessingML, "0");
            _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "dxa");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "tblCellMar", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "top", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "w", OpenXmlNamespaces.WordprocessingML, "0");
            _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "dxa");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "left", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "w", OpenXmlNamespaces.WordprocessingML, "108");
            _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "dxa");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "bottom", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "w", OpenXmlNamespaces.WordprocessingML, "0");
            _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "dxa");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "right", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "w", OpenXmlNamespaces.WordprocessingML, "108");
            _writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "dxa");
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }
    }
}
