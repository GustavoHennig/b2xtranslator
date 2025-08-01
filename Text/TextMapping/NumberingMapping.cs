using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.txt.TextModel;

namespace b2xtranslator.txt.TextMapping
{
    public class NumberingMapping : AbstractMapping,
          IMapping<ListTable>
    {
        private ConversionContext _ctx;
        private WordDocument _parentDoc;

        private enum LevelJustification
        {
            left = 0,
            center,
            right
        }

        public NumberingMapping(ConversionContext ctx, WordDocument parentDoc, IWriter writer)
            : base(writer)
        {
            _ctx = ctx;
            _parentDoc = parentDoc;
        }

        public void Apply(ListTable rglst)
        {
            _writer.WriteStartElement("w", "numbering", OpenXmlNamespaces.WordprocessingML);

            for (int i = 0; i < rglst.Count; i++)
            {
                var lstf = rglst[i];

                //start abstractNum
                _writer.WriteStartElement("w", "abstractNum", OpenXmlNamespaces.WordprocessingML);

                _writer.WriteAttributeString("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML, i.ToString());

                //nsid
                _writer.WriteStartElement("w", "nsid", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, string.Format("{0:X8}", lstf.lsid));
                _writer.WriteEndElement();

                //multiLevelType
                _writer.WriteStartElement("w", "multiLevelType", OpenXmlNamespaces.WordprocessingML);
                if (lstf.fHybrid)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "hybridMultilevel");
                else if (lstf.fSimpleList)
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "singleLevel");
                else
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "multilevel");
                _writer.WriteEndElement();

                //template
                _writer.WriteStartElement("w", "tmpl", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, string.Format("{0:X8}", lstf.tplc));
                _writer.WriteEndElement();

                //writes the levels
                for (int j = 0; j < lstf.rglvl.Length; j++)
                {
                    var lvl = lstf.rglvl[j];

                    _writer.WriteStartElement("w", "lvl", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "ilvl", OpenXmlNamespaces.WordprocessingML, j.ToString());

                    //starts at
                    _writer.WriteStartElement("w", "start", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, lvl.iStartAt.ToString());
                    _writer.WriteEndElement();

                    //number format
                    _writer.WriteStartElement("w", "numFmt", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, GetNumberFormat(lvl.nfc));
                    _writer.WriteEndElement();

                    //suffix
                    _writer.WriteStartElement("w", "suff", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, lvl.ixchFollow.ToString());
                    _writer.WriteEndElement();

                    //style
                    //The style id is used for a reverse reference. 
                    //It can happen that the reference points to the wrong style.
                    short styleIndex = lstf.rgistd[j];
                    if(styleIndex != ListData.ISTD_NIL && styleIndex > 0 && styleIndex < _ctx.Doc.Styles.Styles.Count)
                    {
                        _writer.WriteStartElement("w", "pStyle", OpenXmlNamespaces.WordprocessingML);
                        _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, StyleSheetMapping.MakeStyleId(_ctx.Doc.Styles.Styles[styleIndex]));
                        _writer.WriteEndElement();
                    }

                    //Number level text
                    _writer.WriteStartElement("w", "lvlText", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, getLvlText(lvl.xst));
                    _writer.WriteEndElement();

                    //jc
                    _writer.WriteStartElement("w", "lvlJc", OpenXmlNamespaces.WordprocessingML);
                    _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ((LevelJustification)lvl.jc).ToString());
                    _writer.WriteEndElement();

                    //pPr
                    lvl.grpprlPapx.Convert(new ParagraphPropertiesMapping(_writer, _ctx, _parentDoc,  null));

                    //rPr
                    lvl.grpprlChpx.Convert(new CharacterPropertiesMapping(_writer, _parentDoc, new RevisionData(lvl.grpprlChpx), lvl.grpprlPapx, false));

                    _writer.WriteEndElement();
                }

                //end abstractNum
                _writer.WriteEndElement();
            }

            //write the overrides
            for (int i = 0; i < _ctx.Doc.ListFormatOverrideTable.Count; i++)
            {
                var lfo = _ctx.Doc.ListFormatOverrideTable[i];

                //start num
                _writer.WriteStartElement("w", "num", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "numId", OpenXmlNamespaces.WordprocessingML, (i+1).ToString());

                int index = FindIndexbyId(rglst, lfo.lsid);

                _writer.WriteStartElement("w", "abstractNumId", OpenXmlNamespaces.WordprocessingML);
                _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, index.ToString());

                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
            _writer.Flush();
        }

        public static int FindIndexbyId(List<ListData> list, int id)
        {
            int ret = -1;
            for (int i = 0; i < list.Count; i++)
			{
                if (list[i].lsid == id)
                {
                    ret = i;
                    break;
                }
			}
            return ret;
        }

        /// <summary>
        /// Converts the number text of the binary format to the number text of OOXML.
        /// OOXML uses different placeholders for the numbers.
        /// </summary>
        /// <param name="numberText">The number text of the binary format</param>
        /// <returns></returns>
        private string getLvlText(string numberText)
        {
            string ret = numberText;

            ret = ret.Replace(new string((char)0x0000, 1), "%1");
            ret = ret.Replace(new string((char)0x0001, 1), "%2");
            ret = ret.Replace(new string((char)0x0002, 1), "%3");
            ret = ret.Replace(new string((char)0x0003, 1), "%4");
            ret = ret.Replace(new string((char)0x0004, 1), "%5");
            ret = ret.Replace(new string((char)0x0005, 1), "%6");
            ret = ret.Replace(new string((char)0x0006, 1), "%7");
            ret = ret.Replace(new string((char)0x0007, 1), "%8");
            ret = ret.Replace(new string((char)0x0008, 1), "%9");

            return ret;
        }

        /// <summary>
        /// Converts the number format code of the binary format.
        /// </summary>
        /// <param name="nfc">The number format code</param>
        /// <returns>The OOXML attribute value</returns>
        public static string GetNumberFormat(int nfc)
        {
            switch (nfc)
            {
                case 0:
                    return "decimal";
                case 1:
                    return "upperRoman";
                case 2:
                    return "lowerRoman";
                case 3:
                    return "upperLetter";
                case 4:
                    return "lowerLetter";
                case 5:
                    return "ordinal";
                case 6:
                    return "cardinalText";
                case 7:
                    return "ordinalText";
                case 8:
                    return "hex";
                case 9:
                    return "chicago";
                case 10:
                    return "ideographDigital";
                case 11:
                    return "japaneseCounting";
                case 12:
                    return "aiueo";
                case 13:
                    return "iroha";
                case 14:
                    return "decimalFullWidth";
                case 15:
                    return "decimalHalfWidth";
                case 16:
                    return "japaneseLegal";
                case 17:
                    return "japaneseDigitalTenThousand";
                case 18:
                    return "decimalEnclosedCircle";
                case 19:
                    return "decimalFullWidth2";
                case 20:
                    return "aiueoFullWidth";
                case 21:
                    return "irohaFullWidth";
                case 22:
                    return "decimalZero";
                case 23:
                    return "bullet";
                case 24:
                    return "ganada";
                case 25:
                    return "chosung";
                case 26:
                    return "decimalEnclosedFullstop";
                case 27:
                    return "decimalEnclosedParen";
                case 28:
                    return "decimalEnclosedCircleChinese";
                case 29:
                    return "ideographEnclosedCircle";
                case 30:
                    return "ideographTraditional";
                case 31:
                    return "ideographZodiac";
                case 32:
                    return "ideographZodiacTraditional";
                case 33:
                    return "taiwaneseCounting";
                case 34:
                    return "ideographLegalTraditional";
                case 35:
                    return "taiwaneseCountingThousand";
                case 36:
                    return "taiwaneseDigital";
                case 37:
                    return "chineseCounting";
                case 38:
                    return "chineseLegalSimplified";
                case 39:
                    return "chineseCountingThousand";
                case 40:
                    return "koreanDigital";
                case 41:
                    return "koreanCounting";
                case 42:
                    return "koreanLegal";
                case 43:
                    return "koreanDigital2";
                case 44:
                    return "vietnameseCounting";
                case 45:
                    return "russianLower";
                case 46:
                    return "russianUpper";
                case 47:
                    return "none";
                case 48:
                    return "numberInDash";
                case 49:
                    return "hebrew1";
                case 50:
                    return "hebrew2";
                case 51:
                    return "arabicAlpha";
                case 52:
                    return "arabicAbjad";
                case 53:
                    return "hindiVowels";
                case 54:
                    return "hindiConsonants";
                case 55:
                    return "hindiNumbers";
                case 56:
                    return "hindiCounting";
                case 57:
                    return "thaiLetters";
                case 58:
                    return "thaiNumbers";
                case 59:
                    return "thaiCounting";
                default:
                    return "decimal";
            }
        }
    }
}
