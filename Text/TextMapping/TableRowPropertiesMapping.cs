using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Tools;
using System;
using b2xtranslator.txt.TextModel;

namespace b2xtranslator.txt.TextMapping
{
    class TableRowPropertiesMapping :
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private INode _trPr;
        private INode _tblPrEx;
        //private INode _tblBorders;
        //private BorderCode brcLeft, brcTop, brcBottom, brcRight, brcHorz, brcVert;
        private CharacterPropertyExceptions _rowEndChpx;

        public TableRowPropertiesMapping(IWriter writer, CharacterPropertyExceptions rowEndChpx)
            : base(writer)
        {
            _trPr = _nodeFactory.CreateElement("w", "trPr", OpenXmlNamespaces.WordprocessingML);
            _tblPrEx = _nodeFactory.CreateElement("w", "tblPrEx", OpenXmlNamespaces.WordprocessingML);
            //_tblBorders = _nodeFactory.CreateElement("w", "tblBorders", OpenXmlNamespaces.WordprocessingML);
            _rowEndChpx = rowEndChpx;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            //delete infos
            var rev = new RevisionData(_rowEndChpx);
            if (_rowEndChpx != null && rev.Type == RevisionData.RevisionType.Deleted)
            {
                var del = _nodeFactory.CreateElement("w", "del", OpenXmlNamespaces.WordprocessingML);
                _trPr.AppendChild(del);
            }

            foreach (var sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)  
                {
                    case SinglePropertyModifier.OperationCode.sprmTDefTable:
                        //SprmTDefTable tdef = new SprmTDefTable(sprm.Arguments);
                        break;

                    //header row
                    case SinglePropertyModifier.OperationCode.sprmTTableHeader:
                        bool fHeader = Utils.ByteToBool(sprm.Arguments[0]);
                        if(fHeader)
                        {
                            var header = _nodeFactory.CreateElement("w", "tblHeader", OpenXmlNamespaces.WordprocessingML);
                            _trPr.AppendChild(header);
                        }
                        break;

                    //width after
                    case SinglePropertyModifier.OperationCode.sprmTWidthAfter:
                        var wAfter = _nodeFactory.CreateElement("w", "wAfter", OpenXmlNamespaces.WordprocessingML);
                        var wAfterValue = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                        wAfterValue.Value = BitConverter.ToInt16(sprm.Arguments, 1).ToString();
                        wAfter.Attributes.Append(wAfterValue);
                        var wAfterType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                        wAfterType.Value = "dxa";
                        wAfter.Attributes.Append(wAfterType);
                        _trPr.AppendChild(wAfter);
                        break;

                    //width before
                    case SinglePropertyModifier.OperationCode.sprmTWidthBefore:
                        short before = BitConverter.ToInt16(sprm.Arguments, 1);
                        if (before != 0)
                        {
                            var wBefore = _nodeFactory.CreateElement("w", "wBefore", OpenXmlNamespaces.WordprocessingML);
                            var wBeforeValue = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                            wBeforeValue.Value = before.ToString();
                            wBefore.Attributes.Append(wBeforeValue);
                            var wBeforeType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                            wBeforeType.Value = "dxa";
                            wBefore.Attributes.Append(wBeforeType);
                            _trPr.AppendChild(wBefore);
                        }
                        break;

                    //row height
                    case SinglePropertyModifier.OperationCode.sprmTDyaRowHeight:
                        var rowHeight = _nodeFactory.CreateElement("w", "trHeight", OpenXmlNamespaces.WordprocessingML);
                        var rowHeightVal = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                        var rowHeightRule = _nodeFactory.CreateAttribute("w", "hRule", OpenXmlNamespaces.WordprocessingML);
                        short rH = BitConverter.ToInt16(sprm.Arguments, 0);
                        if (rH > 0)
                        {
                            rowHeightRule.Value = "atLeast";
                        }
                        else
                        {
                            rowHeightRule.Value = "exact";
                            rH *= -1;
                        }
                        rowHeightVal.Value = rH.ToString();
                        rowHeight.Attributes.Append(rowHeightVal);
                        rowHeight.Attributes.Append(rowHeightRule);
                        _trPr.AppendChild(rowHeight);
                        break;

                    //can't split
                    case SinglePropertyModifier.OperationCode.sprmTFCantSplit:
                    case SinglePropertyModifier.OperationCode.sprmTFCantSplit90:
                        appendFlagElement(_trPr, sprm, "cantSplit", true);
                        break;

                    //div id
                    case SinglePropertyModifier.OperationCode.sprmTIpgp:
                        appendValueElement(_trPr, "divId", BitConverter.ToInt32(sprm.Arguments, 0).ToString(), true);
                        break;

                    ////borders 80 exceptions
                    //case SinglePropertyModifier.OperationCode.sprmTTableBorders80:
                    //    byte[] brc80 = new byte[4];
                    //    //top border
                    //    Array.Copy(sprm.Arguments, 0, brc80, 0, 4);
                    //    brcTop = new BorderCode(brc80);
                    //    //left
                    //    Array.Copy(sprm.Arguments, 4, brc80, 0, 4);
                    //    brcLeft = new BorderCode(brc80);
                    //    //bottom
                    //    Array.Copy(sprm.Arguments, 8, brc80, 0, 4);
                    //    brcBottom = new BorderCode(brc80);
                    //    //right
                    //    Array.Copy(sprm.Arguments, 12, brc80, 0, 4);
                    //    brcRight = new BorderCode(brc80);
                    //    //inside H
                    //    Array.Copy(sprm.Arguments, 16, brc80, 0, 4);
                    //    brcHorz = new BorderCode(brc80);
                    //    //inside V
                    //    Array.Copy(sprm.Arguments, 20, brc80, 0, 4);
                    //    brcVert = new BorderCode(brc80);
                    //    break;

                    ////border exceptions
                    //case SinglePropertyModifier.OperationCode.sprmTTableBorders:
                    //    byte[] brc = new byte[8];
                    //    //top border
                    //    Array.Copy(sprm.Arguments, 0, brc, 0, 8);
                    //    brcTop = new BorderCode(brc);
                    //    //left
                    //    Array.Copy(sprm.Arguments, 8, brc, 0, 8);
                    //    brcLeft = new BorderCode(brc);
                    //    //bottom
                    //    Array.Copy(sprm.Arguments, 16, brc, 0, 8);
                    //    brcBottom = new BorderCode(brc);
                    //    //right
                    //    Array.Copy(sprm.Arguments, 24, brc, 0, 8);
                    //    brcRight = new BorderCode(brc);
                    //    //inside H
                    //    Array.Copy(sprm.Arguments, 32, brc, 0, 8);
                    //    brcHorz = new BorderCode(brc);
                    //    //inside V
                    //    Array.Copy(sprm.Arguments, 40, brc, 0, 8);
                    //    brcVert = new BorderCode(brc);
                    //    break;
                }
            }

            ////set borders
            //if (brcTop != null)
            //{
            //    XmlNode topBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
            //    appendBorderAttributes(brcTop, topBorder);
            //    addOrSetBorder(_tblBorders, topBorder);
            //}
            //if (brcLeft != null)
            //{
            //    XmlNode leftBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
            //    appendBorderAttributes(brcLeft, leftBorder);
            //    addOrSetBorder(_tblBorders, leftBorder);
            //}
            //if (brcBottom != null)
            //{
            //    XmlNode bottomBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
            //    appendBorderAttributes(brcBottom, bottomBorder);
            //    addOrSetBorder(_tblBorders, bottomBorder);
            //}
            //if (brcRight != null)
            //{
            //    XmlNode rightBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
            //    appendBorderAttributes(brcRight, rightBorder);
            //    addOrSetBorder(_tblBorders, rightBorder);
            //}
            //if (brcHorz != null)
            //{
            //     XmlNode insideHBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideH", OpenXmlNamespaces.WordprocessingML);
            //     appendBorderAttributes(brcHorz, insideHBorder);
            //     addOrSetBorder(_tblBorders, insideHBorder);
            //}
            //if (brcVert != null)
            //{
            //    XmlNode insideVBorder = _nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideV", OpenXmlNamespaces.WordprocessingML);
            //    appendBorderAttributes(brcVert, insideVBorder);
            //    addOrSetBorder(_tblBorders, insideVBorder);
            //}
            //if (_tblBorders.ChildNodes.Count > 0)
            //{
            //    _tblPrEx.AppendChild(_tblBorders);
            //}

            //set exceptions
            if (_tblPrEx.ChildNodes.Count > 0)
            {
                _trPr.AppendChild(_tblPrEx);
            }

            //write Properties
            if (_trPr.ChildNodes.Count > 0 || _trPr.Attributes.Count > 0)
            {
                _trPr.WriteTo(_writer);
            }
        }
    }
}
