using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using System.Collections;
using b2xtranslator.Tools;
using b2xtranslator.txt.TextModel;

namespace b2xtranslator.txt.TextMapping
{
    public class TableCellPropertiesMapping : 
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private int _gridIndex;
        private int _cellIndex;
        private INode _tcPr;
        private INode _tcMar;
        private INode _tcBorders;
        private List<short> _grid;
        private short[] _tGrid;

        private short _width;
        private Global.CellWidthType _ftsWidth;
        private TC80 _tcDef;

        private BorderCode _brcTop, _brcLeft, _brcRight, _brcBottom;

        /// <summary>
        /// The grind span of this cell
        /// </summary>
        private int _gridSpan;
        public int GridSpan
        {
            get { return _gridSpan; }
        }

        private enum VerticalCellAlignment
        {
            top,
            center,
            bottom
        }

        public TableCellPropertiesMapping(IWriter writer, List<short> tableGrid, int gridIndex, int cellIndex)
            : base(writer)
        {
            _tcPr = _nodeFactory.CreateElement("w", "tcPr", OpenXmlNamespaces.WordprocessingML);
            _tcMar = _nodeFactory.CreateElement("w", "tcMar", OpenXmlNamespaces.WordprocessingML);
            _tcBorders = _nodeFactory.CreateElement("w", "tcBorders", OpenXmlNamespaces.WordprocessingML);
            _gridIndex = gridIndex;
            _grid = tableGrid;
            _cellIndex = cellIndex;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            //int lastBdr = getLastTabelBorderOccurrence(tapx.grpprl);

            for (int i=0; i< tapx.grpprl.Count; i++)
            {
                var sprm = tapx.grpprl[i];

                switch (sprm.OpCode)
	            {
                    //Table definition SPRM
                    case  SinglePropertyModifier.OperationCode.sprmTDefTable:
                        var tdef = new SprmTDefTable(sprm.Arguments);
                        _tGrid = tdef.rgdxaCenter;
                        if (tdef.rgTc80.Length <= _cellIndex) break;
                        _tcDef = tdef.rgTc80[_cellIndex];

                        appendValueElement(_tcPr, "textDirection", _tcDef.textFlow.ToString(), false);

                        if (_tcDef.vertMerge == Global.VerticalMergeFlag.fvmMerge)
                            appendValueElement(_tcPr, "vMerge", "continue", false);
                        else if (_tcDef.vertMerge == Global.VerticalMergeFlag.fvmRestart)
                            appendValueElement(_tcPr, "vMerge", "restart", false);
                        else if (_tcDef.vertMerge == Global.VerticalMergeFlag.fvmRestart)
                            appendValueElement(_tcPr, "vMerge", "restart", false);

                        appendValueElement(_tcPr, "vAlign", _tcDef.vertAlign.ToString(), false);

                        if (_tcDef.fFitText)
                            appendValueElement(_tcPr, "tcFitText", "", false);

                        if (_tcDef.fNoWrap)
                            appendValueElement(_tcPr, "noWrap", "", true);

                        //_width = _tcDef.wWidth;
                        _width = (short)(tdef.rgdxaCenter[_cellIndex + 1] - tdef.rgdxaCenter[_cellIndex]);
                        _ftsWidth = _tcDef.ftsWidth;

                        //borders
                        // if the sprm has a higher priority than the last sprmTTableBorder sprm in the list
                        //if (i > lastBdr)
                        //{
                        //    _brcTop = _tcDef.brcTop;
                        //    _brcLeft = _tcDef.brcLeft;
                        //    _brcRight = _tcDef.brcRight;
                        //    _brcBottom = _tcDef.brcBottom;
                        //}

                        break;

                    //margins
                    case SinglePropertyModifier.OperationCode.sprmTCellPadding:
                        byte first = sprm.Arguments[0];
                        byte lim = sprm.Arguments[1];
                        byte ftsMargin = sprm.Arguments[3];
                        short wMargin = BitConverter.ToInt16(sprm.Arguments, 4);

                        if (_cellIndex >= first && _cellIndex < lim)
                        {
                            var borderBits = new BitArray(new byte[] { sprm.Arguments[2] });
                            if (borderBits[0] == true)
                                appendDxaElement(_tcMar, "top", wMargin.ToString(), true);
                            if (borderBits[1] == true)
                                appendDxaElement(_tcMar, "left", wMargin.ToString(), true);
                            if (borderBits[2] == true)
                                appendDxaElement(_tcMar, "bottom", wMargin.ToString(), true);
                            if (borderBits[3] == true)
                                appendDxaElement(_tcMar, "right", wMargin.ToString(), true);
                        }
                        break;

                    //shading
                    case SinglePropertyModifier.OperationCode.sprmTDefTableShd:
                        //cell shading for cells 0-20
                        apppendCellShading(sprm.Arguments, _cellIndex);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDefTableShd2nd:
                        //cell shading for cells 21-42
                        apppendCellShading(sprm.Arguments, _cellIndex - 21);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDefTableShd3rd:
                        //cell shading for cells 43-62
                        apppendCellShading(sprm.Arguments, _cellIndex - 43);
                        break;

                    //width
                    case SinglePropertyModifier.OperationCode.sprmTCellWidth:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (_cellIndex >= first && _cellIndex < lim)
                        {
                            _ftsWidth = (Global.CellWidthType)sprm.Arguments[2];
                            _width = BitConverter.ToInt16(sprm.Arguments, 3);
                        }
                        break;

                    //vertical alignment
                    case SinglePropertyModifier.OperationCode.sprmTVertAlign:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (_cellIndex >= first && _cellIndex < lim)
                            appendValueElement(_tcPr, "vAlign", ((VerticalCellAlignment)sprm.Arguments[2]).ToString(), true);
                        break;

                    //Autofit
                    case SinglePropertyModifier.OperationCode.sprmTFitText:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (_cellIndex >= first && _cellIndex < lim)
                            appendValueElement(_tcPr, "tcFitText", sprm.Arguments[2].ToString(), true);
                        break;

                    //borders (cell definition)
                    case SinglePropertyModifier.OperationCode.sprmTSetBrc:
                        byte min = sprm.Arguments[0];
                        byte max = sprm.Arguments[1];
                        int bordersToApply = sprm.Arguments[2] ;

                        if (_cellIndex >= min && _cellIndex < max)
                        {
                            var brcBytes = new byte[8];
                            Array.Copy(sprm.Arguments, 3, brcBytes, 0, 8);
                            var border = new BorderCode(brcBytes);
                            if(Utils.BitmaskToBool(bordersToApply, 0x01))
                            {
                                _brcTop = border;
                            }
                            if(Utils.BitmaskToBool(bordersToApply, 0x02))
                            {
                                _brcLeft = border;
                            }
                            if (Utils.BitmaskToBool(bordersToApply, 0x04))
                            {
                                _brcBottom = border;
                            }
                            if (Utils.BitmaskToBool(bordersToApply, 0x08))
                            {
                                _brcRight = border;
                            }
                        }
                        break;
                }
            }

            //width
            var tcW = _nodeFactory.CreateElement("w", "tcW", OpenXmlNamespaces.WordprocessingML);
            var tcWType = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
            var tcWVal = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
            tcWType.Value = _ftsWidth.ToString();
            tcWVal.Value = _width.ToString();
            tcW.Attributes.Append(tcWType);
            tcW.Attributes.Append(tcWVal);
            _tcPr.AppendChild(tcW);

            //grid span
            _gridSpan = 1;
            if (_gridIndex < _grid.Count && _width > _grid[_gridIndex])
            {
                //check the number of merged cells
                int w = _grid[_gridIndex];
                for (int i = _gridIndex + 1; i < _grid.Count; i++)
                {
                    _gridSpan++;
                    w += _grid[i];
                    if (w >= _width)
                        break;
                }

                appendValueElement(_tcPr, "gridSpan", _gridSpan.ToString(), true);
            }

            //append margins
            if (_tcMar.ChildNodes.Count > 0)
            {
                _tcPr.AppendChild(_tcMar);
            }

            //append borders (removed)

            if (_tcBorders.ChildNodes.Count > 0)
            {
                _tcPr.AppendChild(_tcBorders);
            }

            //write Properties
            if (_tcPr.ChildNodes.Count > 0 || _tcPr.Attributes.Count > 0)
            {
                _tcPr.WriteTo(_writer);
            }
        }

        private void apppendCellShading(byte[] sprmArg, int cellIndex)
        {
            //shading descriptor can have 10 bytes (Word 2000) or 2 bytes (Word 97)
            int shdLength = 2;
            if (sprmArg.Length % 10 == 0)
                shdLength = 10;

            var shdBytes = new byte[shdLength];

            //multiple cell can be formatted with the same SHD.
            //in this case there is only 1 SHD for all cells in the row.
            if (cellIndex * shdLength >= sprmArg.Length)
            {
                //use the first SHD
                cellIndex = 0;
            }

            Array.Copy(sprmArg, cellIndex * shdBytes.Length, shdBytes, 0, shdBytes.Length);
            
            var shd = new ShadingDescriptor(shdBytes);
            appendShading(_tcPr, shd);
        }


        /// <summary>
        /// Returns the index of the last occurence of an sprmTTableBorders or sprmTTableBorders80 sprm.
        /// </summary>
        /// <param name="grpprl">The grpprl of sprms</param>
        /// <returns>The index or -1 if no sprm is in the list</returns>
        private int getLastTabelBorderOccurrence(List<SinglePropertyModifier> grpprl)
        {
            int index = -1;

            for (int i = 0; i < grpprl.Count; i++)
            {
                if (grpprl[i].OpCode == SinglePropertyModifier.OperationCode.sprmTTableBorders ||
                    grpprl[i].OpCode == SinglePropertyModifier.OperationCode.sprmTTableBorders80)
                {
                    index = i;
                }
            }

            return index;
        }
    }
}
