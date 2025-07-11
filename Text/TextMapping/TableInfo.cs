using b2xtranslator.DocFileFormat;
using b2xtranslator.Tools;

namespace b2xtranslator.txt.TextMapping
{
    public class TableInfo
    {
        /// <summary>
        /// 
        /// </summary>
        public bool fInTable;

        /// <summary>
        /// 
        /// </summary>
        public bool fTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTableCell;

        /// <summary>
        /// 
        /// </summary>
        public uint iTap;

        public TableInfo(ParagraphPropertyExceptions papx)
        {
            foreach (var sprm in papx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInTable)
                {
                    fInTable = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFTtp)
                {
                    fTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInnerTableCell)
                {
                    fInnerTableCell = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInnerTtp)
                {
                    fInnerTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPItap)
                {
                    iTap = BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (iTap > 0)
                        fInTable = true;
                }
                if ((int)sprm.OpCode == 0x66A)
                {
                    //add value!
                    iTap = BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (iTap > 0)
                        fInTable = true;
                }
            }
        }
    }
}
