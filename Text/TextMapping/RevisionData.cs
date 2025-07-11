using System;
using System.Collections.Generic;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.txt.TextMapping
{
    public class RevisionData
    {
        public enum RevisionType
        {
            NoRevision,
            Inserted,
            Deleted,
            Changed
        }

        public DateAndTime Dttm;
        public short Isbt;
        public RevisionType Type;
        public List<SinglePropertyModifier> Changes;
        public int RsidDel;
        public int RsidProp;
        public int Rsid;

        public RevisionData()
        {
            Changes = new List<SinglePropertyModifier>();
        }

        /// <summary>
        /// Collects the revision data of a CHPX
        /// </summary>
        /// <param name="chpx"></param>
        public RevisionData(CharacterPropertyExceptions chpx)
        {
            bool collectRevisionData = true;
            Changes = new List<SinglePropertyModifier>();

            foreach (var sprm in chpx.grpprl)
            {
                switch ((int)sprm.OpCode)
                {
                    //revision data
                    case 0xCA89:
                        //revision mark
                        collectRevisionData = false;
                        //author 
                        Isbt = BitConverter.ToInt16(sprm.Arguments, 1);
                        //date
                        var dttmBytes = new byte[4];
                        Array.Copy(sprm.Arguments, 3, dttmBytes, 0, 4);
                        Dttm = new DateAndTime(dttmBytes);
                        break;
                    case 0x0801:
                        //revision mark
                        collectRevisionData = false;
                        break;
                    case 0x4804:
                        //author
                        Isbt = BitConverter.ToInt16(sprm.Arguments, 0);
                        break;
                    case 0x6805:
                        //date
                        Dttm = new DateAndTime(sprm.Arguments);
                        break;
                    case 0x0800:
                        //delete mark
                        Type = RevisionType.Deleted;
                        break;
                    case 0x6815:
                        RsidProp = BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case 0x6816:
                        Rsid = BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case 0x6817:
                        RsidDel = BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                }

                //put the sprm on the revision stack
                if (collectRevisionData)
                {
                    Changes.Add(sprm);
                }
            }

            //type
            if (Type != RevisionType.Deleted)
            {
                if (collectRevisionData)
                {
                    //no mark was found, so this CHPX doesn't contain revision data
                    Type = RevisionType.NoRevision;
                }
                else
                {
                    if (Changes.Count > 0)
                    {
                        Type = RevisionType.Changed;
                    }
                    else
                    {
                        Type = RevisionType.Inserted;
                        Changes.Clear();
                    }
                }
            }
        }
    }
}
