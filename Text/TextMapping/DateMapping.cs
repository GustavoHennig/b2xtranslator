using System;
using System.Text;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.txt.TextMapping
{
    public class DateMapping : AbstractMapping,
          IMapping<DateAndTime>
    {
        INode _parent;

        /// <summary>
        /// Writes a date attribute to the given writer
        /// </summary>
        /// <param name="writer"></param>
        public DateMapping(IWriter writer)
            : base(writer)
        {
        }

        /// <summary>
        /// Appends a date attribute to the given Element
        /// </summary>
        /// <param name="parent"></param>
        public DateMapping(INode parent)
            : base(null)
        {
            _parent = parent;
            _nodeFactory = parent.OwnerDocument;
        }

        public void Apply(DateAndTime dttm)
        {
            var date = new StringBuilder();
            date.Append(string.Format("{0:0000}", dttm.yr));
            date.Append("-");
            date.Append(string.Format("{0:00}", dttm.mon));
            date.Append("-");
            date.Append(string.Format("{0:00}", dttm.dom));
            date.Append("T");
            date.Append(string.Format("{0:00}", dttm.hr));
            date.Append(":");
            date.Append(string.Format("{0:00}", dttm.mint));
            date.Append(":00Z");

            var xml = _nodeFactory.CreateAttribute("w", "date", OpenXmlNamespaces.WordprocessingML);
            xml.Value = date.ToString() ;

            //append or write
            if (_writer != null)
            {
                xml.WriteTo(_writer);
            }
            else if (_parent != null)
            {
                _parent.Attributes.Append(xml);
            }
        }
    }
}
