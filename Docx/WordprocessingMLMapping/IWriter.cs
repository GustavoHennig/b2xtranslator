using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.doc.WordprocessingMLMapping
{
    public interface IWriter
    {
        void WriteAttributeString(string v1, string v2, string wordprocessingML, string rsid);
        void WriteElementString(string v1, string v2, string wordprocessingML, string v3);
        void WriteEndElement();
        void WriteStartElement(string v1, string v2, string wordprocessingML);
    }
}
