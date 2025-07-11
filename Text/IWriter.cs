using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt
{
    public interface IWriter
    {
        void Flush();
        void WriteAttributeString(string? prefix, string localName, string? ns = null, string? value = null);
        void WriteChar(char c);
        void WriteChars(char[] chars, int v1, int v2);
        void WriteElementString(params string[] values);
        void WriteEndAttribute();
        void WriteEndDocument();
        void WriteEndElement();
        void WriteNode(INode node);
        void WriteStartAttribute(params string[] attrs);
        void WriteStartDocument();
        void WriteStartElement(params string[] values);
        void WriteString(string v);
    }
}
