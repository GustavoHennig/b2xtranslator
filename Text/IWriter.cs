using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt
{
    /// <summary>
    /// This interface was created as a proof of concept (PoC) to abstract content writing,
    /// decoupling it from a specific XML implementation.
    /// In the future, it could be extended to support other output formats as well.
    /// </summary>
    public interface IWriter
    {
        void Flush();
        void WriteAttributeString(string? prefix, string localName, string? ns = null, string? value = null);
        void WriteChar(char c);
        void WriteChars(char[] chars, int index, int count);
        void WriteElementString(string? prefix, string localName, string? ns = null, string? value = null);
        void WriteEndAttribute();
        void WriteEndDocument();
        void WriteEndElement();
        void WriteNode(INode node);
        void WriteStartAttribute(string? prefix, string localName, string? ns = null, string? value = null);
        void WriteStartDocument();
        void WriteStartElement(string prefix, string? localName = null, string? ns = null, string? value = null);
        void WriteString(string value);
    }
}
