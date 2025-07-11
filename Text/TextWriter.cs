using System;
using System.Collections.Generic;
using System.Text;

namespace b2xtranslator.txt
{
    public class TextWriter : IWriter
    {
        private readonly StringBuilder _sb = new StringBuilder();
        private readonly Stack<string> _elementStack = new Stack<string>();

        public void Flush()
        {
            // No-op for StringBuilder
        }

        public void WriteAttributeString(string? prefix, string localName, string? ns, string? value)
        {

            if ("w".Equals(prefix))
            {
                if ("tab".Equals(localName))
                {
                    _sb.Append("\t");
                }
                else if ("br".Equals(localName))
                {
                    _sb.Append(Environment.NewLine);
                }
                else if ("lang".Equals(localName))
                {
                    // Ignore language attribute for plain text export
                }
                else if ("val".Equals(localName) && value != null)
                {
                   // _sb.Append(value);

                }
            }
        }

        public void WriteChars(char[] chars, int index, int count)
        {
            _sb.Append(chars, index, count);
        }

        public void WriteChar(char c)
        {
            _sb.Append(c);
        }

        public void WriteElementString(params string[] values)
        {
            //foreach (var v in values)
            //{
            //    _sb.Append(v);
            //}
        }

        public void WriteEndAttribute()
        {
            // Ignore for plain text
        }

        public void WriteEndDocument()
        {
            // Ignore for plain text
        }

        public void WriteEndElement()
        {
            if (_elementStack.Count > 0)
            {
                string element = _elementStack.Pop();
                
                // Add appropriate separators for table elements
                if (element == "tc")  // Table cell
                {
                    _sb.Append("\t");
                }
                else if (element == "tr")  // Table row
                {
                    _sb.Append(Environment.NewLine);
                }
            }
        }

        public void WriteNode(INode node)
        {
            //if (node != null)
            //{
            //    _sb.Append(node.ToString());
            //}
        }

        public void WriteStartAttribute(params string[] attrs)
        {
            // Ignore for plain text
        }

        public void WriteStartDocument()
        {
            // Ignore for plain text
        }

        public void WriteStartElement(params string[] values)
        {
            // Track element names for proper closing separators
            if (values.Length >= 2)
            {
                string elementName = values[1]; // The local name is usually the second parameter
                _elementStack.Push(elementName);
            }
        }

        public void WriteString(string v)
        {
            _sb.Append(v);
        }

        public override string ToString()
        {
            return _sb.ToString();
        }

    }
}
