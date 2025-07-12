using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace b2xtranslator.txt
{
    public class TextWriter : IWriter
    {
        /*
           <w:p>
              <w:r>
                <w:t>Hello world</w:t>
              </w:r>
        </w:p>
         */

        [DebuggerDisplay("{Prefix}:{LocalName}={ToString()}")]
        class TextElement
        {
            public string? Prefix { get; private set; }
            public string LocalName { get; private set; }
            public readonly StringBuilder Content = new StringBuilder();
            public StringBuilder PureContent = new StringBuilder();
            public TextElement? Parent { get; private set; }

            public List<INode> Nodes { get; set; }
            public List<IAttribute> Attributes { get; set; }
            public TextElement(TextElement? parent, string? prefix, string localName, string? value)
            {
                Parent = parent;
                Prefix = prefix;
                LocalName = localName;
                if (value != null)
                {
                    this.Content.Append(value);
                }
            }

            public void Append(string value)
            {
                if (value != null)
                {
                    Content.Append(value);
                }
            }
            public void Append(char[] chars, int index, int count)
            {
                Content.Append(chars, index, count);
            }

            public void Append(char ch)
            {
                Content.Append(ch);
            }
            public override string ToString()
            {
                return Content.ToString();
            }
        }

        //private readonly StringBuilder _mainSb = new StringBuilder();
        private readonly TextElement _rootTextElement;
        private TextElement _currentTextElement;
        private readonly Stack<TextElement> _elementStack;

        public TextWriter()
        {
            _rootTextElement = new TextElement(null, null, "root", null);
            _currentTextElement = _rootTextElement;
            _elementStack = new Stack<TextElement>();
        }

        public void Flush()
        {
            // No-op for StringBuilder
        }

        public void WriteAttributeString(string? prefix, string localName, string? ns, string? value)
        {
            _currentTextElement.Attributes ??= new List<IAttribute>();
            _currentTextElement.Attributes.Add(new TextAttribute(prefix,localName, value));
        }

        public void WriteChars(char[] chars, int index, int count)
        {
            _currentTextElement.Append(chars, index, count);
        }

        public void WriteChar(char c)
        {
            _currentTextElement.Append(c);
        }

        public void WriteElementString(string? prefix, string localName, string? ns, string? value)
        {

            if ("w".Equals(prefix))
            {
                if ("tab".Equals(localName))
                {
                    _currentTextElement.PureContent.Append("\t");
                }
                else if ("br".Equals(localName))
                {
                    _currentTextElement.PureContent.Append("\n");
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
                var element = _elementStack.Pop();

                // Restaura o elemento pai como atual
                _currentTextElement = element.Parent ?? _rootTextElement;

                if (element.Content.ToString().Contains("List Number 1 (Level 3)"))
                {

                }
                Debug.Print($"Closing {element.Prefix}:{element.LocalName}");
                Debug.Print($"Content: {element.ToString()}");

                // Para outros elementos estruturais, o conte�do j� foi propagado
                // dos w:t filhos, ent�o n�o precisamos fazer nada

                // Adiciona separadores ap�s o conte�do, baseado no tipo de elemento
                if ("w".Equals(element.Prefix))
                {
                    if ("tc".Equals(element.LocalName))  // Table cell
                    {
                        _currentTextElement.PureContent.Append("\t");
                    }
                    else if ("tr".Equals(element.LocalName))  // Table row
                    {
                        _currentTextElement.PureContent.Append("\n"); // do not use NewLine
                    }
                    else if ("p".Equals(element.LocalName))  // Paragraph
                    {
                        if (!"tc".Equals(element.Parent?.LocalName))
                        {
                            _currentTextElement.PureContent.Append("\n"); // do not use NewLine
                        }
                    }
                }

                _currentTextElement.PureContent.Append(element.PureContent);

                // Propaga conte�do APENAS de elementos w:t
                if ("w".Equals(element.Prefix) && "t".Equals(element.LocalName))
                {
                    _currentTextElement.PureContent.Append(element.Content);
                }
            }
        }


        public void WriteNode(INode node)
        {
            if (node != null)
            {
                _currentTextElement.Nodes ??= new List<INode>();
                _currentTextElement.Nodes.Add(node);

            }
        }

        public void WriteStartAttribute(string? prefix, string localName, string? ns, string? value)
        {
            // Ignore for plain text
        }

        public void WriteStartDocument()
        {
            // Ignore for plain text
        }

        public void WriteStartElement(string? prefix, string localName, string? ns, string? value)
        {
            //if (("w".Equals(prefix) && "t".Equals(localName)) ||
            //         localName == "tc" ||
            //        localName == "tr")
            {

                _currentTextElement = new TextElement(_currentTextElement, prefix, localName, value);
                _elementStack.Push(_currentTextElement);
            }
        }

        public void WriteString(string v)
        {
            _currentTextElement.Append(v);
        }

        public override string ToString()
        {
            // Unpop all elements from the stack and append their content to the root
            while (_elementStack.Count > 0)
            {
                WriteEndElement();
            }
            return _rootTextElement.PureContent.ToString();
        }

    }
}
