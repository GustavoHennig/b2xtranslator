using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using b2xtranslator.Tools;

namespace b2xtranslator.txt.TextModel
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

        // Hyperlink state tracking
        private string? _pendingHyperlinkUrl = null;
        private StringBuilder _hyperlinkDescription = new StringBuilder();
        private bool _isInHyperlinkDescription = false;
        private int _hyperlinkFieldCharCount = 0;

        // Track if we've output any content to prevent leading newline
        private bool _isFirstStructuralElement = true;

        // Flag to control URL extraction
        private bool _extractUrls = true;

        // Symbol handling state
        private bool _isInSymbolElement = false;
        private string? _symbolFont = null;
        private string? _symbolChar = null;

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
                    Content.Append(value);
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

        public TextWriter(bool extractUrls = true)
        {
            _rootTextElement = new TextElement(null, null, "root", null);
            _currentTextElement = _rootTextElement;
            _elementStack = new Stack<TextElement>();
            _extractUrls = extractUrls;
        }

        public void Flush()
        {
            // No-op for StringBuilder
        }

        public void WriteAttributeString(string? prefix, string localName, string? ns, string? value)
        {
            _currentTextElement.Attributes ??= new List<IAttribute>();
            _currentTextElement.Attributes.Add(new TextAttribute(prefix, localName, value));

            // Track symbol attributes
            if (_isInSymbolElement && "w".Equals(prefix))
            {
                if ("font".Equals(localName))
                {
                    _symbolFont = value;
                }
                else if ("char".Equals(localName))
                {
                    _symbolChar = value;
                }
            }
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

                _currentTextElement = element.Parent ?? _rootTextElement;

                if (element.PureContent.ToString().Contains("Apache Tika"))
                {
                }

                if ("w".Equals(element.Prefix))
                {
                    if ("tc".Equals(element.LocalName))  // Table cell
                    {
                        _currentTextElement.PureContent.Append("\t");
                    }
                    else if ("tr".Equals(element.LocalName))  // Table row
                    {
                        if (!_isFirstStructuralElement)
                        {
                            _currentTextElement.PureContent.Append("\n"); // do not use NewLine
                        }
                        _isFirstStructuralElement = false;
                    }
                    else if ("p".Equals(element.LocalName))  // Paragraph
                    {
                        if (!"tc".Equals(element.Parent?.LocalName))
                        {
                            if (!_isFirstStructuralElement)
                            {
                                _currentTextElement.PureContent.Append("\n"); // do not use NewLine
                            }
                            _isFirstStructuralElement = false;
                        }
                    }
                    else if ("instrText".Equals(element.LocalName))
                    {
                        string content = element.Content.ToString();
                        string trimmedContent = content.Trim();

                        if (trimmedContent.StartsWith("HYPERLINK "))
                        {
                            // Extract URL from field instruction - handle both quoted and unquoted formats
                            string url;
                            if (trimmedContent.StartsWith("HYPERLINK \""))
                            {
                                // Use a regular expression to find the first quoted string, which is the URL.
                                var match = System.Text.RegularExpressions.Regex.Match(trimmedContent, @"""([^""]+)""");
                                if (match.Success)
                                {
                                    url = match.Groups[1].Value;
                                }
                                else
                                {
                                    // Fallback to the old logic if the regex fails for some reason.
                                    url = trimmedContent.Replace("HYPERLINK \"", "").Replace("\"", "").Trim();
                                }
                            }
                            else
                            {
                                // Unquoted format: HYPERLINK http://example.com
                                url = trimmedContent.Replace("HYPERLINK ", "").Trim();
                            }
                            _pendingHyperlinkUrl = url;
                            _hyperlinkDescription.Clear();
                            _hyperlinkFieldCharCount = 0;
                            // Don't start collecting description yet - wait for field separator or next instrText
                        }
                        else if (_pendingHyperlinkUrl != null && !string.IsNullOrEmpty(trimmedContent))
                        {
                            // Collect hyperlink description text from subsequent instrText elements
                            _hyperlinkDescription.Append(content); // Use original content to preserve spaces
                            _isInHyperlinkDescription = true;
                        }
                        else if (_pendingHyperlinkUrl != null && string.IsNullOrWhiteSpace(trimmedContent) && !string.IsNullOrEmpty(content))
                        {
                            // Collect whitespace characters (like spaces) in hyperlink descriptions
                            _hyperlinkDescription.Append(content);
                            _isInHyperlinkDescription = true;
                        }
                        // Note: All other instrText elements (field instructions) are skipped for text output
                        // Skip truly empty instrText elements during hyperlink processing
                    }
                    else if ("fldChar".Equals(element.LocalName))
                    {
                        // Handle field character markers
                        if (_pendingHyperlinkUrl != null)
                        {
                            _hyperlinkFieldCharCount++;

                            // Output hyperlink after we've seen enough field characters
                            // Some hyperlinks have 1 fldChar, others have 3 (begin, separator, end)
                            if (_hyperlinkFieldCharCount >= 1 && _hyperlinkDescription.Length > 0)
                            {
                                // We have both URL and description - output complete hyperlink
                                OutputHyperlink();
                            }
                            else if (_hyperlinkFieldCharCount >= 3)
                            {
                                // Three field chars means we're definitely at the end
                                OutputHyperlink();
                            }
                        }
                    }
                    else if ("sym".Equals(element.LocalName))
                    {
                        // Handle symbol conversion
                        try
                        {
                            if (_symbolFont != null && _symbolChar != null)
                            {
                                string symbolText = SymbolMapping.ConvertSymbolHex(_symbolChar, _symbolFont);
                                _currentTextElement.PureContent.Append(symbolText);
                            }
                            else
                            {
                                _currentTextElement.PureContent.Append("?");
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error processing symbol: Font={_symbolFont}, Char={_symbolChar}, Error={ex.Message}");
                            _currentTextElement.PureContent.Append("?");
                        }

                        // Reset symbol state
                        _isInSymbolElement = false;
                        _symbolFont = null;
                        _symbolChar = null;
                    }
                }

                _currentTextElement.PureContent.Append(element.PureContent);

                // Propaga conte�do APENAS de elementos w:t
                if ("w".Equals(element.Prefix) && "t".Equals(element.LocalName))
                {
                    string textContent = element.Content.ToString();

                    // If we're collecting hyperlink description, capture it
                    if (_isInHyperlinkDescription && _pendingHyperlinkUrl != null)
                    {
                        _hyperlinkDescription.Append(textContent);
                    }
                    else
                    {
                        // Normal text processing
                        _currentTextElement.PureContent.Append(textContent);
                    }
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

                // Track symbol elements
                if ("w".Equals(prefix) && "sym".Equals(localName))
                {
                    _isInSymbolElement = true;
                    _symbolFont = null;
                    _symbolChar = null;
                }
            }
        }

        public void WriteString(string v)
        {
            _currentTextElement.Append(v);
        }

        public override string ToString()
        {
            // Ensure all elements are closed and their content is appended to the root element
            while (_elementStack.Count > 0)
            {
                WriteEndElement();
            }

            return _rootTextElement.PureContent.ToString();
        }

        /// <summary>
        /// Helper method to get attribute value from an element
        /// </summary>
        private string? GetAttributeValue(TextElement element, string attributeName)
        {
            if (element.Attributes == null) return null;

            foreach (var attr in element.Attributes)
            {
                if (attributeName.Equals(attr.LocalName))
                {
                    return attr.Value;
                }
            }
            return null;
        }

        /// <summary>
        /// Output the complete hyperlink with both URL and description
        /// </summary>
        private void OutputHyperlink()
        {
            if (_pendingHyperlinkUrl == null) return;

            string description = _hyperlinkDescription.ToString().Trim();

            // Format the hyperlink output based on _extractUrls flag
            if (_extractUrls)
            {
                if (!string.IsNullOrEmpty(description) && !description.Equals(_pendingHyperlinkUrl, StringComparison.OrdinalIgnoreCase))
                {
                    // If we have description text that's different from URL, show: "description (url)"
                    _currentTextElement.PureContent.Append($"{description} ({_pendingHyperlinkUrl})");
                }
                else
                {
                    // If no description or description is the same as URL, just show the URL
                    _currentTextElement.PureContent.Append(_pendingHyperlinkUrl);
                }
            }
            else
            {
                // If URL extraction is disabled, just show the description text
                if (!string.IsNullOrEmpty(description))
                {
                    _currentTextElement.PureContent.Append(description);
                }
            }

            // Reset hyperlink state
            _pendingHyperlinkUrl = null;
            _hyperlinkDescription.Clear();
            _isInHyperlinkDescription = false;
            _hyperlinkFieldCharCount = 0;
        }

    }
}
