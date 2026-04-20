using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private bool _isInsideField = false;
        private StringBuilder _currentFieldInstruction = new StringBuilder();

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

            // Track symbol font and char attributes for symbol conversion
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
                    // Language attribute is ignored for plain text export
                }
                else if ("val".Equals(localName) && value != null)
                {
                    // Value attribute is ignored for plain text export
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
                        if (_isInsideField)
                        {
                            _currentFieldInstruction.Append(content);
                        }
                    }
                    else if ("fldChar".Equals(element.LocalName))
                    {
                        string? fieldCharType = GetAttributeValue(element, "fldCharType");
                        if ("begin".Equals(fieldCharType, StringComparison.OrdinalIgnoreCase))
                        {
                            _isInsideField = true;
                            _currentFieldInstruction.Clear();
                            _isInHyperlinkDescription = false;
                        }
                        else if ("separate".Equals(fieldCharType, StringComparison.OrdinalIgnoreCase))
                        {
                            BeginFieldResult();
                        }
                        else if ("end".Equals(fieldCharType, StringComparison.OrdinalIgnoreCase))
                        {
                            if (_pendingHyperlinkUrl != null)
                            {
                                OutputHyperlink();
                            }

                            _isInsideField = false;
                            _currentFieldInstruction.Clear();
                            _isInHyperlinkDescription = false;
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
                            TraceLogger.Warning("Error processing symbol: Font={0}, Char={1}, Error={2}", _symbolFont, _symbolChar, ex.Message);
                            _currentTextElement.PureContent.Append("?");
                        }

                        // Reset symbol state
                        _isInSymbolElement = false;
                        _symbolFont = null;
                        _symbolChar = null;
                    }
                }

                _currentTextElement.PureContent.Append(element.PureContent);

                // Propagate content ONLY from w:t elements
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

        }

        public void WriteStartDocument()
        {
            // Ignore for plain text
        }

        public void WriteStartElement(string? prefix, string localName, string? ns, string? value)
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
        /// Helper method to get attribute value from an element.
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
        /// Output the complete hyperlink with both URL and description.
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
        }

        private void BeginFieldResult()
        {
            string instruction = _currentFieldInstruction.ToString().Trim();
            _currentFieldInstruction.Clear();

            if (instruction.StartsWith("HYPERLINK ", StringComparison.OrdinalIgnoreCase))
            {
                _pendingHyperlinkUrl = ExtractHyperlinkUrl(instruction);
                _hyperlinkDescription.Clear();
                _isInHyperlinkDescription = _pendingHyperlinkUrl != null;
            }
            else
            {
                _isInHyperlinkDescription = false;
            }
        }

        private static string? ExtractHyperlinkUrl(string instruction)
        {
            instruction = SanitizeFieldInstruction(instruction);

            var bookmarkMatch = Regex.Match(
                instruction,
                @"^HYPERLINK\s+\\l\s+""([^""]+)""",
                RegexOptions.IgnoreCase);
            if (bookmarkMatch.Success)
            {
                return $@"\l ""{bookmarkMatch.Groups[1].Value}""";
            }

            if (instruction.StartsWith("HYPERLINK \"", StringComparison.OrdinalIgnoreCase))
            {
                var match = Regex.Match(instruction, @"""([^""]+)""");
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }

                return instruction.Replace("HYPERLINK \"", "").Replace("\"", "").Trim();
            }

            var unquotedMatch = Regex.Match(
                instruction,
                @"^HYPERLINK\s+([^\s]+)",
                RegexOptions.IgnoreCase);
            if (unquotedMatch.Success)
            {
                return unquotedMatch.Groups[1].Value.Trim();
            }

            return null;
        }

        private static string SanitizeFieldInstruction(string instruction)
        {
            if (string.IsNullOrEmpty(instruction))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(instruction.Length);
            foreach (char c in instruction)
            {
                if (c == '\r' || c == '\n' || c == '\t')
                {
                    builder.Append(' ');
                }
                else if (!char.IsControl(c))
                {
                    builder.Append(c);
                }
            }

            return Regex.Replace(builder.ToString(), @"\s+", " ").Trim();
        }

    }
}
