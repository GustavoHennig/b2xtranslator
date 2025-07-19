using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt.TextModel
{
    public  class TextAttribute : IAttribute
    {
        public string? Prefix { get; private set; }
        public string LocalName { get; private set; }
        public string? Value { get; set; } = null;
        public TextAttribute(string? prefix, string locaName, string? value)
        {
            Prefix = prefix;
            LocalName = locaName;
            Value = value;
        }
        public void WriteTo(IWriter writer)
        {
            writer.WriteAttributeString(Prefix, LocalName, null, Value);
        }
    }
}
