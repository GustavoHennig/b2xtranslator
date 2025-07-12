using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt
{
    public  class TextAttribute : IAttribute
    {
        public string? Prefix { get; private set; }
        public string LocalName { get; private set; }
        public string? Value { get; set; } = null;
        public TextAttribute(string? prefix, string locaName, string? value)
        {
            this.Prefix = prefix;
            this.LocalName = locaName;
            this.Value = value;
        }
        public void WriteTo(IWriter writer)
        {
            writer.WriteAttributeString(this.Prefix, LocalName, null, this.Value);
        }
    }
}
