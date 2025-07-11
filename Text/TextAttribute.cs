using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt
{
    public  class TextAttribute : IAttribute
    {
        public string Name { get; private set; }
        public string Value { get; set; }
        public TextAttribute(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }
        public void WriteTo(IWriter writer)
        {
            writer.WriteAttributeString(this.Name, this.Value);
        }
    }
}
