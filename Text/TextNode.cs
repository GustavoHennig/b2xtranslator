using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt
{
    public class TextNode : INode
    {
        public List<IAttribute> Attributes { get; } = new List<IAttribute>();

        public IDocument OwnerDocument {get; private set; }

      

        public List<INode> ChildNodes {get; } = new List<INode>();

        string INode.Name => "";

        public string Value { get; set; }

        public TextNode(IDocument ownerDocument)
        {
            OwnerDocument = ownerDocument;
        }
        public void AppendChild(INode tblInd)
        {
            ChildNodes.Add(tblInd);
        }

        public void RemoveChild(INode bdr)
        {
            if (ChildNodes.Contains(bdr))
            {
                ChildNodes.Remove(bdr);
            }
        }

        public void WriteTo(IWriter writer)
        {
            writer.WriteNode(this);
        }
    }
}
