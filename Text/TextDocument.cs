using b2xtranslator.OpenXmlLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt
{

    public interface IDocument
    {
        IAttribute CreateAttribute(string v1, string v2, string wordprocessingML);
        INode CreateElement(string v1, string v2, string wordprocessingML);
    }

    public class TextDocument : IDocument
    {
        public IWriter FootnotesWriter { get; internal set; }
        public IWriter EndnotesWriter { get; internal set; }
        public IWriter MainDocumentWriter { get; internal set; }
        public IWriter CommentsPartWriter { get; internal set; }


        public string FilePath { get; internal set; }

        protected TextDocument(string fileName, IWriter writer)
        {
            this.FilePath = fileName;
            this.FootnotesWriter = writer;
            this.EndnotesWriter = writer;
            this.MainDocumentWriter = writer;
            this.CommentsPartWriter = writer;
        }

        public static TextDocument Create(string fileName, IWriter? writer = null)
        {
            var doc = new TextDocument(fileName, writer ?? new TextWriter());

            return doc;
        }

        public IAttribute CreateAttribute(string v1, string v2, string wordprocessingML)
        {
            return new TextAttribute(v1, v2)
            {
                Value = wordprocessingML,
            };
        }

        public INode CreateElement(string v1, string v2, string wordprocessingML)
        {
            return new TextNode(this)
            {
                Value = v1,
            };
        }
    }
}
