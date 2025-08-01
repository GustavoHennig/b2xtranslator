﻿using b2xtranslator.OpenXmlLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace b2xtranslator.txt.TextModel
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
            FilePath = fileName;
            FootnotesWriter = writer;
            EndnotesWriter = writer;
            MainDocumentWriter = writer;
            CommentsPartWriter = writer;
        }

        public static TextDocument Create(string fileName, IWriter? writer = null, bool extractUrls = true)
        {
            var doc = new TextDocument(fileName, writer ?? new TextWriter(extractUrls));

            return doc;
        }

        public IAttribute CreateAttribute(string? prefix, string localName, string? namespaceURI)
        {
            return new TextAttribute(prefix, localName, null)
            {
            };
        }

        public INode CreateElement(string? prefix, string localName, string? namespaceURI)
        {
            return new TextNode(this, $"{prefix}:{localName}")
            {
            };
        }
    }
}
