using System.Xml;

namespace b2xtranslator.txt
{
    public abstract class AbstractMapping
    {
        protected IWriter _writer;
        protected IDocument _nodeFactory;

        public AbstractMapping(IWriter writer)
        {
            this._writer = writer;
            this._nodeFactory = TextDocument.Create("", writer);
        }
    }
}
