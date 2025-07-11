
namespace b2xtranslator.txt
{

    public interface IAttribute
    {
        string Name { get; }
        string Value { get; set; }

        void WriteTo(IWriter writer);
    }

    public interface INode
    {
        List<IAttribute> Attributes { get; }
        IDocument OwnerDocument { get; }
        List<INode> ChildNodes { get; }
        string Name { get; }
        string Value { get; set; }

        void AppendChild(INode tblInd);
        void RemoveChild(INode bdr);
        void WriteTo(IWriter writer);
    }
}