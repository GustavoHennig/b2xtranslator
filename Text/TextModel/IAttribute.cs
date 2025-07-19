namespace b2xtranslator.txt.TextModel
{
    public interface IAttribute
    {
        string? Value { get; set; }
        string? Prefix { get; }
        string LocalName { get; }
        void WriteTo(IWriter writer);
    }
}