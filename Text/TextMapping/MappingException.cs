using System;

namespace b2xtranslator.txt.TextMapping
{
    public class MappingException : Exception
    {
        public MappingException(string message)
            : base(message)
        { }
    }
}
