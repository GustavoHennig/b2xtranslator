using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.DocFileFormat
{
    public class SectionPropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Creates a SEPX which doesn't modify anything.<br/>
        /// The grpprl list is empty (for Word 95 support)
        /// </summary>
        public SectionPropertyExceptions() : base()
        {
        }

        /// <summary>
        /// Parses the bytes to retrieve a SectionPropertyExceptions
        /// </summary>
        /// <param name="bytes">The bytes starting with the grpprl</param>
        public SectionPropertyExceptions(byte[] bytes)
            : base(bytes)
        {
        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            ((IMapping<SectionPropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
