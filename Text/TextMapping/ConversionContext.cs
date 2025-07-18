using System.Collections.Generic;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib.WordprocessingML;

namespace b2xtranslator.txt.TextMapping
{
    public class ConversionContext
    {
        
        /// <summary>
        /// The source of the conversion.
        /// </summary>
        public WordDocument Doc { get; set; }

        /// <summary>
        /// This is the target of the conversion.<br/>
        /// The result will be written to the parts of this document.
        /// </summary>
        public TextDocument TextDoc { get; set; }

        /// <summary>
        /// Flag indicating whether to extract URLs from hyperlinks in text output.
        /// </summary>
        public bool ExtractUrls { get; set; }

        /// <summary>
        /// A list thta contains all revision ids.
        /// </summary>
        public List<string> AllRsids;

        public ConversionContext(WordDocument doc, bool extractUrls = true)
        {
            Doc = doc;
            ExtractUrls = extractUrls;
            AllRsids = new List<string>();
        }

        /// <summary>
        /// Adds a new RSID to the list
        /// </summary>
        /// <param name="rsid"></param>
        public void AddRsid(string rsid)
        {
            if (!AllRsids.Contains(rsid))
                AllRsids.Add(rsid);
        }
    }
}
