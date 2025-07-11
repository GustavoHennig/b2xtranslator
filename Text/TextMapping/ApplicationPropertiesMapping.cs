using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.txt.TextMapping
{
    public class ApplicationPropertiesMapping : AbstractMapping,
          IMapping<DocumentProperties>
    {

        public ApplicationPropertiesMapping(AppPropertiesPart appPart, IWriter writer)
            : base(writer)
        {
        }

        public void Apply(DocumentProperties dop)
        {
            //start Properties
            _writer.WriteStartElement("w", "Properties", OpenXmlNamespaces.WordprocessingML);

            //Application
            //AppVersion
            //Company
            //DigSig
            //DocSecurity
            //HeadingPairs
            //HiddenSlides
            //HLinks
            //HyperlinkBase
            //HyperlinksChanged
            //LinksUpToDate
            //Manager
            //MMClips
            //Notes
            //PresentationFormat
            //ScaleCrop
            //SharedDoc
            //Slides
            //Template
            //TitlesOfParts
            //TotalTime

            //WordCount statistics

            //CharactersWithSpaces
            _writer.WriteStartElement("CharactersWithSpaces");
            _writer.WriteString(dop.cChWS.ToString());
            _writer.WriteEndElement();

            //Characters
            _writer.WriteStartElement("Characters");
            _writer.WriteString(dop.cCh.ToString());
            _writer.WriteEndElement();

            //Lines
            _writer.WriteStartElement("Lines");
            _writer.WriteString(dop.cLines.ToString());
            _writer.WriteEndElement();

            //Pages
            _writer.WriteStartElement("Pages");
            _writer.WriteString(dop.cPg.ToString());
            _writer.WriteEndElement();

            //Paragraphs
            _writer.WriteStartElement("Paragraphs");
            _writer.WriteString(dop.cParas.ToString());
            _writer.WriteEndElement();

            //Words
            _writer.WriteStartElement("Words");
            _writer.WriteString(dop.cWords.ToString());
            _writer.WriteEndElement();

            //end Properties
            _writer.WriteEndElement();
        }
    }
}
