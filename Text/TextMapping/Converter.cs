using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.WordprocessingMLMapping;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace b2xtranslator.txt.TextMapping
{
    public class Converter
    {
        public static OpenXmlPackage.DocumentType DetectOutputType(WordDocument doc)
        {
            var returnType = OpenXmlPackage.DocumentType.Document;

            //detect the document type
            if (doc.FIB.fDot)
            {
                //template
                if (doc.CommandTable.MacroDatas != null && doc.CommandTable.MacroDatas.Count > 0)
                {
                    //macro enabled template
                    returnType = OpenXmlPackage.DocumentType.MacroEnabledTemplate;
                }
                else
                {
                    //without macros
                    returnType = OpenXmlPackage.DocumentType.Template;
                }
            }
            else
            {
                //no template
                if (doc.CommandTable.MacroDatas != null && doc.CommandTable.MacroDatas.Count > 0)
                {
                    //macro enabled document
                    returnType = OpenXmlPackage.DocumentType.MacroEnabledDocument;
                }
                else
                {
                    returnType = OpenXmlPackage.DocumentType.Document;
                }
            }

            return returnType;
        }


        public static string GetConformFilename(string choosenFilename, OpenXmlPackage.DocumentType outType)
        {
            string outExt = ".docx";
            switch (outType)
            {
                case OpenXmlPackage.DocumentType.Document:
                    outExt = ".docx";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    outExt = ".docm";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    outExt = ".dotm";
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    outExt = ".dotx";
                    break;
                default:
                    outExt = ".docx";
                    break;
            }

            string inExt = Path.GetExtension(choosenFilename);
            if (inExt != null)
            {
                return choosenFilename.Replace(inExt, outExt);
            }
            else
            {
                return choosenFilename + outExt;
            }
        }

        public static void ConvertFiles(string inputFilePath, string outputFilePath)
        {
            //open the reader
            using (var reader = new StructuredStorageReader(inputFilePath))
            {
                //parse the input document
                var doc = new WordDocument(reader);


                var textDoc = TextDocument.Create(outputFilePath);
                //start time
                var start = DateTime.Now;
                TraceLogger.Info("Converting file {0} into {1}", inputFilePath, outputFilePath);

                //convert the document
                string output = ConvertFileToString(inputFilePath);

                File.WriteAllText(outputFilePath, output);


                var end = DateTime.Now;
                var diff = end.Subtract(start);


                TraceLogger.Info("Conversion of file {0} finished in {1} seconds", inputFilePath, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
            }
        }

        public static string ConvertFileToString(string inputFilePath)
        {
            //open the reader
            using (var reader = new StructuredStorageReader(inputFilePath))
            {
                //parse the input document
                var doc = new WordDocument(reader);

                var textDoc = TextDocument.Create("");
               
                //convert the document
                return Converter.ConvertToString(doc, textDoc);

            }
        }

        public static void Convert(WordDocument doc, TextDocument textDocument)
        {
            //Create a new conversion context
            var context = new ConversionContext(doc);
            string output = ConvertToString(doc, textDocument);
            File.WriteAllText(context.TextDoc.FilePath, output);

        }

        public static string ConvertToString(WordDocument doc, TextDocument textDocument)
        {
            var context = new ConversionContext(doc);
            {
                //Setup the context
                context.TextDoc = textDocument;



                //convert the command table
                doc.CommandTable.Convert(new CommandTableMapping(context, textDocument.FootnotesWriter));

                //Write styles.xml
                doc.Styles.Convert(new StyleSheetMapping(context, doc, textDocument.FootnotesWriter));

                //Write numbering.xml
                doc.ListTable.Convert(new NumberingMapping(context, doc, context.TextDoc.MainDocumentWriter));

                //Write fontTable.xml
                doc.FontTable.Convert(new FontTableMapping(context, new TextWriter()));

                //write document.xml and the header and footers
                doc.Convert(new MainDocumentMapping(context, context.TextDoc.MainDocumentWriter));

                //write the footnotes
                doc.Convert(new FootnotesMapping(context));

                //write the endnotes
                doc.Convert(new EndnotesMapping(context));

                //write the comments
                doc.Convert(new CommentsMapping(context));

                //write settings.xml at last because of the rsid list
                //doc.DocumentProperties.Convert(new SettingsMapping(context, docx.MainDocumentPart.SettingsPart, writer));

                //convert the glossary subdocument
                if (doc.Glossary != null)
                {
                    doc.Glossary.Convert(new GlossaryMapping(context, context.TextDoc.MainDocumentWriter));
                    //doc.Glossary.FontTable.Convert(new FontTableMapping(context, docx.MainDocumentPart.GlossaryPart.FontTablePart, writer));
                    //doc.Glossary.Styles.Convert(new StyleSheetMapping(context, doc.Glossary, docx.MainDocumentPart.GlossaryPart.StyleDefinitionsPart));

                    //write settings.xml at last because of the rsid list
                    //doc.Glossary.DocumentProperties.Convert(new SettingsMapping(context, docx.MainDocumentPart.GlossaryPart.SettingsPart, writer));
                }

                return context.TextDoc.MainDocumentWriter.ToString();

            }
        }
    }
}
