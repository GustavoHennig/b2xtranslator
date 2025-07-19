using b2xtranslator.DocFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.Shell;
using System.Globalization;
using System.Diagnostics;
using b2xtranslator.txt.TextMapping;
using b2xtranslator.txt.TextModel;
using TextWriter = b2xtranslator.txt.TextModel.TextWriter;

namespace b2xtranslator.txt
{
    public class DocTextExtractor
    {
       
        public static void ConvertDocToTxt(string docFilePath, string outputFilePath)
        {
            var stopwatch = Stopwatch.StartNew();

            using (var reader = new StructuredStorageReader(docFilePath))
            {
                var doc = new WordDocument(reader);


                var textDoc = TextDocument.Create(outputFilePath, null, CommandLineTranslator.ExtractUrls);
             
                TraceLogger.Info("Converting file {0} into {1}", docFilePath, outputFilePath);

                string output = ExtractTextFromFile(docFilePath, CommandLineTranslator.ExtractUrls);

                File.WriteAllText(outputFilePath, output);


                stopwatch.Stop();
                TraceLogger.Info("Conversion of file {0} finished in {1} seconds", docFilePath, stopwatch.Elapsed.TotalSeconds.ToString(CultureInfo.InvariantCulture));
            }
        }

        public static string ExtractTextFromFile(string docFilePath, bool extractUrls = true)
        {
            //open the reader
            using (var reader = new StructuredStorageReader(docFilePath))
            {
                //parse the input document
                var doc = new WordDocument(reader);

                var textDoc = TextDocument.Create("", null,  extractUrls);
               
                //convert the document
                return ConvertToString(doc, textDoc, extractUrls);
            }
        }

        public static string ConvertToString(WordDocument doc, TextDocument textDocument, bool extractUrls = true)
        {
            var context = new ConversionContext(doc, extractUrls);
            {
                //Setup the context
                context.TextDoc = textDocument;



                //convert the command table (skip for Word 95 files where it may be null)
                if (doc.CommandTable != null)
                {
                    doc.CommandTable.Convert(new CommandTableMapping(context, textDocument.FootnotesWriter));
                }

                //Write styles (skip for Word 95 files where it may be null)
                if (doc.Styles != null)
                {
                    doc.Styles.Convert(new StyleSheetMapping(context, doc, textDocument.FootnotesWriter));
                }

                //Write numbering (skip for Word 95 files where it may be null)
                if (doc.ListTable != null)
                {
                    doc.ListTable.Convert(new NumberingMapping(context, doc, context.TextDoc.MainDocumentWriter));
                }

                //Write fontTable (skip for Word 95 files where it may be null)
                if (doc.FontTable != null)
                {
                    doc.FontTable.Convert(new FontTableMapping(context, new TextWriter(context.ExtractUrls)));
                }

                //write document and the header and footers
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


                // TODO: Put a final cleanup here if needed, for example:
                string cleanText = context.TextDoc.MainDocumentWriter.ToString().Replace("\u2002", " ");

                return cleanText;

            }
        }
    }
}
