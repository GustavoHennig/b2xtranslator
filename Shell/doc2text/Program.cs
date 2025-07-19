using b2xtranslator.Common.Exceptions;
using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.Shell;
using b2xtranslator.StructuredStorage.Common;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.txt;
using b2xtranslator.txt.TextMapping;
using System;
using System.Globalization;
using System.IO;

namespace b2xtranslator.doc2x
{
    public class Program : CommandLineTranslator
    {
        public static string ToolName = "doc2text";
        public static string ContextMenuInputExtension = ".doc";
        public static string ContextMenuText = "Convert to .txt";

        public static void Main(string[] args)
        {
            ParseArgs(args, ToolName);

            InitializeLogger();

            PrintWelcome(ToolName, "-1");

            // convert
            try
            {
                //copy processing file
                var procFile = new ProcessingFile(InputFile);

                //make output file name
                if (ChoosenOutputFile == null)
                {
                    if (InputFile.Contains("."))
                    {
                        ChoosenOutputFile = InputFile.Remove(InputFile.LastIndexOf(".")) + ".txt";
                    }
                    else
                    {
                        ChoosenOutputFile = InputFile + ".txt";
                    }
                }


               DocTextExtractor.ConvertDocToTxt(
                    InputFile,
                    ChoosenOutputFile
                );

            }
            catch (DirectoryNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (FileNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ReadBytesAmountMismatchException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MagicNumberException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (UnspportedFileVersionException ex)
            {
                TraceLogger.Error("File {0} has been created with a Word version older than Word 97.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ByteParseException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MappingException ex)
            {
                TraceLogger.Error("There was an error while converting file {0}: {1}", InputFile, ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (EncryptedFileException ex)
            {
                TraceLogger.Error("File {0} is encrypted and cannot be converted.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (Exception ex)
            {
                TraceLogger.Error("Conversion of file {0} failed: {1}", InputFile, ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
        }
    }
}
