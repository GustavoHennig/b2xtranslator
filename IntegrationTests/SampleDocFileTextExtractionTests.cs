using System;
using System.IO;
using System.Linq;
using Xunit;
using System.Diagnostics;
using Xunit.Sdk;
using b2xtranslator.txt;
using b2xtranslator.txt.TextMapping;

namespace b2xtranslator.Tests
{
    public class SampleDocFileTextExtractionTests
    {
        public static IEnumerable<object[]> DocFiles()
        {
            // Determine solution root and examples folder
            var baseDir = AppContext.BaseDirectory;
            var examplesDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "samples"));
            var examplesLocalDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "samples-local"));

            if (!Directory.Exists(examplesDir))
                throw new DirectoryNotFoundException($"Examples directory not found: {examplesDir}");

            // Find all .doc files
            var docFiles = Directory.GetFiles(examplesDir, "*.doc").ToList();


            if (Directory.Exists(examplesLocalDir))
            {
                docFiles.AddRange(Directory.GetFiles(examplesLocalDir, "*.doc"));
            }

            docFiles
                .Where(w => !File.Exists(Path.ChangeExtension(w, ".expected.txt")))
                .ToList()
                .ForEach(doc =>
                {
                    Debug.Print($"Expected file not found for {doc}, skipping test.");
                    File.WriteAllText(Path.ChangeExtension(doc, ".expected.txt"), "Expected text not provided.");
                });


            return docFiles
                .Where(w => File.Exists(Path.ChangeExtension(w, ".expected.txt")))
                .Select(doc => new object[] { doc, Path.ChangeExtension(doc, ".expected.txt") });
        }

        [Theory]
        [MemberData(nameof(DocFiles))]
        public void ExtractedText_EqualsExpectedFile(string docPath, string expectedPath)
        {
            if (!File.Exists(expectedPath))
            {
                Debug.Print($"Expected file not found: {expectedPath}");
                throw SkipException.ForSkip($"Expected file not found: {expectedPath}");
            }
            string result;
            string resultOriginal;
            string expected;
            try
            {
                expected = NormalizeText(File.ReadAllText(expectedPath));
                resultOriginal = Converter.ConvertFileToString(docPath);
                result = NormalizeText(resultOriginal);
                bool isEqual = string.Equals(result, expected, StringComparison.InvariantCultureIgnoreCase);
                if (!isEqual)
                {
                    Debug.Print($"Mismatch in {docPath}");
                    File.WriteAllText(Path.ChangeExtension(docPath, ".actual.txt"), resultOriginal);
                }
                else
                {
                    //Rewrite expected to make all line-breaks match
                    //File.WriteAllText(Path.ChangeExtension(docPath, ".expected.txt"), resultOriginal);
                    File.Delete(Path.ChangeExtension(docPath, ".actual.txt"));
                }
                File.Delete(Path.ChangeExtension(docPath, ".error.txt"));

            }
            catch (Exception ex)
            {
                File.WriteAllText(Path.ChangeExtension(docPath, ".error.txt"), ex.ToString());
                File.Delete(Path.ChangeExtension(docPath, ".actual.txt"));
                throw;
            }
            Assert.Equal(expected, result, true, true, true, true);


        }
        /// <summary>
        /// Normalizes text by standardizing line breaks to '\n' and trimming trailing whitespace from each line.
        /// </summary>
        public static string NormalizeText(string text)
        {
            if (text == null) return null;
            // Replace CRLF and CR with LF
            var normalized = text
                .Replace("\r\n", "\n")
                .Replace("\r", "\n").Replace("\t", "").Replace("  ", " ");
            // Trim trailing whitespace from each line
            var lines = normalized.Split('\n').Select(line => line.TrimEnd());
            return string.Join("\n", lines);
        }


    }
}
