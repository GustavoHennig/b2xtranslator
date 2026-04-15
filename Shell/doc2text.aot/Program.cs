using b2xtranslator.txt;
using System;
using System.IO;

if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: doc2text.aot <input.doc> [output.txt]");
    return 1;
}

string inputFile = Path.GetFullPath(args[0]);

if (!File.Exists(inputFile))
{
    Console.Error.WriteLine($"File not found: {inputFile}");
    return 1;
}

string? outputFile = args.Length >= 2 ? args[1] : null;

try
{
    string text = DocTextExtractor.ExtractTextFromFile(inputFile);

    if (outputFile != null)
        File.WriteAllText(outputFile, text);
    else
        Console.Write(text);

    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    return 2;
}
