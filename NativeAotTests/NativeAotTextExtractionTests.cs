using System.Diagnostics;
using System.Runtime.InteropServices;
using Xunit;
using Xunit.Sdk;

namespace b2xtranslator.NativeAotTests;

public class NativeAotTextExtractionTests : IClassFixture<NativeAotPublishFixture>
{
    private readonly NativeAotPublishFixture _fixture;

    public NativeAotTextExtractionTests(NativeAotPublishFixture fixture)
    {
        _fixture = fixture;
    }

    public static IEnumerable<object[]> DocFiles()
    {
        var baseDir = AppContext.BaseDirectory;
        var examplesDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "samples"));
        var examplesLocalDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "samples-local"));

        if (!Directory.Exists(examplesDir))
            throw new DirectoryNotFoundException($"Examples directory not found: {examplesDir}");

        var docFiles = Directory.GetFiles(examplesDir, "*.doc").ToList();

        if (Directory.Exists(examplesLocalDir))
            docFiles.AddRange(Directory.GetFiles(examplesLocalDir, "*.doc"));

        return docFiles
            .Where(w => File.Exists(Path.ChangeExtension(w, ".expected.txt")))
            .Select(doc => new object[] { doc, Path.ChangeExtension(doc, ".expected.txt") });
    }

    [Theory]
    [MemberData(nameof(DocFiles))]
    public void ExtractedText_EqualsExpectedFile(string docPath, string expectedPath)
    {
        if (!File.Exists(expectedPath))
            throw SkipException.ForSkip($"Expected file not found: {expectedPath}");

        string expected = NormalizeText(File.ReadAllText(expectedPath));
        string outputFile = Path.Combine(Path.GetTempPath(), $"nativeaot-test-{Guid.NewGuid()}.txt");

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _fixture.ExecutablePath,
                Arguments = $"\"{docPath}\" \"{outputFile}\"",
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi)!;
            string stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                if (!string.IsNullOrEmpty(stderr) && stderr.Contains(expected, StringComparison.InvariantCultureIgnoreCase))
                {
                    File.Delete(Path.ChangeExtension(docPath, ".actual.txt"));
                    File.Delete(Path.ChangeExtension(docPath, ".error.txt"));
                    return;
                }

                File.WriteAllText(Path.ChangeExtension(docPath, ".error.txt"), stderr);
                Assert.Fail($"NativeAOT executable failed (exit code {process.ExitCode}): {stderr}");
            }

            string resultOriginal = File.ReadAllText(outputFile);
            string result = NormalizeText(resultOriginal);
            bool isEqual = string.Equals(result, expected, StringComparison.InvariantCultureIgnoreCase);

            if (!isEqual)
                File.WriteAllText(Path.ChangeExtension(docPath, ".actual.txt"), resultOriginal);
            else
                File.Delete(Path.ChangeExtension(docPath, ".actual.txt"));

            File.Delete(Path.ChangeExtension(docPath, ".error.txt"));
            Assert.Equal(expected, result, true, true, true, true);
        }
        finally
        {
            if (File.Exists(outputFile))
                File.Delete(outputFile);
        }
    }

    public static string NormalizeText(string text)
    {
        if (text == null) return null!;
        var normalized = text
            .Replace("\r\n", "\n")
            .Replace("\r", "\n")
            .Replace("\t", "")
            .Replace("  ", " ")
            .Replace("\n\n", "\n")
            .Replace("\n\n", "\n");

        var lines = normalized.Split('\n').Select(line => line.Trim()).Where(w => !string.IsNullOrWhiteSpace(w));
        var result = string.Join("\n", lines);

        return result.TrimEnd(' ', '\n', '\r');
    }
}

public sealed class NativeAotPublishFixture : IDisposable
{
    public string ExecutablePath { get; }
    public string PublishDirectory { get; }

    public NativeAotPublishFixture()
    {
        var baseDir = AppContext.BaseDirectory;
        var solutionRoot = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
        var projectPath = Path.Combine(solutionRoot, "Shell", "doc2text.aot", "doc2text.aot.csproj");

        string rid = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "win-x64"
            : RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ? "linux-x64"
            : RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "osx-x64"
            : "win-x64";

        if (RuntimeInformation.OSArchitecture == Architecture.Arm64)
            rid = rid.Replace("x64", "arm64");

        PublishDirectory = Path.Combine(solutionRoot, "artifacts", "nativeaot", rid);

        string exeName = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "doc2text.aot.exe" : "doc2text.aot";
        ExecutablePath = Path.Combine(PublishDirectory, exeName);

        if (File.Exists(ExecutablePath))
            return;

        Directory.CreateDirectory(PublishDirectory);

        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = $"publish \"{projectPath}\" -c Release -r {rid} -o \"{PublishDirectory}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi)!;
        string stdout = process.StandardOutput.ReadToEnd();
        string stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"NativeAOT publish failed (exit code {process.ExitCode}).\nOutput:\n{stdout}\nError:\n{stderr}");
        }

        if (!File.Exists(ExecutablePath))
            throw new FileNotFoundException($"NativeAOT executable not found at {ExecutablePath} after publish.");
    }

    public void Dispose()
    {
    }
}
