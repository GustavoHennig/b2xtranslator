# Issue #012: Make Text Extraction Configurable

## Problem Description
The text extraction process is not configurable. Users cannot choose to selectively include or exclude certain types of content from the document, such as headers, footers, comments, or textboxes. This limits the tool's flexibility for different use cases.

## Severity
**LOW** (Enhancement)

## Impact
- Users lack fine-grained control over the text extraction process.
- Inability to generate "clean" text output by excluding metadata or non-essential content.
- The shell application is less powerful than it could be.

## Root Cause Analysis
This is a missing feature. The `doc2text` shell application and the underlying `TextConverter` library were not designed with configurability in mind.

1.  **No Options API**: The `TextConverter` class does not accept any configuration object.
2.  **No CLI Parsing**: The `doc2text` `Program.cs` does not parse any command-line arguments beyond the input/output file paths.
3.  **Hardcoded Logic**: The `TextMapping` processes all content types it encounters without any conditional logic to skip sections.

## Known Affected Components
- **Shell Application**: `Shell/doc2text/`
- **Text Converter**: `Text/TextMapping/TextConverter.cs`
- **Text Mapping Logic**: `Text/TextMapping/TextMapping.cs`

## Reproduction Steps
N/A (This is a feature request).

## Proposed Solutions

### 1. Create a Configuration Class
- In the `Text/` project, create a new class, e.g., `TextExtractionSettings`.
- Add boolean properties like `IncludeHeadersAndFooters`, `IncludeTextBoxes`, `IncludeComments`, `IncludeFootnotes`, etc. All should default to `true`.

### 2. Update the Converter
- Modify the `TextConverter.Convert` method to accept an instance of `TextExtractionSettings`.
- Pass this settings object down to the `TextMapping` instance.

### 3. Update the Shell Application
- Use a simple command-line parsing library (or manual parsing) in `Shell/doc2text/Program.cs`.
- Add flags like `--no-headers-footers`, `--no-textboxes`, `--no-comments`.
- When a flag is present, create the `TextExtractionSettings` object with the corresponding property set to `false`.
- Pass the settings object to the `TextConverter`.

### 4. Update the Mapping Logic
- In `TextMapping.cs`, before processing a given element (e.g., a header, a comment), check the settings object.
- If the corresponding setting is `false`, skip the processing of that element and its children.

## Files to Investigate
- `Shell/doc2text/Program.cs`
- `Text/TextMapping/TextConverter.cs`
- `Text/TextMapping/TextMapping.cs`
- `Doc/DocFileFormat/WordDocument.cs` (to see how different content parts are stored and accessed)

## Success Criteria
1.  **Flags Work**: Users can control the text extraction output using command-line flags.
2.  **Content Excluded**: When a flag like `--no-headers-footers` is used, the corresponding content is absent from the output file.
3.  **Default Behavior Unchanged**: If no flags are provided, the output is the same as it was before, with all content included.

## Current Status (Verified)
**FEATURE NOT IMPLEMENTED** - Testing shows configurable extraction is not available:
- No command-line argument parsing beyond input/output files
- Attempting `--help` results in file not found error 
- No configuration options available for selective content extraction
- Feature needs to be implemented from scratch

## Detailed Resolution Plan

### Phase 1: Command-Line Argument Infrastructure (1-2 days)

#### 1.1 Add Command-Line Parsing Library
**File: `Shell/doc2text/doc2text.csproj`**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
  </PropertyGroup>
  
  <ItemGroup>
    <!-- Add command-line parsing support -->
    <PackageReference Include="System.CommandLine" Version="2.0.0-beta4.22272.1" />
  </ItemGroup>
  
  <!-- Existing references -->
</Project>
```

#### 1.2 Redesign Program.cs for Argument Parsing
**File: `Shell/doc2text/Program.cs`**

```csharp
using System.CommandLine;
using System.CommandLine.Invocation;

namespace doc2text
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            var rootCommand = new RootCommand("Convert Word documents to plain text")
            {
                CreateInputOption(),
                CreateOutputOption(),
                CreateNoHeadersFootersOption(),
                CreateNoTextBoxesOption(),
                CreateNoCommentsOption(),
                CreateNoFootnotesOption(),
                CreateHelpOption()
            };

            rootCommand.SetHandler(async (context) =>
            {
                var settings = BuildExtractionSettings(context);
                var inputFile = context.ParseResult.GetValueForOption(InputOption);
                var outputFile = context.ParseResult.GetValueForOption(OutputOption);
                
                await ConvertDocument(inputFile, outputFile, settings);
            });

            return await rootCommand.InvokeAsync(args);
        }

        private static Option<string> CreateInputOption() =>
            new Option<string>(
                name: "--input",
                description: "Input Word document file")
            {
                IsRequired = true,
                ArgumentHelpName = "input.doc"
            };

        private static Option<string> CreateOutputOption() =>
            new Option<string>(
                name: "--output", 
                description: "Output text file")
            {
                IsRequired = true,
                ArgumentHelpName = "output.txt"
            };

        private static Option<bool> CreateNoHeadersFootersOption() =>
            new Option<bool>(
                name: "--no-headers-footers",
                description: "Exclude headers and footers from extraction");

        private static Option<bool> CreateNoTextBoxesOption() =>
            new Option<bool>(
                name: "--no-textboxes",
                description: "Exclude text box content from extraction");

        private static Option<bool> CreateNoCommentsOption() =>
            new Option<bool>(
                name: "--no-comments",
                description: "Exclude comments from extraction");

        private static Option<bool> CreateNoFootnotesOption() =>
            new Option<bool>(
                name: "--no-footnotes",
                description: "Exclude footnotes and endnotes from extraction");
    }
}
```

### Phase 2: Text Extraction Settings Framework (1-2 days)

#### 2.1 Create Configuration Classes
**File: `Text/TextExtractionSettings.cs`** (create)

```csharp
namespace b2xtranslator.Text
{
    public class TextExtractionSettings
    {
        public bool IncludeHeadersAndFooters { get; set; } = true;
        public bool IncludeTextBoxes { get; set; } = true;
        public bool IncludeComments { get; set; } = true;
        public bool IncludeFootnotes { get; set; } = true;
        public bool IncludeEndnotes { get; set; } = true;
        public bool IncludePictures { get; set; } = false;  // Alt text only
        public bool IncludeHyperlinks { get; set; } = true;
        public bool IncludeTables { get; set; } = true;
        
        // Advanced settings
        public bool PreserveParagraphBreaks { get; set; } = true;
        public bool PreserveLineBreaks { get; set; } = false;
        public string ListBulletCharacter { get; set; } = "•";
        public int MaxNestingDepth { get; set; } = 10;
        
        public static TextExtractionSettings Default => new TextExtractionSettings();
        
        public static TextExtractionSettings CreateFromCommandLine(ParseResult parseResult)
        {
            return new TextExtractionSettings
            {
                IncludeHeadersAndFooters = !parseResult.GetValueForOption("--no-headers-footers"),
                IncludeTextBoxes = !parseResult.GetValueForOption("--no-textboxes"),
                IncludeComments = !parseResult.GetValueForOption("--no-comments"),
                IncludeFootnotes = !parseResult.GetValueForOption("--no-footnotes")
            };
        }
    }
}
```

#### 2.2 Create Extraction Context
**File: `Text/TextExtractionContext.cs`** (create)

```csharp
namespace b2xtranslator.Text
{
    public class TextExtractionContext
    {
        public TextExtractionSettings Settings { get; }
        public ITextWriter Writer { get; }
        public WordDocument Document { get; }
        
        // Processing state
        public bool IsInHeaderFooter { get; set; }
        public bool IsInComment { get; set; }
        public bool IsInTextBox { get; set; }
        public bool IsInFootnote { get; set; }
        public int NestingLevel { get; set; }
        
        public TextExtractionContext(
            TextExtractionSettings settings, 
            ITextWriter writer, 
            WordDocument document)
        {
            Settings = settings ?? TextExtractionSettings.Default;
            Writer = writer ?? throw new ArgumentNullException(nameof(writer));
            Document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        public bool ShouldProcessCurrentContent()
        {
            if (IsInHeaderFooter && !Settings.IncludeHeadersAndFooters) return false;
            if (IsInComment && !Settings.IncludeComments) return false;
            if (IsInTextBox && !Settings.IncludeTextBoxes) return false;
            if (IsInFootnote && !Settings.IncludeFootnotes) return false;
            if (NestingLevel > Settings.MaxNestingDepth) return false;
            
            return true;
        }
    }
}
```

### Phase 3: Update Text Mapping Architecture (2-3 days)

#### 3.1 Enhance Base Mapping Classes
**File: `Text/AbstractMapping.cs`**

```csharp
public abstract class AbstractMapping : IMapping
{
    protected TextExtractionContext Context { get; }
    
    protected AbstractMapping(TextExtractionContext context)
    {
        Context = context ?? throw new ArgumentNullException(nameof(context));
    }
    
    public virtual void Apply(IVisitable x)
    {
        if (Context.ShouldProcessCurrentContent())
        {
            ApplyCore(x);
        }
    }
    
    protected abstract void ApplyCore(IVisitable x);
    
    protected void WriteText(string text)
    {
        if (Context.ShouldProcessCurrentContent())
        {
            Context.Writer.WriteText(text);
        }
    }
}
```

#### 3.2 Update Document Mapping
**File: `Text/TextMapping/DocumentMapping.cs`**

```csharp
public class DocumentMapping : AbstractMapping
{
    public DocumentMapping(TextExtractionContext context) : base(context) { }
    
    protected override void ApplyCore(IVisitable x)
    {
        var doc = x as WordDocument;
        
        // Process main document content
        ProcessMainDocument(doc);
        
        // Process headers/footers if enabled
        if (Context.Settings.IncludeHeadersAndFooters)
        {
            ProcessHeadersAndFooters(doc);
        }
        
        // Process comments if enabled
        if (Context.Settings.IncludeComments)
        {
            ProcessComments(doc);
        }
        
        // Process footnotes if enabled
        if (Context.Settings.IncludeFootnotes)
        {
            ProcessFootnotes(doc);
        }
        
        // Process textboxes if enabled
        if (Context.Settings.IncludeTextBoxes)
        {
            ProcessTextBoxes(doc);
        }
    }
    
    private void ProcessHeadersAndFooters(WordDocument doc)
    {
        Context.IsInHeaderFooter = true;
        try
        {
            // Process headers and footers
            foreach (var header in doc.Headers)
            {
                header.Convert(new HeaderMapping(Context));
            }
        }
        finally
        {
            Context.IsInHeaderFooter = false;
        }
    }
}
```

### Phase 4: Conditional Content Processing (1-2 days)

#### 4.1 Update Specific Mapping Classes
**File: `Text/TextMapping/ParagraphMapping.cs`**

```csharp
public class ParagraphMapping : AbstractMapping
{
    public ParagraphMapping(TextExtractionContext context) : base(context) { }
    
    protected override void ApplyCore(IVisitable x)
    {
        var paragraph = x as Paragraph;
        
        // Check if this paragraph should be processed based on context
        if (IsSpecialParagraph(paragraph))
        {
            ProcessSpecialParagraph(paragraph);
        }
        else
        {
            ProcessRegularParagraph(paragraph);
        }
    }
    
    private bool IsSpecialParagraph(Paragraph paragraph)
    {
        // Detect if paragraph is in header, footer, comment, etc.
        return paragraph.IsInHeaderFooter || 
               paragraph.IsComment || 
               paragraph.IsInTextBox;
    }
}
```

#### 4.2 Create Conditional Processing Helpers
**File: `Text/TextMapping/ConditionalProcessor.cs`** (create)

```csharp
public static class ConditionalProcessor
{
    public static void ProcessConditionally<T>(
        TextExtractionContext context,
        T item,
        Action<T> processor,
        Func<TextExtractionSettings, bool> shouldProcess)
        where T : class
    {
        if (shouldProcess(context.Settings))
        {
            processor(item);
        }
    }
    
    public static void ProcessWithState<T>(
        TextExtractionContext context,
        T item,
        Action<T> processor,
        Action<TextExtractionContext> enterState,
        Action<TextExtractionContext> exitState)
        where T : class
    {
        enterState(context);
        try
        {
            if (context.ShouldProcessCurrentContent())
            {
                processor(item);
            }
        }
        finally
        {
            exitState(context);
        }
    }
}
```

### Phase 5: Integration and Testing (1-2 days)

#### 5.1 Update Main Converter Class
**File: `Text/TextConverter.cs`**

```csharp
public static class TextConverter
{
    public static void Convert(
        WordDocument document, 
        TextWriter writer, 
        TextExtractionSettings settings = null)
    {
        settings = settings ?? TextExtractionSettings.Default;
        var textWriter = new TextWriter(writer);
        var context = new TextExtractionContext(settings, textWriter, document);
        
        var documentMapping = new DocumentMapping(context);
        documentMapping.Apply(document);
    }
    
    // Async version
    public static async Task ConvertAsync(
        WordDocument document, 
        TextWriter writer, 
        TextExtractionSettings settings = null,
        CancellationToken cancellationToken = default)
    {
        await Task.Run(() => Convert(document, writer, settings), cancellationToken);
    }
}
```

#### 5.2 Create Unit Tests
**File: `UnitTests/ConfigurableExtractionTests.cs`** (create)

```csharp
[TestFixture]
public class ConfigurableExtractionTests
{
    [Test]
    public void ShouldExcludeHeadersWhenConfigured()
    {
        var settings = new TextExtractionSettings
        {
            IncludeHeadersAndFooters = false
        };
        
        var result = ConvertWithSettings("samples/header_image.doc", settings);
        
        // Should not contain header content
        Assert.That(result, Does.Not.Contain("Header Text"));
    }
    
    [Test]
    public void ShouldExcludeTextBoxesWhenConfigured()
    {
        var settings = new TextExtractionSettings
        {
            IncludeTextBoxes = false
        };
        
        var result = ConvertWithSettings("samples/news-example.doc", settings);
        
        // Should not contain TextBox content
        Assert.That(result, Does.Not.Contain("Caixa de texto"));
    }
    
    [Test]
    public void ShouldExcludeCommentsWhenConfigured()
    {
        var settings = new TextExtractionSettings
        {
            IncludeComments = false
        };
        
        var result = ConvertWithSettings("samples/footnote_and_annotation.doc", settings);
        
        // Should not contain comment content
        Assert.That(result, Does.Not.Contain("[Comment:"));
    }
    
    private string ConvertWithSettings(string filePath, TextExtractionSettings settings)
    {
        using var doc = new WordDocument(filePath);
        using var writer = new StringWriter();
        
        TextConverter.Convert(doc, writer, settings);
        return writer.ToString();
    }
}
```

### Phase 6: Documentation and Help System (1 day)

#### 6.1 Enhanced Help System
**File: `Shell/doc2text/Program.cs`** (enhance)

```csharp
private static void ShowHelp()
{
    Console.WriteLine("doc2text - Convert Word documents to plain text");
    Console.WriteLine();
    Console.WriteLine("Usage: doc2text --input <input.doc> --output <output.txt> [options]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  --input <file>          Input Word document (.doc file)");
    Console.WriteLine("  --output <file>         Output text file");
    Console.WriteLine("  --no-headers-footers    Exclude headers and footers");
    Console.WriteLine("  --no-textboxes          Exclude text box content");
    Console.WriteLine("  --no-comments           Exclude comments and annotations");
    Console.WriteLine("  --no-footnotes          Exclude footnotes and endnotes");
    Console.WriteLine("  --help                  Show this help message");
    Console.WriteLine();
    Console.WriteLine("Examples:");
    Console.WriteLine("  doc2text --input document.doc --output document.txt");
    Console.WriteLine("  doc2text --input document.doc --output clean.txt --no-headers-footers --no-comments");
}
```

#### 6.2 Create Configuration File Support (Optional)
**File: `Text/ConfigurationFile.cs`** (create)

```csharp
public class ConfigurationFile
{
    public static TextExtractionSettings LoadFromFile(string configPath)
    {
        if (!File.Exists(configPath))
            return TextExtractionSettings.Default;
            
        var json = File.ReadAllText(configPath);
        return JsonSerializer.Deserialize<TextExtractionSettings>(json);
    }
    
    public static void SaveToFile(TextExtractionSettings settings, string configPath)
    {
        var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions 
        { 
            WriteIndented = true 
        });
        File.WriteAllText(configPath, json);
    }
}
```

## Implementation Timeline
- **Day 1-2**: Command-line infrastructure and argument parsing
- **Day 3-4**: Settings framework and extraction context
- **Day 5-7**: Text mapping updates and conditional processing
- **Day 8-9**: Integration, testing, and validation
- **Day 10**: Documentation and help system

## Success Metrics
1. All command-line options work as documented
2. Content exclusion works correctly for each option
3. Default behavior unchanged when no options specified
4. Performance impact < 5% when using selective extraction
5. Comprehensive test coverage for all configuration options

## Testing Commands
```bash
# Build the solution
dotnet build b2xtranslator.sln

# Test help system
dotnet run --project Shell/doc2text/doc2text.csproj -- --help

# Test default extraction (should include everything)
dotnet run --project Shell/doc2text/doc2text.csproj -- --input samples/news-example.doc --output output_full.txt

# Test selective extraction
dotnet run --project Shell/doc2text/doc2text.csproj -- --input samples/news-example.doc --output output_no_textbox.txt --no-textboxes

# Test multiple exclusions
dotnet run --project Shell/doc2text/doc2text.csproj -- --input samples/header_image.doc --output output_clean.txt --no-headers-footers --no-comments

# Verify content differences
diff output_full.txt output_no_textbox.txt  # Should show missing TextBox content
wc -w output_full.txt output_clean.txt     # Should show word count differences

# Run unit tests
dotnet test UnitTests/UnitTests.csproj --filter "Category=ConfigurableExtraction"
```
