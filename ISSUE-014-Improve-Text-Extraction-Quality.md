# Issue #014: Improve General Text Extraction Quality

## Current Status
**A comprehensive resolution plan is now available.**

## Problem Description
This is a meta/tracking issue for the overall goal of improving the quality of plain text extraction. While specific bugs and features are tracked in other issues, this serves as an epic to group all efforts related to making the text output cleaner, more complete, and more accurate.

## Severity
**HIGH** - The core value proposition depends on extraction quality

## Impact
- The core value proposition of the `doc2text` tool is directly tied to the quality of its output.
- Low-quality output (with missing content, extra markers, or bad formatting) makes the tool less useful.
- Poor text extraction quality affects all downstream processing and analysis
- Competitive disadvantage compared to other text extraction tools
- Reduced user adoption and confidence

## Scope
This epic covers a range of improvements, many of which are tracked in their own issues:
- **List Handling**: `ISSUE-005-List-Extraction.md`
- **Field Markers**: `ISSUE-006-Internal-Field-Markers.md`
- **Missing Content**: `ISSUE-007-Missing-Paragraph-Content.md`
- **Whitespace**: `ISSUE-008-Alternative-Space-Characters.md`
- **TextBoxes**: `ISSUE-010-TextBox-Content-Extraction.md`
- **Symbols**: `ISSUE-011-Symbol-Handling.md`

The goal is to systematically address these issues to raise the overall quality bar.

## Root Cause Analysis
The root cause is a combination of many small factors: incomplete implementation of the spec, unhandled edge cases, and missing features in the `TextMapping` logic. Each sub-issue has its own detailed root cause analysis.

## Testing Strategy

### Golden Master Testing
- The primary testing strategy is "Golden Master" (or Approval) testing, which is already in place with the `samples/*.doc` and `samples/*.expected.txt` files.
- The goal is to continuously improve the `.actual.txt` output so that it matches the `.expected.txt` "golden master".

### Test Corpus Expansion
- To improve quality, the corpus of test documents must be expanded. We should actively seek out and add real-world documents that challenge the extractor.
- This includes documents with complex tables, mixed languages, right-to-left text, complex layouts, and various embedded objects.

## Proposed Solutions

### 1. Prioritize and Execute
- Systematically work through the backlog of linked issues.
- Prioritize issues based on their impact on text quality (e.g., missing paragraphs are higher priority than incorrect symbols).

### 2. Comparative Analysis
- Regularly compare the output of `doc2text` against other established tools (like `antiword`, `wvWare`, `catdoc`, or commercial libraries) on a large corpus of documents.
- Identify systematic gaps where `b2xtranslator` is weaker and create new issues to address them.

### 3. Refine Expected Output
- As the extractor improves, the `.expected.txt` files should be reviewed and refined. The goal is to make the expected output a clean, human-readable version of the document's text content.

## Comprehensive Resolution Plan

### Phase 1: Quality Assessment and Baseline Establishment (1-2 days)

#### 1.1 Create Quality Metrics Framework
**File: `IntegrationTests/QualityMetrics.cs`** (create)

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace b2xtranslator.IntegrationTests
{
    public class QualityMetrics
    {
        public class ExtractionQuality
        {
            public string DocumentName { get; set; }
            public int ExpectedWordCount { get; set; }
            public int ActualWordCount { get; set; }
            public double WordCountAccuracy => (double)ActualWordCount / ExpectedWordCount;
            
            public int ExpectedParagraphs { get; set; }
            public int ActualParagraphs { get; set; }
            public double ParagraphAccuracy => (double)ActualParagraphs / ExpectedParagraphs;
            
            public List<string> MissingContent { get; set; } = new List<string>();
            public List<string> ExtraContent { get; set; } = new List<string>();
            public List<string> FieldMarkers { get; set; } = new List<string>();
            
            public double OverallQualityScore { get; set; }
        }
        
        public static ExtractionQuality AnalyzeQuality(string expectedPath, string actualPath)
        {
            var expected = File.ReadAllText(expectedPath);
            var actual = File.ReadAllText(actualPath);
            
            var quality = new ExtractionQuality
            {
                DocumentName = Path.GetFileNameWithoutExtension(expectedPath),
                ExpectedWordCount = CountWords(expected),
                ActualWordCount = CountWords(actual),
                ExpectedParagraphs = CountParagraphs(expected),
                ActualParagraphs = CountParagraphs(actual)
            };
            
            // Detect field markers in output
            quality.FieldMarkers = DetectFieldMarkers(actual);
            
            // Calculate missing/extra content
            DetectContentDifferences(expected, actual, quality);
            
            // Calculate overall quality score (0-100)
            quality.OverallQualityScore = CalculateQualityScore(quality);
            
            return quality;
        }
        
        private static int CountWords(string text)
        {
            return Regex.Split(text.Trim(), @"\s+").Where(w => !string.IsNullOrEmpty(w)).Count();
        }
        
        private static int CountParagraphs(string text)
        {
            return text.Split(new[] { "\n\n", "\r\n\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        
        private static List<string> DetectFieldMarkers(string text)
        {
            var markers = new List<string>();
            
            // Common field markers to detect
            var patterns = new[]
            {
                @"\{\s*\\[a-zA-Z]+\s*\}",  // RTF-like markers
                @"MERGEFIELD\s+\w+",       // Mail merge fields
                @"HYPERLINK\s+",           // Hyperlink fields
                @"INCLUDEPICTURE\s+",      // Picture fields
                @"FORMTEXT\s*",            // Form fields
                @"\{\s*\d+\s*\}",         // Numbered placeholders
                @"«\s*\w+\s*»"            // French quote markers
            };
            
            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                markers.AddRange(matches.Cast<Match>().Select(m => m.Value));
            }
            
            return markers.Distinct().ToList();
        }
        
        private static void DetectContentDifferences(string expected, string actual, ExtractionQuality quality)
        {
            // Split into significant phrases for comparison
            var expectedPhrases = ExtractSignificantPhrases(expected);
            var actualPhrases = ExtractSignificantPhrases(actual);
            
            quality.MissingContent = expectedPhrases.Except(actualPhrases).Take(10).ToList();
            quality.ExtraContent = actualPhrases.Except(expectedPhrases).Take(10).ToList();
        }
        
        private static IEnumerable<string> ExtractSignificantPhrases(string text)
        {
            // Extract phrases of 3+ words that are likely to be meaningful content
            var sentences = Regex.Split(text, @"[.!?]+");
            var phrases = new List<string>();
            
            foreach (var sentence in sentences)
            {
                var words = Regex.Split(sentence.Trim(), @"\s+")
                    .Where(w => w.Length > 2 && !string.IsNullOrWhiteSpace(w))
                    .ToArray();
                    
                if (words.Length >= 3)
                {
                    phrases.Add(string.Join(" ", words.Take(5)).ToLowerInvariant());
                }
            }
            
            return phrases;
        }
        
        private static double CalculateQualityScore(ExtractionQuality quality)
        {
            var wordScore = Math.Min(quality.WordCountAccuracy, 1.0) * 40;  // 40 points max
            var paragraphScore = Math.Min(quality.ParagraphAccuracy, 1.0) * 20;  // 20 points max
            var contentScore = Math.Max(0, 30 - quality.MissingContent.Count * 3);  // 30 points max
            var cleanScore = Math.Max(0, 10 - quality.FieldMarkers.Count);  // 10 points max
            
            return wordScore + paragraphScore + contentScore + cleanScore;
        }
    }
}
```

#### 1.2 Create Quality Assessment Tool
**File: `Tools/QualityAssessment/QualityAssessment.cs`** (create)

```csharp
using System;
using System.IO;
using System.Linq;
using b2xtranslator.IntegrationTests;

namespace b2xtranslator.Tools.QualityAssessment
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: QualityAssessment <samples-directory>");
                return;
            }
            
            var samplesDir = args[0];
            var expectedFiles = Directory.GetFiles(samplesDir, "*.expected.txt");
            
            Console.WriteLine("=== Text Extraction Quality Report ===");
            Console.WriteLine($"Analyzing {expectedFiles.Length} documents...\n");
            
            var allQualities = new List<QualityMetrics.ExtractionQuality>();
            
            foreach (var expectedFile in expectedFiles)
            {
                var actualFile = expectedFile.Replace(".expected.txt", ".actual.txt");
                
                if (!File.Exists(actualFile))
                {
                    Console.WriteLine($"⚠ Missing: {Path.GetFileName(actualFile)}");
                    continue;
                }
                
                var quality = QualityMetrics.AnalyzeQuality(expectedFile, actualFile);
                allQualities.Add(quality);
                
                PrintDocumentQuality(quality);
            }
            
            PrintOverallSummary(allQualities);
        }
        
        private static void PrintDocumentQuality(QualityMetrics.ExtractionQuality quality)
        {
            var status = quality.OverallQualityScore >= 80 ? "✅" : 
                        quality.OverallQualityScore >= 60 ? "⚠" : "❌";
            
            Console.WriteLine($"{status} {quality.DocumentName} (Score: {quality.OverallQualityScore:F1}/100)");
            Console.WriteLine($"   Words: {quality.ActualWordCount}/{quality.ExpectedWordCount} ({quality.WordCountAccuracy:P1})");
            Console.WriteLine($"   Paragraphs: {quality.ActualParagraphs}/{quality.ExpectedParagraphs} ({quality.ParagraphAccuracy:P1})");
            
            if (quality.FieldMarkers.Any())
            {
                Console.WriteLine($"   Field Markers: {string.Join(", ", quality.FieldMarkers.Take(3))}");
            }
            
            if (quality.MissingContent.Any())
            {
                Console.WriteLine($"   Missing: {string.Join("; ", quality.MissingContent.Take(2))}...");
            }
            
            Console.WriteLine();
        }
        
        private static void PrintOverallSummary(List<QualityMetrics.ExtractionQuality> qualities)
        {
            var avgScore = qualities.Average(q => q.OverallQualityScore);
            var excellentCount = qualities.Count(q => q.OverallQualityScore >= 80);
            var goodCount = qualities.Count(q => q.OverallQualityScore >= 60 && q.OverallQualityScore < 80);
            var poorCount = qualities.Count(q => q.OverallQualityScore < 60);
            
            Console.WriteLine("=== Overall Summary ===");
            Console.WriteLine($"Average Quality Score: {avgScore:F1}/100");
            Console.WriteLine($"Excellent (80-100): {excellentCount} documents");
            Console.WriteLine($"Good (60-79): {goodCount} documents");
            Console.WriteLine($"Poor (<60): {poorCount} documents");
            
            // Identify most common issues
            var allFieldMarkers = qualities.SelectMany(q => q.FieldMarkers).ToList();
            var commonMarkers = allFieldMarkers.GroupBy(m => m)
                .OrderByDescending(g => g.Count())
                .Take(5);
            
            if (commonMarkers.Any())
            {
                Console.WriteLine("\nMost Common Field Markers:");
                foreach (var marker in commonMarkers)
                {
                    Console.WriteLine($"  {marker.Key}: {marker.Count()} occurrences");
                }
            }
        }
    }
}
```

### Phase 2: Systematic Quality Improvements (2-3 days)

#### 2.1 Enhanced Content Preservation Framework
**File: `Text/TextMapping/ContentPreservationMapping.cs`** (create)

```csharp
public class ContentPreservationMapping : AbstractMapping
{
    private readonly ContentPreservationSettings _settings;
    private readonly List<string> _preservedElements = new List<string>();
    
    public ContentPreservationMapping(ConversionContext context, ContentPreservationSettings settings) 
        : base(context)
    {
        _settings = settings;
    }
    
    public override void Apply(IVisitable x)
    {
        // Track all content elements to ensure nothing is lost
        if (x is IContentElement element)
        {
            ProcessContentElement(element);
        }
    }
    
    private void ProcessContentElement(IContentElement element)
    {
        try
        {
            // Attempt to extract content with multiple strategies
            var content = ExtractContentWithFallback(element);
            
            if (!string.IsNullOrEmpty(content))
            {
                _preservedElements.Add(content);
                WriteContent(content);
            }
            else if (_settings.LogMissingContent)
            {
                LogMissingContent(element);
            }
        }
        catch (Exception ex)
        {
            if (_settings.ContinueOnError)
            {
                LogContentError(element, ex);
            }
            else
            {
                throw;
            }
        }
    }
    
    private string ExtractContentWithFallback(IContentElement element)
    {
        // Strategy 1: Standard extraction
        try
        {
            return element.ExtractText();
        }
        catch
        {
            // Strategy 2: Raw content extraction
            try
            {
                return element.ExtractRawContent();
            }
            catch
            {
                // Strategy 3: Property-based extraction
                return ExtractFromProperties(element);
            }
        }
    }
}

public class ContentPreservationSettings
{
    public bool LogMissingContent { get; set; } = true;
    public bool ContinueOnError { get; set; } = true;
    public bool PreservePlaceholderText { get; set; } = false;
    public bool ExtractAltText { get; set; } = true;
    public bool PreserveHiddenText { get; set; } = false;
}
```

#### 2.2 Advanced Text Cleaning Pipeline
**File: `Text/TextMapping/TextCleaningPipeline.cs`** (create)

```csharp
public class TextCleaningPipeline
{
    private readonly List<ITextCleaner> _cleaners = new List<ITextCleaner>();
    
    public TextCleaningPipeline()
    {
        // Register default cleaners in order of execution
        _cleaners.Add(new FieldMarkerCleaner());
        _cleaners.Add(new WhitespaceCleaner());
        _cleaners.Add(new UnicodeNormalizer());
        _cleaners.Add(new ParagraphFormatter());
        _cleaners.Add(new ListFormatter());
    }
    
    public string CleanText(string input, TextCleaningSettings settings)
    {
        var result = input;
        
        foreach (var cleaner in _cleaners)
        {
            if (cleaner.ShouldApply(settings))
            {
                result = cleaner.Clean(result, settings);
            }
        }
        
        return result;
    }
}

public interface ITextCleaner
{
    bool ShouldApply(TextCleaningSettings settings);
    string Clean(string input, TextCleaningSettings settings);
}

public class FieldMarkerCleaner : ITextCleaner
{
    public bool ShouldApply(TextCleaningSettings settings) => settings.RemoveFieldMarkers;
    
    public string Clean(string input, TextCleaningSettings settings)
    {
        // Remove common field markers
        var patterns = new[]
        {
            @"\{\s*\\[a-zA-Z]+\s*\}",           // RTF-like markers
            @"MERGEFIELD\s+\w+(\s+\\[^}]*)?",   // Mail merge fields
            @"HYPERLINK\s+""[^""]*""",          // Hyperlink fields
            @"INCLUDEPICTURE\s+""[^""]*""",     // Picture fields
            @"FORMTEXT\s*",                      // Form fields
            @"«\s*[^»]*\s*»",                   // French quote markers
            @"\{\s*\d+\s*\}",                   // Numbered placeholders
            @"\\[a-z]+\d*\s*"                   // Other RTF-like commands
        };
        
        var result = input;
        foreach (var pattern in patterns)
        {
            result = Regex.Replace(result, pattern, "", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        }
        
        return result;
    }
}

public class WhitespaceCleaner : ITextCleaner
{
    public bool ShouldApply(TextCleaningSettings settings) => settings.NormalizeWhitespace;
    
    public string Clean(string input, TextCleaningSettings settings)
    {
        // Normalize various whitespace characters
        var result = input;
        
        // Replace non-breaking spaces with regular spaces
        result = result.Replace('\u00A0', ' ');  // Non-breaking space
        result = result.Replace('\u2007', ' ');  // Figure space
        result = result.Replace('\u202F', ' ');  // Narrow non-breaking space
        
        // Normalize line endings
        result = Regex.Replace(result, @"\r\n|\r|\n", "\n");
        
        // Remove excessive whitespace but preserve paragraph breaks
        result = Regex.Replace(result, @"[ \t]+", " ");           // Multiple spaces/tabs → single space
        result = Regex.Replace(result, @"\n{3,}", "\n\n");       // Multiple line breaks → double line break
        result = Regex.Replace(result, @"[ \t]+\n", "\n");       // Trailing whitespace on lines
        
        return result.Trim();
    }
}

public class ListFormatter : ITextCleaner
{
    public bool ShouldApply(TextCleaningSettings settings) => settings.FormatLists;
    
    public string Clean(string input, TextCleaningSettings settings)
    {
        // Improve list formatting by ensuring proper bullet points and numbering
        var lines = input.Split('\n');
        var result = new StringBuilder();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            
            // Detect and format list items
            if (IsListItem(line))
            {
                var formattedLine = FormatListItem(line, settings);
                result.AppendLine(formattedLine);
            }
            else
            {
                result.AppendLine(line);
            }
        }
        
        return result.ToString();
    }
    
    private bool IsListItem(string line)
    {
        // Detect various list item patterns
        return Regex.IsMatch(line, @"^\s*[-*•·]\s+") ||           // Bullet points
               Regex.IsMatch(line, @"^\s*\d+[.)]\s+") ||          // Numbered lists
               Regex.IsMatch(line, @"^\s*[a-zA-Z][.)]\s+");       // Lettered lists
    }
    
    private string FormatListItem(string line, TextCleaningSettings settings)
    {
        // Ensure consistent bullet character
        line = Regex.Replace(line, @"^(\s*)[-*·]\s+", $"$1{settings.BulletCharacter} ");
        
        // Ensure consistent numbering format
        line = Regex.Replace(line, @"^(\s*)(\d+)[.)]\s+", "$1$2. ");
        
        return line;
    }
}

public class TextCleaningSettings
{
    public bool RemoveFieldMarkers { get; set; } = true;
    public bool NormalizeWhitespace { get; set; } = true;
    public bool FormatLists { get; set; } = true;
    public bool NormalizeUnicode { get; set; } = true;
    public string BulletCharacter { get; set; } = "•";
    public bool PreserveParagraphStructure { get; set; } = true;
}
```

### Phase 3: Advanced Content Detection and Extraction (2-3 days)

#### 3.1 Content Classification System
**File: `Text/TextMapping/ContentClassifier.cs`** (create)

```csharp
public enum ContentType
{
    MainText,
    Header,
    Footer,
    Table,
    List,
    TextBox,
    Comment,
    Footnote,
    Hyperlink,
    Field,
    Image,
    Chart,
    Symbol,
    FormField
}

public class ContentClassifier
{
    private readonly Dictionary<ContentType, IContentDetector> _detectors;
    
    public ContentClassifier()
    {
        _detectors = new Dictionary<ContentType, IContentDetector>
        {
            { ContentType.Header, new HeaderDetector() },
            { ContentType.Footer, new FooterDetector() },
            { ContentType.Table, new TableDetector() },
            { ContentType.List, new ListDetector() },
            { ContentType.TextBox, new TextBoxDetector() },
            { ContentType.Comment, new CommentDetector() },
            { ContentType.Footnote, new FootnoteDetector() },
            { ContentType.Hyperlink, new HyperlinkDetector() },
            { ContentType.Field, new FieldDetector() },
            { ContentType.Symbol, new SymbolDetector() }
        };
    }
    
    public ContentType ClassifyContent(IVisitable content)
    {
        foreach (var detector in _detectors)
        {
            if (detector.Value.CanHandle(content))
            {
                return detector.Key;
            }
        }
        
        return ContentType.MainText;
    }
    
    public ExtractionStrategy GetExtractionStrategy(ContentType contentType)
    {
        return contentType switch
        {
            ContentType.MainText => new StandardTextExtraction(),
            ContentType.Table => new TableExtractionStrategy(),
            ContentType.List => new ListExtractionStrategy(),
            ContentType.TextBox => new TextBoxExtractionStrategy(),
            ContentType.Symbol => new SymbolExtractionStrategy(),
            ContentType.Field => new FieldExtractionStrategy(),
            _ => new StandardTextExtraction()
        };
    }
}

public interface IContentDetector
{
    bool CanHandle(IVisitable content);
}

public abstract class ExtractionStrategy
{
    public abstract string ExtractText(IVisitable content, ExtractionContext context);
}
```

#### 3.2 Smart Content Reconstruction
**File: `Text/TextMapping/ContentReconstructor.cs`** (create)

```csharp
public class ContentReconstructor
{
    private readonly ContentClassifier _classifier;
    private readonly TextCleaningPipeline _cleaningPipeline;
    
    public ContentReconstructor()
    {
        _classifier = new ContentClassifier();
        _cleaningPipeline = new TextCleaningPipeline();
    }
    
    public ReconstructedDocument ReconstructDocument(WordDocument document, ReconstructionSettings settings)
    {
        var result = new ReconstructedDocument();
        var context = new ExtractionContext(settings);
        
        // Phase 1: Extract and classify all content
        var contentItems = ExtractAllContent(document, context);
        
        // Phase 2: Organize content by type and importance
        var organizedContent = OrganizeContent(contentItems);
        
        // Phase 3: Reconstruct document structure
        result.MainContent = ReconstructMainContent(organizedContent);
        result.Metadata = ExtractMetadata(organizedContent);
        result.QualityMetrics = CalculateQualityMetrics(result);
        
        return result;
    }
    
    private List<ClassifiedContent> ExtractAllContent(WordDocument document, ExtractionContext context)
    {
        var content = new List<ClassifiedContent>();
        
        // Extract content from all document parts
        foreach (var part in document.GetAllDocumentParts())
        {
            foreach (var element in part.GetContentElements())
            {
                var contentType = _classifier.ClassifyContent(element);
                var strategy = _classifier.GetExtractionStrategy(contentType);
                var text = strategy.ExtractText(element, context);
                
                if (!string.IsNullOrEmpty(text))
                {
                    content.Add(new ClassifiedContent
                    {
                        Type = contentType,
                        Text = text,
                        SourceElement = element,
                        Position = element.GetDocumentPosition(),
                        Priority = GetContentPriority(contentType)
                    });
                }
            }
        }
        
        return content.OrderBy(c => c.Position).ToList();
    }
    
    private string ReconstructMainContent(OrganizedContent organized)
    {
        var builder = new StringBuilder();
        
        // Add title if available
        if (!string.IsNullOrEmpty(organized.Title))
        {
            builder.AppendLine(organized.Title);
            builder.AppendLine();
        }
        
        // Add main text content with proper formatting
        foreach (var paragraph in organized.MainText)
        {
            builder.AppendLine(paragraph);
        }
        
        // Add tables with simplified formatting
        foreach (var table in organized.Tables)
        {
            builder.AppendLine();
            builder.AppendLine(FormatTableAsText(table));
            builder.AppendLine();
        }
        
        // Add footnotes at the end if present
        if (organized.Footnotes.Any())
        {
            builder.AppendLine();
            builder.AppendLine("--- Footnotes ---");
            foreach (var footnote in organized.Footnotes)
            {
                builder.AppendLine(footnote);
            }
        }
        
        // Clean and format the final result
        return _cleaningPipeline.CleanText(builder.ToString(), new TextCleaningSettings());
    }
}

public class ReconstructedDocument
{
    public string MainContent { get; set; }
    public DocumentMetadata Metadata { get; set; }
    public QualityMetrics QualityMetrics { get; set; }
}

public class DocumentMetadata
{
    public string Title { get; set; }
    public string Author { get; set; }
    public DateTime? CreationDate { get; set; }
    public int WordCount { get; set; }
    public int ParagraphCount { get; set; }
    public List<string> Languages { get; set; } = new List<string>();
}
```

### Phase 4: Quality Assurance and Testing Framework (1-2 days)

#### 4.1 Automated Quality Regression Tests
**File: `IntegrationTests/QualityRegressionTests.cs`** (create)

```csharp
[TestFixture]
public class QualityRegressionTests
{
    private static readonly string SamplesDirectory = Path.Combine(TestContext.CurrentContext.TestDirectory, "samples");
    
    [Test]
    [TestCaseSource(nameof(GetQualityTestCases))]
    public void ShouldMaintainMinimumQualityScore(string documentPath, double minimumScore)
    {
        var actualPath = GenerateActualOutput(documentPath);
        var expectedPath = documentPath.Replace(".doc", ".expected.txt");
        
        var quality = QualityMetrics.AnalyzeQuality(expectedPath, actualPath);
        
        Assert.That(quality.OverallQualityScore, Is.GreaterThanOrEqualTo(minimumScore),
            $"Quality score {quality.OverallQualityScore:F1} is below minimum {minimumScore:F1} for {quality.DocumentName}");
    }
    
    [Test]
    public void ShouldNotContainFieldMarkers()
    {
        var testDocuments = GetCriticalTestDocuments();
        
        foreach (var doc in testDocuments)
        {
            var actualPath = GenerateActualOutput(doc);
            var content = File.ReadAllText(actualPath);
            
            var fieldMarkers = QualityMetrics.DetectFieldMarkers(content);
            
            Assert.That(fieldMarkers, Is.Empty,
                $"Document {Path.GetFileName(doc)} contains field markers: {string.Join(", ", fieldMarkers)}");
        }
    }
    
    [Test]
    public void ShouldPreserveMainContent()
    {
        var testCases = new[]
        {
            ("samples/simple.doc", new[] { "This is a simple document", "with some text" }),
            ("samples/news-example.doc", new[] { "News Example", "important information" }),
            ("samples/table test.doc", new[] { "table", "cells", "content" })
        };
        
        foreach (var (docPath, expectedPhrases) in testCases)
        {
            var actualPath = GenerateActualOutput(docPath);
            var content = File.ReadAllText(actualPath).ToLowerInvariant();
            
            foreach (var phrase in expectedPhrases)
            {
                Assert.That(content, Contains.Substring(phrase.ToLowerInvariant()),
                    $"Document {Path.GetFileName(docPath)} missing expected phrase: '{phrase}'");
            }
        }
    }
    
    private static IEnumerable<TestCaseData> GetQualityTestCases()
    {
        var qualityBaseline = new Dictionary<string, double>
        {
            { "simple.doc", 95.0 },
            { "news-example.doc", 85.0 },
            { "table test.doc", 80.0 },
            { "footnote_and_annotation.doc", 75.0 },
            { "header_image.doc", 70.0 },
            { "complex-layout.doc", 65.0 }
        };
        
        foreach (var baseline in qualityBaseline)
        {
            var docPath = Path.Combine(SamplesDirectory, baseline.Key);
            if (File.Exists(docPath))
            {
                yield return new TestCaseData(docPath, baseline.Value).SetName($"Quality_{baseline.Key}");
            }
        }
    }
}
```

#### 4.2 Performance and Quality Monitoring
**File: `Tools/QualityMonitoring/QualityMonitor.cs`** (create)

```csharp
public class QualityMonitor
{
    private readonly string _baselineDirectory;
    private readonly QualityMetrics _metrics;
    
    public QualityMonitor(string baselineDirectory)
    {
        _baselineDirectory = baselineDirectory;
        _metrics = new QualityMetrics();
    }
    
    public QualityReport GenerateQualityReport(string samplesDirectory)
    {
        var report = new QualityReport
        {
            GeneratedAt = DateTime.UtcNow,
            SamplesDirectory = samplesDirectory
        };
        
        var expectedFiles = Directory.GetFiles(samplesDirectory, "*.expected.txt");
        
        foreach (var expectedFile in expectedFiles)
        {
            var actualFile = expectedFile.Replace(".expected.txt", ".actual.txt");
            
            if (File.Exists(actualFile))
            {
                var quality = _metrics.AnalyzeQuality(expectedFile, actualFile);
                report.DocumentQualities.Add(quality);
                
                // Compare with baseline if available
                var baseline = LoadBaseline(quality.DocumentName);
                if (baseline != null)
                {
                    var regression = DetectRegression(quality, baseline);
                    if (regression != null)
                    {
                        report.Regressions.Add(regression);
                    }
                }
                
                // Update baseline
                SaveBaseline(quality);
            }
        }
        
        CalculateSummaryStatistics(report);
        return report;
    }
    
    private QualityRegression DetectRegression(QualityMetrics.ExtractionQuality current, QualityMetrics.ExtractionQuality baseline)
    {
        const double regressionThreshold = 5.0; // 5 point decrease considered regression
        
        var scoreDiff = current.OverallQualityScore - baseline.OverallQualityScore;
        
        if (scoreDiff < -regressionThreshold)
        {
            return new QualityRegression
            {
                DocumentName = current.DocumentName,
                CurrentScore = current.OverallQualityScore,
                BaselineScore = baseline.OverallQualityScore,
                ScoreDifference = scoreDiff,
                RegressionType = ClassifyRegression(current, baseline)
            };
        }
        
        return null;
    }
    
    private RegressionType ClassifyRegression(QualityMetrics.ExtractionQuality current, QualityMetrics.ExtractionQuality baseline)
    {
        if (current.FieldMarkers.Count > baseline.FieldMarkers.Count)
            return RegressionType.FieldMarkersIncreased;
        
        if (current.WordCountAccuracy < baseline.WordCountAccuracy - 0.1)
            return RegressionType.ContentLoss;
        
        if (current.MissingContent.Count > baseline.MissingContent.Count)
            return RegressionType.MissingContent;
        
        return RegressionType.GeneralQualityDecrease;
    }
}

public class QualityReport
{
    public DateTime GeneratedAt { get; set; }
    public string SamplesDirectory { get; set; }
    public List<QualityMetrics.ExtractionQuality> DocumentQualities { get; set; } = new List<QualityMetrics.ExtractionQuality>();
    public List<QualityRegression> Regressions { get; set; } = new List<QualityRegression>();
    
    public double AverageQualityScore { get; set; }
    public int ExcellentDocuments { get; set; }
    public int GoodDocuments { get; set; }
    public int PoorDocuments { get; set; }
}

public class QualityRegression
{
    public string DocumentName { get; set; }
    public double CurrentScore { get; set; }
    public double BaselineScore { get; set; }
    public double ScoreDifference { get; set; }
    public RegressionType RegressionType { get; set; }
}

public enum RegressionType
{
    ContentLoss,
    FieldMarkersIncreased,
    MissingContent,
    GeneralQualityDecrease
}
```

## Implementation Timeline
- **Day 1-2**: Quality metrics framework and assessment baseline
- **Day 3-4**: Content preservation and text cleaning pipeline
- **Day 5-7**: Advanced content detection and smart reconstruction
- **Day 8-9**: Quality assurance framework and regression testing
- **Day 10**: Integration, optimization, and documentation

## Success Criteria
1.  **High Fidelity**: The extracted text is a complete and accurate representation of the source document's content.
2.  **Clean Output**: The output contains no internal markers, field codes, or other conversion artifacts.
3.  **Good Formatting**: Paragraph breaks, table layouts (even if simplified), and list structures are preserved in a readable way.
4.  **Competitive Quality**: The output quality is on par with or better than other leading open-source text extraction tools.
5.  **Quality Metrics**: Average quality score across all sample documents ≥ 80/100
6.  **Regression Prevention**: Automated tests prevent quality degradation
7.  **Content Preservation**: ≥ 95% of main document content preserved

## Testing Commands
```bash
# Build the solution
dotnet build b2xtranslator.sln

# Run quality assessment
dotnet run --project Tools/QualityAssessment/QualityAssessment.csproj -- samples/

# Generate actual output for all samples
for doc in samples/*.doc; do
    base=$(basename "$doc" .doc)
    dotnet run --project Shell/doc2text/doc2text.csproj -- "$doc" "samples/${base}.actual.txt"
done

# Run quality regression tests
dotnet test IntegrationTests/IntegrationTests.csproj --filter "Category=QualityRegression"

# The main measure of success is the passing of the integration test suite
dotnet test IntegrationTests/IntegrationTests.csproj

# Generate comprehensive quality report
dotnet run --project Tools/QualityMonitoring/QualityMonitor.csproj -- samples/ --output quality-report.html

# Compare with baseline
dotnet run --project Tools/QualityAssessment/QualityAssessment.csproj -- samples/ --compare-baseline

# Performance and quality benchmarking
dotnet run --project Tools/Benchmarking/QualityBenchmark.csproj -- samples/ --iterations 10
```
