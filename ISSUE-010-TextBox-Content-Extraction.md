# Issue #010: Text from TextBoxes is not Extracted

## Problem Description
Text contained within TextBoxes in Word documents is not being extracted during the text conversion process. This leads to incomplete output where significant content may be missing. The `README.md` identifies `news-example.doc` as a key example of this failure.

## Severity
**MEDIUM** - Can lead to significant content loss, but does not crash the application.

## Impact
- Incomplete text extraction, resulting in loss of important information.
- The context and meaning of the document can be lost if key text is in a textbox.
- Reduced reliability for documents that use textboxes for layout.

## Root Cause Analysis
TextBoxes are not part of the main document text stream. They are typically stored as "shapes" or Office Drawing objects. The current text extraction logic in `TextMapping` likely only traverses the primary content stream (`MainDocument` part) and does not parse these drawing objects to extract their text content.

1.  **Shape Processing**: The parser does not identify and process shape containers that hold text.
2.  **Text Traversal**: The text extraction visitor does not traverse into the drawing layer.
3.  **Data Location**: Textbox content is stored in a different part of the binary file that is not being read during text extraction.

## Known Affected Components
- **Text Mapping**: `Text/TextMapping/`
- **Doc Parser**: `Doc/DocFileFormat/`
- **Office Drawing Library**: `Common/OfficeDrawing/`

## Reproduction Steps

### Basic Reproduction
1.  Use the sample file `news-example.doc` (should be added to `samples-local/`).
2.  Run the text extraction tool:
    ```bash
    dotnet run --project Shell/doc2text/doc2text.csproj -- samples-local/news-example.doc output.txt
    ```
3.  Open `output.txt` and observe that the text from the text boxes is missing.

## Testing Strategy

### Integration Tests
- Add `news-example.doc` and other documents containing text boxes to the test samples.
- Create `.expected.txt` files that include the text from within the text boxes, preferably in the order they appear on the page.
- Add these files to the `SampleDocFileTextExtractionTests.cs` test suite.

## Proposed Solutions

### 1. Parse Drawing Objects
- Investigate how TextBox content is stored in the `.doc` format. This will likely involve the `OfficeArtContent` and related structures in the `Common/OfficeDrawing/` project.
- Extend the `WordDocument` parser to identify and extract drawing objects and their associated text runs.

### 2. Extend TextMapping
- Modify the `TextMapping` logic to traverse the collection of parsed shapes.
- When a shape containing text is found, extract its content and append it to the output stream.
- A key challenge will be to correctly order the textbox content relative to the main document text. A simple approach is to append all textbox content at the end. A more advanced solution would attempt to interleave it based on position.

## Files to Investigate
- `Text/TextMapping/TextMapping.cs`
- `Doc/DocFileFormat/WordDocument.cs`
- `Common/OfficeDrawing/` (entire directory)
- `Docx/WordprocessingMLMapping/DrawingMapping.cs` (for reference on how drawings are handled for DOCX conversion)

## Success Criteria
1.  **Text Extracted**: Text from simple TextBoxes is present in the extracted text output.
2.  **No Duplication**: Text is not extracted multiple times.
3.  **Order**: Text from textboxes is included in the output, ideally at the end of the document content to avoid complex layout calculations.

## Current Status (Verified)
**PROBLEM PERSISTS** - Testing shows that TextBox content is still missing from extracted text:
- Expected output includes: "Existe uma Caixa de texto aqui" (line 37)
- Actual output: Missing this TextBox content entirely
- The issue remains active and needs resolution

## Detailed Resolution Plan

### Phase 1: Investigation and Analysis (1-2 days)

#### 1.1 Understand TextBox Binary Storage
```bash
# Analyze the binary structure of news-example.doc
hexdump -C samples/news-example.doc | grep -A5 -B5 "Caixa\|texto"
```

**Tasks:**
- Study how TextBoxes are stored in Word binary format
- Examine the OfficeDrawing structures in `Common/OfficeDrawing/`
- Identify where TextBox text content is stored (likely in drawing objects)
- Map the binary structure to understand parsing requirements

#### 1.2 Code Flow Analysis
**Investigate current parsing flow:**
- `Doc/DocFileFormat/WordDocument.cs` - Document loading
- `Doc/DocFileFormat/OfficeArtContent.cs` - Drawing objects
- `Text/TextMapping/DocumentMapping.cs` - Text extraction logic

**Key Questions:**
- Are drawing objects being parsed at all?
- Is TextBox text stored in the main text stream or separate streams?
- Are there existing classes for TextBox representation?

### Phase 2: Parser Enhancement (2-3 days)

#### 2.1 Extend OfficeDrawing Parser
**File: `Common/OfficeDrawing/`**

Create new classes if needed:
```csharp
// New class for TextBox representation
public class TextBoxContainer : RegularContainer
{
    public string TextContent { get; set; }
    public Rectangle Bounds { get; set; }
}

// Enhance existing Shape class
public class Shape 
{
    // Add TextBox detection
    public bool IsTextBox => ShapeType == OfficeShapeType.TextBox;
    public string ExtractText() { /* Implementation */ }
}
```

#### 2.2 Enhance WordDocument Parser
**File: `Doc/DocFileFormat/WordDocument.cs`**

```csharp
// Add TextBox collection
public class WordDocument 
{
    public List<TextBoxContent> TextBoxes { get; set; } = new List<TextBoxContent>();
    
    // Add method to parse drawing objects for text
    private void ParseDrawingObjectText() 
    {
        // Implementation to extract text from OfficeArtContent
    }
}
```

### Phase 3: Text Mapping Integration (1-2 days)

#### 3.1 Update Text Mapping Logic
**File: `Text/TextMapping/DocumentMapping.cs`**

```csharp
public class DocumentMapping : AbstractOpenXmlMapping
{
    public override void Apply(IVisitable x)
    {
        var doc = x as WordDocument;
        
        // Existing main document processing...
        ProcessMainDocument(doc);
        
        // NEW: Process TextBox content
        ProcessTextBoxContent(doc);
    }
    
    private void ProcessTextBoxContent(WordDocument doc)
    {
        if (doc.TextBoxes?.Any() == true)
        {
            foreach (var textBox in doc.TextBoxes)
            {
                // Extract and append TextBox content
                _writer.WriteText(textBox.Content);
                _writer.WriteText(Environment.NewLine);
            }
        }
    }
}
```

#### 3.2 Configuration Support
**File: `Text/TextExtractionSettings.cs`** (create if needed)

```csharp
public class TextExtractionSettings
{
    public bool IncludeTextBoxes { get; set; } = true;
    public TextBoxPlacement TextBoxPlacement { get; set; } = TextBoxPlacement.AtEnd;
}

public enum TextBoxPlacement
{
    AtEnd,       // Append all TextBoxes at document end
    InPosition,  // Try to place based on document position
    Separate     // Extract to separate section
}
```

### Phase 4: Testing Implementation (1 day)

#### 4.1 Unit Tests
**File: `UnitTests/TextBoxExtractionTests.cs`** (create)

```csharp
[TestFixture]
public class TextBoxExtractionTests
{
    [Test]
    public void ShouldExtractTextBoxContent()
    {
        // Test TextBox parsing logic
        var doc = ParseDocument("samples/news-example.doc");
        Assert.That(doc.TextBoxes.Count, Is.GreaterThan(0));
        Assert.That(doc.TextBoxes[0].Content, Contains.Substring("Caixa de texto"));
    }
    
    [Test]
    public void ShouldIncludeTextBoxInOutput()
    {
        var result = ConvertToText("samples/news-example.doc");
        Assert.That(result, Contains.Substring("Existe uma Caixa de texto aqui"));
    }
}
```

#### 4.2 Integration Test Updates
Update `IntegrationTests/SampleDocFileTextExtractionTests.cs`:
- Ensure news-example.doc test passes with TextBox content
- Add verification for TextBox-specific test cases

### Phase 5: Advanced Features (Optional - 1-2 days)

#### 5.1 Position-Based TextBox Placement
```csharp
// Enhanced placement logic based on TextBox coordinates
private void ProcessTextBoxContentWithPosition(WordDocument doc)
{
    var sortedTextBoxes = doc.TextBoxes
        .OrderBy(tb => tb.Position.Y)  // Sort by vertical position
        .ThenBy(tb => tb.Position.X);  // Then by horizontal position
    
    // Integrate with main text flow based on position
}
```

#### 5.2 TextBox Metadata Preservation
```csharp
public class TextBoxContent
{
    public string Text { get; set; }
    public Rectangle Bounds { get; set; }
    public string OriginalFormat { get; set; }
    public int ZOrder { get; set; }  // Layering information
}
```

## Implementation Timeline
- **Day 1-2**: Investigation and binary analysis
- **Day 3-5**: Parser enhancement and integration
- **Day 6**: Testing and debugging
- **Day 7**: Optional advanced features and cleanup

## Success Metrics
1. `news-example.doc` conversion includes "Existe uma Caixa de texto aqui"
2. All existing tests continue to pass
3. New unit tests for TextBox extraction pass
4. Performance impact < 10% for documents without TextBoxes
5. Support for basic TextBox formatting preservation

## Testing Commands
```bash
# Build the solution
dotnet build b2xtranslator.sln

# Run integration tests
dotnet test IntegrationTests/IntegrationTests.csproj

# Test specific TextBox functionality
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc output.txt
grep "Caixa de texto" output.txt  # Should find the TextBox content

# Verify no regression on other samples
dotnet test IntegrationTests/IntegrationTests.csproj --filter "Category=SampleFiles"
```
