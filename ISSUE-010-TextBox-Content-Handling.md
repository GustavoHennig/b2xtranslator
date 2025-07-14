# Issue #010: TextBox Content Handling

## Problem Description
The b2xtranslator fails to properly extract and include content from TextBox elements during document conversion, resulting in missing text content that was contained within text boxes in the original Word document. This affects document completeness and information preservation.

## Severity
**MEDIUM** - Affects content completeness but doesn't prevent basic conversion

## Impact
- Loss of content contained within text boxes
- Incomplete document conversion and information preservation
- Missing important information that was intentionally placed in text boxes
- Poor document fidelity when text boxes contain significant content
- Accessibility issues for users who need all document content
- Inconsistent conversion results for documents with complex layouts

## Root Cause Analysis
TextBox content handling issues typically stem from:

1. **Missing TextBox Parsing**: TextBox objects not being recognized or processed during parsing
2. **Drawing Object Limitations**: TextBoxes treated as drawing objects rather than text containers
3. **Content Extraction Gaps**: TextBox text content not extracted during conversion process
4. **Mapping Incomplete**: Text mapping logic doesn't include TextBox content processing
5. **Binary Structure Parsing**: TextBox binary structures not properly interpreted

## Known Affected Files
Based on README.md mentions:
- `news-example.doc` - "Handle TextBox content: 'news-example.doc'"

This indicates that the document contains TextBox elements with content that should be extracted but currently isn't being processed.

## Related Context
The README.md also mentions a roadmap item to make the text extraction shell app accept extraction settings such as:
- `--no-textboxes` - Option to ignore TextBox elements

This suggests that TextBox handling should be:
1. Implemented for content extraction
2. Configurable to allow inclusion/exclusion of TextBox content

## Reproduction Steps

### Basic TextBox Content Testing
1. Build the solution:
   ```bash
   dotnet build b2xtranslator.sln
   ```

2. Test with known TextBox-containing file:
   ```bash
   # Test news-example.doc which contains TextBox content
   dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc output.txt
   cat output.txt  # Check if TextBox content is included
   ```

3. Compare with expected output:
   ```bash
   # Compare with expected output if available
   if [ -f samples/news-example.expected.txt ]; then
       diff samples/news-example.expected.txt output.txt
       echo "Expected content length: $(wc -w < samples/news-example.expected.txt)"
       echo "Actual content length: $(wc -w < output.txt)"
   fi
   ```

### Systematic TextBox Testing
```bash
# Test for TextBox content across sample files
textbox_content_test() {
    echo "Testing TextBox content extraction..."
    
    # Test news-example.doc specifically
    if [ -f "samples/news-example.doc" ]; then
        echo "=== Testing news-example.doc ==="
        
        dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc temp-textbox.txt
        
        echo "Content analysis:"
        echo "  Total characters: $(wc -c < temp-textbox.txt)"
        echo "  Total words: $(wc -w < temp-textbox.txt)"
        echo "  Total lines: $(wc -l < temp-textbox.txt)"
        
        # Check for potential TextBox indicators
        echo "Potential TextBox content indicators:"
        grep -n "textbox\|text box\|box\|frame" temp-textbox.txt -i || echo "  No TextBox indicators found"
        
        # Sample content
        echo "Sample content (first 5 lines):"
        head -5 temp-textbox.txt
        
        rm -f temp-textbox.txt
        echo ""
    fi
    
    # Test other potential TextBox-containing files
    local potential_textbox_files=(
        "news-example.doc"
        "footnote_and_annotation.doc"
        "header_image.doc"
        "FloatingPictures.doc"
    )
    
    for file in "${potential_textbox_files[@]}"; do
        if [ -f "samples/$file" ]; then
            local basename=$(basename "$file" .doc)
            echo "Testing $basename for TextBox content..."
            
            dotnet run --project Shell/doc2text/doc2text.csproj -- "samples/$file" "temp-$basename.txt"
            
            # Check content completeness
            if [ -f "samples/$basename.expected.txt" ]; then
                local expected_words=$(wc -w < "samples/$basename.expected.txt")
                local actual_words=$(wc -w < "temp-$basename.txt")
                
                if [ $expected_words -gt 0 ]; then
                    local percentage=$((actual_words * 100 / expected_words))
                    echo "  Content preservation: $percentage% ($actual_words/$expected_words words)"
                    
                    if [ $percentage -lt 80 ]; then
                        echo "  WARNING: Significant content loss - may include missing TextBox content"
                    fi
                fi
            else
                echo "  Generated $(wc -w < "temp-$basename.txt") words (no expected output for comparison)"
            fi
            
            rm -f "temp-$basename.txt"
        fi
    done
}

textbox_content_test
```

### TextBox Detection Analysis
```bash
# Analyze documents for TextBox structures
analyze_textbox_structures() {
    echo "Analyzing TextBox structures in sample documents..."
    
    # Binary analysis for TextBox markers
    if [ -f "samples/news-example.doc" ]; then
        echo "=== Binary analysis of news-example.doc ==="
        
        # Look for TextBox-related binary markers
        strings samples/news-example.doc | grep -i "textbox\|text.*box\|shape\|frame" | head -10
        
        # Look for drawing object indicators
        hexdump -C samples/news-example.doc | grep -E "textbox|shape|draw" -i | head -5
        
        echo ""
    fi
    
    # Test with reference tools for comparison
    if command -v antiword >/dev/null; then
        echo "Reference comparison with antiword:"
        antiword samples/news-example.doc > antiword-textbox.txt 2>/dev/null
        echo "  Antiword output words: $(wc -w < antiword-textbox.txt)"
        
        # Check if antiword extracts TextBox content
        grep -i "textbox\|text box" antiword-textbox.txt || echo "  No TextBox indicators in antiword output"
        rm -f antiword-textbox.txt
    fi
    
    if command -v catdoc >/dev/null; then
        echo "Reference comparison with catdoc:"
        catdoc samples/news-example.doc > catdoc-textbox.txt 2>/dev/null
        echo "  Catdoc output words: $(wc -w < catdoc-textbox.txt)"
        
        # Check if catdoc extracts TextBox content
        grep -i "textbox\|text box" catdoc-textbox.txt || echo "  No TextBox indicators in catdoc output"
        rm -f catdoc-textbox.txt
    fi
}

analyze_textbox_structures
```

## Debug Information to Collect

### TextBox Structure Analysis
- **TextBox Objects**: Identify TextBox objects in document binary structure
- **Content Location**: Where TextBox text content is stored in binary format
- **Drawing Objects**: Relationship between TextBoxes and drawing objects
- **Text Streams**: Whether TextBox content is in main text stream or separate streams

### Document Analysis
- **TextBox Count**: Number of TextBox elements in problematic documents
- **Content Volume**: Amount of text contained within TextBoxes
- **Layout Context**: How TextBoxes relate to main document flow
- **Formatting Information**: TextBox formatting and styling properties

### Conversion Process Analysis
- **Current Processing**: How TextBox objects are currently handled (if at all)
- **Missing Steps**: Identify where TextBox content extraction should occur
- **Binary Parsing**: Whether TextBox binary structures are being parsed
- **Text Extraction**: Whether TextBox text is being extracted during conversion

### Testing Commands
```bash
# Enable detailed TextBox processing logging
export DOTNET_LOGGING_LEVEL=Debug
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc debug-textbox.txt > textbox-debug.log 2>&1

# Analyze binary structure for TextBox data
strings samples/news-example.doc | grep -i "text" > textbox-strings.log
hexdump -C samples/news-example.doc | grep -E "text|box|shape" -i > textbox-hex.log

# Look for TextBox-related terms in conversion output
grep -i "textbox\|text.*box\|shape\|frame\|drawing" debug-textbox.txt textbox-debug.log
```

## Testing Strategy

### Unit Tests for TextBox Processing
```csharp
[Test]
public void ShouldExtractTextBoxContent()
{
    var converter = new DocumentConverter();
    var result = converter.ConvertToText("samples/news-example.doc");
    
    // Should contain substantial content including TextBox text
    Assert.That(result.Length, Is.GreaterThan(100), "Should extract TextBox content");
    Assert.That(result.Split(' ').Length, Is.GreaterThan(20), "Should have significant word count including TextBoxes");
}

[Test]
public void ShouldHandleDocumentsWithTextBoxes()
{
    var converter = new DocumentConverter();
    var result = converter.ConvertToText("samples/news-example.doc");
    
    // Should not be empty and should handle TextBoxes gracefully
    Assert.That(result, Is.Not.Empty, "Should extract content from documents with TextBoxes");
    Assert.That(result, Does.Not.Contain("ERROR"), "Should handle TextBoxes without errors");
}

[Test]
public void TextBoxContentShouldMatchExpected()
{
    if (File.Exists("samples/news-example.expected.txt"))
    {
        var converter = new DocumentConverter();
        var result = converter.ConvertToText("samples/news-example.doc");
        var expected = File.ReadAllText("samples/news-example.expected.txt");
        
        // Content should be reasonably close to expected
        var actualWords = result.Split(new[] { ' ', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
        var expectedWords = expected.Split(new[] { ' ', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
        var percentage = (double)actualWords / expectedWords * 100;
        
        Assert.That(percentage, Is.GreaterThan(80), $"Should extract most content including TextBoxes. Got {percentage:F1}%");
    }
}
```

### Integration TextBox Tests
```bash
# Automated TextBox content testing
dotnet test IntegrationTests/IntegrationTests.csproj --filter "Category=TextBoxContent"

# TextBox extraction validation
validate_textbox_extraction() {
    echo "Validating TextBox content extraction..."
    
    local test_files=("news-example.doc")
    local total_tests=0
    local passed_tests=0
    
    for file in "${test_files[@]}"; do
        if [ -f "samples/$file" ]; then
            ((total_tests++))
            local basename=$(basename "$file" .doc)
            
            echo "Testing: $basename"
            dotnet run --project Shell/doc2text/doc2text.csproj -- "samples/$file" temp-textbox-test.txt
            
            # Check content extraction
            local word_count=$(wc -w < temp-textbox-test.txt)
            echo "  Extracted words: $word_count"
            
            if [ -f "samples/$basename.expected.txt" ]; then
                local expected_words=$(wc -w < "samples/$basename.expected.txt")
                local percentage=$((word_count * 100 / expected_words))
                
                echo "  Expected words: $expected_words"
                echo "  Content percentage: $percentage%"
                
                if [ $percentage -ge 80 ]; then
                    echo "  PASS: Good content extraction"
                    ((passed_tests++))
                else
                    echo "  FAIL: Poor content extraction (possible missing TextBox content)"
                fi
            else
                # Without expected output, check for reasonable content amount
                if [ $word_count -gt 50 ]; then
                    echo "  PASS: Substantial content extracted"
                    ((passed_tests++))
                else
                    echo "  WARN: Limited content extracted"
                fi
            fi
            
            rm -f temp-textbox-test.txt
        fi
    done
    
    echo "TextBox extraction test results: $passed_tests/$total_tests passed"
}

validate_textbox_extraction
```

### Content Completeness Testing
```bash
# Test TextBox content completeness
test_textbox_completeness() {
    echo "Testing TextBox content completeness..."
    
    # Compare with reference implementations
    if [ -f "samples/news-example.doc" ]; then
        echo "=== news-example.doc content comparison ==="
        
        # b2xtranslator output
        dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc b2x-output.txt
        local b2x_words=$(wc -w < b2x-output.txt)
        
        echo "b2xtranslator words: $b2x_words"
        
        # Reference tool comparison
        if command -v antiword >/dev/null; then
            antiword samples/news-example.doc > antiword-output.txt 2>/dev/null
            local antiword_words=$(wc -w < antiword-output.txt)
            echo "antiword words: $antiword_words"
            
            local percentage=$((b2x_words * 100 / antiword_words))
            echo "b2xtranslator vs antiword: $percentage%"
            
            if [ $percentage -lt 70 ]; then
                echo "WARNING: Significant content gap - possible missing TextBox content"
            fi
            
            rm -f antiword-output.txt
        fi
        
        if command -v catdoc >/dev/null; then
            catdoc samples/news-example.doc > catdoc-output.txt 2>/dev/null
            local catdoc_words=$(wc -w < catdoc-output.txt)
            echo "catdoc words: $catdoc_words"
            
            local percentage=$((b2x_words * 100 / catdoc_words))
            echo "b2xtranslator vs catdoc: $percentage%"
            
            rm -f catdoc-output.txt
        fi
        
        rm -f b2x-output.txt
    fi
}

test_textbox_completeness
```

## Proposed Solutions

### 1. TextBox Object Recognition and Parsing
- **Drawing Object Processing**: Enhance drawing object parsing to recognize TextBox elements
- **TextBox Structure Parsing**: Parse TextBox binary structures to extract content
- **Content Stream Identification**: Identify and process TextBox text streams
- **Object Hierarchy Understanding**: Handle TextBox relationships within document structure

### 2. Text Extraction Enhancement
- **TextBox Text Extraction**: Extract text content from TextBox objects
- **Content Integration**: Integrate TextBox content into main text flow appropriately
- **Positioning Context**: Consider TextBox positioning for content placement
- **Formatting Preservation**: Preserve TextBox text formatting where appropriate

### 3. Configuration Options
- **TextBox Inclusion Control**: Implement `--no-textboxes` option for excluding TextBox content
- **Extraction Mode**: Allow different modes for TextBox content handling
- **Content Placement**: Options for where TextBox content appears in output
- **Formatting Options**: Control how TextBox content is formatted in text output

### 4. Enhanced Document Processing
- **Complete Object Model**: Ensure document object model includes TextBox objects
- **Content Completeness**: Verify all document content including TextBoxes is processed
- **Error Handling**: Graceful handling of complex or corrupted TextBox objects
- **Performance Optimization**: Efficient processing of documents with many TextBoxes

## Files to Investigate
- `Text/TextMapping/` - Text conversion mappings for TextBox content
- `Doc/` - Word document parsing for TextBox/drawing objects
- `Text/TextMapping/DrawingMapping.cs` or similar - Drawing object processing
- `Text/TextMapping/TextBoxMapping.cs` or similar - TextBox-specific mapping
- Drawing object and shape processing classes

## Expected Behavior
1. **Content Extraction**: All TextBox text content should be extracted and included
2. **Content Integration**: TextBox content should be appropriately integrated into text output
3. **Configuration Support**: Support for `--no-textboxes` option to exclude TextBox content
4. **Complete Processing**: No loss of content due to unprocessed TextBox elements
5. **Error Handling**: Graceful handling of complex TextBox structures

## Success Criteria
1. **Complete Content**: All TextBox content extracted and included in text output
2. **Content Preservation**: >95% content preservation for documents with TextBoxes
3. **Configuration Options**: Working `--no-textboxes` command-line option
4. **Performance**: No significant performance impact from TextBox processing
5. **Reliability**: Consistent TextBox handling across different document types

## Testing Commands
```bash
# Build and test TextBox content handling
dotnet build b2xtranslator.sln
dotnet test IntegrationTests/IntegrationTests.csproj --filter "Category=TextBoxContent"

# Test news-example.doc specifically
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc output.txt

# Check content completeness
if [ -f samples/news-example.expected.txt ]; then
    echo "Expected content:"
    wc -w samples/news-example.expected.txt
    echo "Actual content:"
    wc -w output.txt
    
    echo "Content comparison:"
    diff samples/news-example.expected.txt output.txt | head -20
fi

# Test TextBox exclusion option (when implemented)
# dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc output-no-textbox.txt --no-textboxes

# Content analysis
echo "Content analysis:"
echo "Total words: $(wc -w < output.txt)"
echo "Total lines: $(wc -l < output.txt)"
echo "Sample content:"
head -10 output.txt

# Binary analysis for TextBox structures
strings samples/news-example.doc | grep -i "text" | head -10
hexdump -C samples/news-example.doc | grep -i "text\|box\|shape" | head -5
```