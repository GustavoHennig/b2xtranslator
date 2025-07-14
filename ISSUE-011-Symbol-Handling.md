# Issue #011: Incorrect Symbol Handling (Partially fixed)

## Problem Description
Special symbols, such as those from the "Symbol" or "Wingdings" fonts, are not being handled correctly during text extraction. This can result in missing characters, incorrect characters (e.g., showing a letter instead of the symbol), or garbled text in the output. The `README.md` mentions `Bug49908.doc` as a sample case.

## Severity
**LOW** - Affects the fidelity of the output, but is less likely to cause major content loss than other issues.

## Impact
- Loss of fidelity and precision in the extracted text.
- Can change the meaning of technical or scientific documents that rely on special symbols.
- Poor rendering of lists that use symbolic bullets.

## Root Cause Analysis
Symbols in Word documents are often represented by a character code within a run of text that is formatted with a specific symbol font (e.g., Symbol, Wingdings, Webdings). The text extraction logic may be reading the character code but failing to interpret it correctly because it doesn't account for the font.

1.  **Font-Specific Encoding**: The extractor treats the character as if it were in a standard text font, ignoring the font-specific encoding.
2.  **Missing Mapping**: There is no mapping from the character code in a given symbol font to its correct Unicode equivalent.

## Known Affected Components
- **Text Mapping**: `Text/TextMapping/`
- **Character Run Parsing**: `Doc/DocFileFormat/Data/CharX.cs` (or similar)

## Reproduction Steps

### Basic Reproduction
1.  Use the sample file `Bug49908.doc`.
2.  Run the text extraction tool:
    ```bash
    dotnet run --project Shell/doc2text/doc2text.csproj -- samples/Bug49908.doc output.txt
    ```
3.  Examine `output.txt` and compare it to the original document to see where symbols are rendered incorrectly.

### Scenario 2 `symbol.doc`
See the image `ISSUE-011-Symbol-Handling.png`
The output should show the exact corresponding UTF-8 code of the symbol presented.

## Testing Strategy

### Integration Tests
- Use `Bug49908.doc` as the primary test case.
- If possible, create a new document with a wider range of symbols from different fonts.
- Create `.expected.txt` files with the correct Unicode characters for the symbols.
- Add these to the `SampleDocFileTextExtractionTests.cs` test suite.

## Proposed Solutions

### 1. Detect Symbol Fonts
- When parsing character runs, detect when a symbol font is being used. The font information is associated with the formatting of the run.

### 2. Implement Character Mapping
- Create a mapping utility that can convert a character code from a known symbol font to a Unicode character.
- For example, if the font is "Symbol" and the character is 'a' (0x61), it should be mapped to the Greek alpha symbol (U+03B1).
- This may require creating explicit mapping tables for common symbol fonts like Symbol and Wingdings.
- The `TextMapping` should use this utility to translate characters from symbol fonts before writing them to the output.

## Files to Investigate
- `Text/TextMapping/TextMapping.cs`
- `Doc/DocFileFormat/Data/CharX.cs`
- `Doc/DocFileFormat/Structures/CHP.cs` (Character Properties, where font info is stored)
- `Common/Tools/Utils.cs` (A potential place for the mapping utility)

## Success Criteria
1.  **Correct Symbols**: Common symbols from the Symbol font are correctly translated to their Unicode equivalents.
2.  **Correct Bullets**: List items that use symbols as bullets are rendered correctly.
3.  **No Regressions**: Standard text characters are unaffected.

## Current Status (Verified)
**PROBLEM PERSISTS** - Testing shows symbol handling issues:
- Expected output: "This is a symbol: ?" (shows placeholder character)
- Actual output: "This is a symbol: " (missing even the placeholder)
- The symbol is completely missing from the extracted text
- Issue remains active and requires resolution

## Detailed Resolution Plan

### Phase 1: Symbol Detection Analysis (1-2 days)

#### 1.1 Binary Symbol Investigation
```bash
# Analyze Bug49908.doc binary structure for symbol data
hexdump -C samples/Bug49908.doc | grep -E "[\x80-\xFF]" | head -20
strings samples/Bug49908.doc | grep -E "[^\x00-\x7F]"

# Check for font references
strings samples/Bug49908.doc | grep -i "symbol\|wingding\|math"
```

**Investigation Tasks:**
- Identify how symbols are stored in Word binary format
- Determine if symbols use special fonts (Symbol, Wingdings, etc.)
- Check if symbols are stored as Unicode or font-specific codes
- Analyze character encoding patterns in binary data

#### 1.2 Current Parsing Analysis
**Files to examine:**
- `Doc/DocFileFormat/CharacterProperties.cs` - Character formatting
- `Text/TextMapping/CharacterPropertiesMapping.cs` - Character conversion
- `Common/Tools/` - Character encoding utilities

**Key Questions:**
- How are font-specific characters currently handled?
- Is there Unicode conversion logic for symbols?
- Are symbols being recognized but not converted properly?

### Phase 2: Character Encoding Enhancement (2-3 days)

#### 2.1 Create Symbol Mapping Tables
**File: `Common/Tools/SymbolMapping.cs`** (create)

```csharp
public static class SymbolMapping
{
    // Symbol font character mappings
    private static readonly Dictionary<byte, string> SymbolFontMap = new()
    {
        { 0x61, "α" },  // Greek alpha
        { 0x62, "β" },  // Greek beta
        { 0x67, "γ" },  // Greek gamma
        { 0x64, "δ" },  // Greek delta
        { 0xD1, "±" },  // Plus-minus
        { 0xD7, "×" },  // Multiplication
        { 0xF7, "÷" },  // Division
        // Add complete mapping tables for Symbol font
    };
    
    private static readonly Dictionary<byte, string> WingdingsFontMap = new()
    {
        { 0x4A, "☺" },  // Smiley face
        { 0x4B, "☻" },  // Black smiley
        { 0x6C, "✓" },  // Check mark
        // Add complete Wingdings mapping
    };
    
    public static string ConvertSymbolCharacter(byte charCode, string fontName)
    {
        return fontName?.ToLower() switch
        {
            "symbol" when SymbolFontMap.TryGetValue(charCode, out var symbol) => symbol,
            "wingdings" when WingdingsFontMap.TryGetValue(charCode, out var wingding) => wingding,
            _ => ConvertToUnicode(charCode, fontName)
        };
    }
    
    private static string ConvertToUnicode(byte charCode, string fontName)
    {
        // Fallback Unicode conversion logic
        if (charCode > 127)
        {
            // Attempt Unicode interpretation
            return ((char)charCode).ToString();
        }
        return "?";  // Placeholder for unknown symbols
    }
}
```

#### 2.2 Enhance Character Processing
**File: `Text/TextMapping/CharacterPropertiesMapping.cs`**

```csharp
public class CharacterPropertiesMapping
{
    public void ProcessCharacterRun(CharacterRun run)
    {
        foreach (var character in run.Characters)
        {
            var fontName = GetFontName(character.FontIndex);
            
            if (IsSymbolFont(fontName))
            {
                // Use symbol mapping
                var symbolText = SymbolMapping.ConvertSymbolCharacter(
                    character.Code, fontName);
                _writer.WriteText(symbolText);
            }
            else
            {
                // Standard character processing
                ProcessStandardCharacter(character);
            }
        }
    }
    
    private bool IsSymbolFont(string fontName)
    {
        var symbolFonts = new[] { "symbol", "wingdings", "webdings", "mt extra" };
        return symbolFonts.Contains(fontName?.ToLower());
    }
}
```

### Phase 3: Unicode Support Enhancement (1-2 days)

#### 3.1 Improve Unicode Handling
**File: `Doc/DocFileFormat/CharacterProperties.cs`**

```csharp
public class CharacterProperties
{
    // Add Unicode support properties
    public bool IsUnicodeCharacter { get; set; }
    public ushort UnicodeValue { get; set; }
    public string FontName { get; set; }
    
    public string GetDisplayCharacter()
    {
        if (IsUnicodeCharacter && UnicodeValue > 0)
        {
            return char.ConvertFromUtf32(UnicodeValue);
        }
        
        // Fallback to font-specific mapping
        return SymbolMapping.ConvertSymbolCharacter((byte)CharacterCode, FontName);
    }
}
```

#### 3.2 Font-Aware Text Extraction
**File: `Text/TextMapping/DocumentMapping.cs`**

```csharp
public class DocumentMapping : AbstractOpenXmlMapping
{
    private readonly Dictionary<int, string> _fontTable = new();
    
    public override void Apply(IVisitable x)
    {
        var doc = x as WordDocument;
        
        // Build font table first
        BuildFontTable(doc);
        
        // Process document with font awareness
        ProcessDocumentContent(doc);
    }
    
    private void BuildFontTable(WordDocument doc)
    {
        foreach (var font in doc.FontTable.Fonts)
        {
            _fontTable[font.Index] = font.Name;
        }
    }
}
```

### Phase 4: Error Handling and Fallbacks (1 day)

#### 4.1 Graceful Symbol Degradation
```csharp
public class SymbolFallbackHandler
{
    public static string HandleUnknownSymbol(byte charCode, string fontName)
    {
        // Log unknown symbol for future mapping
        TraceLogger.DebugInternal($"Unknown symbol: Code={charCode:X2}, Font={fontName}");
        
        // Provide contextual fallback
        return fontName?.ToLower() switch
        {
            "symbol" => GetMathematicalFallback(charCode),
            "wingdings" => GetDecoративeFallback(charCode),
            _ => "?"  // Standard placeholder
        };
    }
    
    private static string GetMathematicalFallback(byte code)
    {
        // Common mathematical symbol approximations
        return code switch
        {
            >= 0x61 and <= 0x7A => "[greek]",  // Greek letters
            >= 0xB1 and <= 0xF7 => "[math]",   // Math operators
            _ => "?"
        };
    }
}
```

### Phase 5: Testing and Validation (1 day)

#### 5.1 Symbol-Specific Unit Tests
**File: `UnitTests/SymbolHandlingTests.cs`** (create)

```csharp
[TestFixture]
public class SymbolHandlingTests
{
    [Test]
    public void ShouldConvertSymbolFontCharacters()
    {
        var mapping = new SymbolMapping();
        
        // Test known Symbol font mappings
        Assert.That(mapping.ConvertSymbolCharacter(0x61, "Symbol"), Is.EqualTo("α"));
        Assert.That(mapping.ConvertSymbolCharacter(0x62, "Symbol"), Is.EqualTo("β"));
    }
    
    [Test]
    public void ShouldHandleWingdingsFont()
    {
        var mapping = new SymbolMapping();
        var result = mapping.ConvertSymbolCharacter(0x4A, "Wingdings");
        Assert.That(result, Is.Not.EqualTo("?"));  // Should not be placeholder
    }
    
    [Test]
    public void ShouldExtractBug49908Symbol()
    {
        var result = ConvertToText("samples/Bug49908.doc");
        Assert.That(result, Does.Not.Match(@"This is a symbol:\s*$"));  // Should have symbol
        Assert.That(result.Length, Is.GreaterThan(20));  // Should have content
    }
}
```

#### 5.2 Integration Testing
```bash
# Test symbol extraction across multiple files
for file in samples/Bug49908.doc samples/equation.doc samples/symbol.doc; do
    if [ -f "$file" ]; then
        echo "Testing symbols in: $(basename "$file")"
        dotnet run --project Shell/doc2text/doc2text.csproj -- "$file" temp-symbol.txt
        
        # Check for symbols vs placeholders
        symbol_count=$(grep -o "[^\x00-\x7F]" temp-symbol.txt | wc -l)
        placeholder_count=$(grep -o "?" temp-symbol.txt | wc -l)
        
        echo "  Unicode symbols: $symbol_count"
        echo "  Placeholders: $placeholder_count"
        
        rm -f temp-symbol.txt
    fi
done
```

### Phase 6: Advanced Symbol Features (Optional - 1-2 days)

#### 6.1 Symbol Context Awareness
```csharp
public class ContextAwareSymbolConverter
{
    public string ConvertWithContext(byte charCode, string fontName, string documentLanguage)
    {
        // Consider document language for symbol interpretation
        // Handle different symbol conventions (US vs European)
        // Provide culturally appropriate symbol representations
    }
}
```

#### 6.2 Symbol Metadata Preservation
```csharp
public class SymbolInfo
{
    public string OriginalCharCode { get; set; }
    public string FontName { get; set; }
    public string UnicodeEquivalent { get; set; }
    public string FallbackRepresentation { get; set; }
    public SymbolCategory Category { get; set; }  // Math, Decorative, etc.
}
```

## Implementation Timeline
- **Day 1-2**: Symbol detection and binary analysis
- **Day 3-4**: Character mapping and font handling
- **Day 5-6**: Unicode support and text extraction
- **Day 7**: Testing, validation, and optimization

## Success Metrics
1. `Bug49908.doc` shows symbol content (not empty space)
2. Symbol fonts (Symbol, Wingdings) properly converted
3. Mathematical symbols display as Unicode equivalents
4. Fallback placeholders only when absolutely necessary
5. Performance impact < 5% for symbol-heavy documents

## Testing Commands
```bash
# Build and test symbol handling
dotnet build b2xtranslator.sln
dotnet test UnitTests/UnitTests.csproj --filter "Category=SymbolHandling"

# Test Bug49908.doc specifically - should show symbol
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/Bug49908.doc output.txt
cat output.txt  # Should show: "This is a symbol: [symbol-char]"

# Symbol analysis across test files
for file in samples/*symbol*.doc samples/Bug49908.doc samples/equation.doc; do
    if [ -f "$file" ]; then
        echo "=== Testing $(basename "$file") ==="
        dotnet run --project Shell/doc2text/doc2text.csproj -- "$file" temp.txt
        echo "Content: $(cat temp.txt)"
        echo "Unicode chars: $(grep -o '[^\x00-\x7F]' temp.txt | wc -l)"
        echo "Placeholders: $(grep -o '?' temp.txt | wc -l)"
        rm -f temp.txt
        echo ""
    fi
done

# Regression testing
dotnet test IntegrationTests/IntegrationTests.csproj
```
