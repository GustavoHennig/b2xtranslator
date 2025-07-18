# Issue #010: TextBox Content Handling and Extraction

## Problem Description
Text contained within TextBoxes in Word documents is not being recognized or extracted during the conversion process. This leads to incomplete output where significant content may be missing. The `news-example.doc` sample highlights this failure: TextBox text (“Existe uma Caixa de texto aqui”) never appears in the converted output.

## Severity
**MEDIUM** – Causes significant content loss but does not crash the application.

## Impact
- Loss of content contained within TextBox elements  
- Incomplete document conversion and information preservation  
- Missing important information intentionally placed in text boxes  
- Poor fidelity for documents with complex layouts  
- Accessibility issues for users needing all content  
- Inconsistent results compared to reference tools (e.g., antiword, catdoc)

## Root Cause Analysis
1. Missing TextBox Parsing: TextBoxes stored as drawing/shape objects are not recognized.  
2. Drawing Object Limitations: Shapes holding text are treated as non‐text containers.  
3. Content Extraction Gaps: Extraction logic traverses only the main document stream, not the drawing layer.  
4. Mapping Incomplete: `TextMapping` does not include TextBox or shape traversal.  
5. Binary Structure Parsing: OfficeDrawing and OfficeArtContent binary structures are not interpreted for text runs.

## Known Affected Components and Files
- **Text Mapping**: `Text/TextMapping/TextMapping.cs`  
- **Document Parser**: `Doc/DocFileFormat/WordDocument.cs`, `Doc/DocFileFormat/OfficeArtContent.cs`  
- **Drawing Objects**: `Common/OfficeDrawing/`  
- **Shell/Text Extraction Tool**: `Shell/doc2text/doc2text.csproj`  
- **Samples**: `samples/news-example.doc`, other TextBox-containing documents

## Related Context
The README outlines a roadmap to support extraction settings such as:
- `--no-textboxes` – option to exclude TextBox elements from output

## Reproduction Steps


- Run conversion on `news-example.doc`:  
   ```bash
   dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc output.txt
   ```

- Inspect `output.txt`, should appear  `Existe uma Caixa de texto aqui` in the end.
``

## Proposed Solutions

1. **Parse Drawing Objects**  
   - Enhance `Common/OfficeDrawing` to detect TextBox shapes.  
   - Extend `OfficeArtContent` parsing to extract text runs from TextBoxes.

2. **Extend WordDocument Parser**  
   - Add `List<TextBoxContent> TextBoxes` in `WordDocument`.  
   - Implement `ParseDrawingObjectText()` to populate TextBox text.

3. **Integrate in TextMapping**  
   - In `Text/TextMapping/DocumentMapping.cs`, after main flow call `ProcessTextBoxContent()`.  
   - Append each TextBox’s `Content` (optionally based on settings).

4. **Configuration Support**  
   - Create `TextExtractionSettings.IncludeTextBoxes` (default true).  
   - Support `--no-textboxes` CLI flag to disable.

5. **Optional Advanced**  
   - Position-based placement using TextBox bounds.  
   - Metadata preservation (bounds, Z-order, formatting).

## Files to Investigate
- `Text/TextMapping/TextMapping.cs`  
- `Doc/DocFileFormat/WordDocument.cs`  
- `Common/OfficeDrawing/` (shape and TextBox support)  
- `Docx/WordprocessingMLMapping/DrawingMapping.cs` (reference)

## Success Criteria
1. Extracted text includes all TextBox content (e.g., “Existe uma Caixa de texto aqui”)  
2. No duplicated text runs  
3. CLI flag `--no-textboxes` correctly excludes TextBoxes  
4. Existing tests continue to pass  
5. Performance impact <10% on non-TextBox documents

## Testing Commands
```bash
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/news-example.doc output.txt
grep "Existe uma Caixa de texto aqui" output.txt
