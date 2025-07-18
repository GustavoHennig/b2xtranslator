# Issue #007: Missing Paragraph Content Issues

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator fails to extract entire paragraphs from certain Word documents, resulting in incomplete text output where significant portions of document content are missing. This is a critical content preservation issue that affects document completeness.

## Affected files:

- fastsavedmix.doc
- 53379.doc
- "Bug50936_1.doc"
- "Bug47958.doc"
- title.doc
- ob_is.doc
- o_kurs.doc (repeating page)
- ESTAT Article comparing RU-LFS-22 12 05_EN.doc
- complex_sections.doc
- Bug53380_3.doc
- Bug52032_3.doc
- Bug51944.doc
- Bug47958.doc
- Bug33519.doc


## Scenarios

This paragraph is missing from the file: `\samples\EL_TechnicalTemplateHandling.doc`

```
Updating the value of any property within this dialog will be reflected in the document after closing the dialogue. This is important since all the properties displayed in the document are updated, independently of their number of occurrences in the document. The Properties Dialog can thus be called any time you need to update the document properties. For all properties, the last setting will be memorized by the dialog. 
NOTE : the options present in this dialog rely on specific bookmarks and fields inside the document in order to be able to add the requested information at the defined spot. If these bookmarks and/or fields are removed or modified, it is possible that the properties dialog will not be able to achieve the desired result. 
Properties Dialog: Front Information Table
```

## Detailed Resolution Plan

The root cause for missing paragraphs can be complex, often stemming from how Word structures its files, especially with features like "Fast Save" or specific paragraph formatting properties.

### Important:
Applying the fixes described in `FASTSAVE-working-example.md` makes `fastsavedmix.doc` work, but it affects all other documents. However, the information needed to make it work is there.

### Phase 1: Enhanced Diagnostics and Logging

Before fixing the issue, we need to pinpoint exactly *why* content is being dropped. This requires adding more detailed logging to the parsing process.

1.  **Log Paragraph Processing Details:**
    In `Text/TextMapping/TextMapping.cs` or wherever paragraphs are iterated, add detailed logging.

    ```csharp
    // In the paragraph processing loop
    foreach (var paragraph in allParagraphs)
    {
        // Log key properties of every paragraph being processed
        _logger.LogDebug($"Processing paragraph at CP range {paragraph.CharacterPositionStart}-{paragraph.CharacterPositionEnd}.");
        _logger.LogTrace($"Paragraph Properties (PAPX): IsHidden={paragraph.IsHidden}, IsInTable={paragraph.IsInTable}");

        if (ShouldSkip(paragraph))
        {
            _logger.LogWarning($"SKIPPING paragraph at CP range {paragraph.CharacterPositionStart}-{paragraph.CharacterPositionEnd} due to skip condition.");
            continue;
        }

        // ... rest of the processing
    }
    ```

### Phase 2: Correctly Handle "Fast-Saved" Documents

The presence of `fastsavedmix.doc` in the list of problematic files strongly suggests that the "Fast Save" feature is a primary cause of missing content. When a document is fast-saved, edits are appended to the end of the file instead of being integrated into the main document structure. A parser must correctly reconstruct the document from these fragments.

1.  **Detect Fast-Saved Mode:**
    In `Doc/DocFileFormat/FileInformationBlock.cs`, ensure the `fFastSaved` bit is correctly read from the FIB.

2.  **Implement Fragment Reconstruction Logic:**
    In `Doc/DocFileFormat/WordDocument.cs`, if `fib.fFastSaved` is true, the parser cannot simply read the text stream linearly. It must follow the chain of edits, which is significantly more complex. The general approach is:
    - The `FIB` contains pointers to the original document's text and the location of the edits.
    - The parser must read the original text and then apply the appended edits (insertions, deletions) in the correct order to reconstruct the final state of the document.
    - This is a major undertaking and requires careful study of the `[MS-DOC]` specification regarding fast-saves.

### Phase 3: Analyze and Handle Paragraph Properties

Paragraphs might be skipped based on their formatting properties (`PAPX`).

1.  **Expose All Relevant `PAPX` Flags:**
    In `Doc/DocFileFormat/Structures/ParagraphProperties.cs`, ensure all flags that could affect visibility are parsed. This includes:
    - `fVanish`: Indicates hidden text.
    - `fTtp`: Indicates a table paragraph.
    - `fAbs`: Indicates an absolutely positioned paragraph (which might be outside the normal flow).

2.  **Make an Explicit Decision on What to Skip:**
    In the `TextMapping`, the decision to skip a paragraph should be deliberate. For text extraction, the default should be to extract everything, including content that might be visually hidden in Word.

    ```csharp
    // In TextMapping.cs
    private bool ShouldExtractParagraph(Paragraph p)
    {
        // For text extraction, we generally want all content.
        // We should NOT skip hidden text unless explicitly configured to.
        if (p.IsHidden) {
             _logger.LogInformation("Including a hidden paragraph.");
        }

        // We might want to skip table structural paragraphs if they don't contain user text.
        if (p.IsTableTerminationCharacter) {
            return false;
        }

        return true;
    }
    ```

### Phase 4: Verify Completeness of Text Traversal

It's possible that the logic that walks the document's text stream is simply not visiting all the character ranges (`CP`).

1.  **Implement a `CP` Coverage Map:**
    Create a simple boolean array or `BitArray` representing all the character positions in the document. As the `TextMapping` processes each paragraph, it marks the corresponding `CP` range in the map as "visited".

2.  **Report on Gaps:**
    After the conversion is complete, check the coverage map for any unvisited ranges. Log these gaps as warnings. This will immediately highlight if and where the parser is failing to walk the entire document.

    ```csharp
    // After conversion
    var unvisitedRanges = FindUnvisitedRanges(coverageMap);
    if (unvisitedRanges.Any())
    {
        _logger.LogError("The following character ranges were not processed:");
        foreach(var range in unvisitedRanges)
        {
            _logger.LogError($"  - CP {range.Start} to {range.End}");
        }
    }
    ```

## Files to Investigate
- `Doc/DocFileFormat/FileInformationBlock.cs`: To check for the `fFastSaved` flag.
- `Doc/DocFileFormat/WordDocument.cs`: The main class for parsing the document, where fast-save logic would be implemented.
- `Doc/DocFileFormat/Structures/ParagraphProperties.cs`: To ensure all visibility-related flags are parsed.
- `Text/TextMapping/TextMapping.cs`: To implement the logging, property checks, and CP coverage map.

## Success Criteria
1.  **Content Completeness**: The word and paragraph counts of the output for the known problematic files should match their `.expected.txt` files with >99% accuracy.
2.  **Fast-Save Support**: The `fastsavedmix.doc` file is converted correctly, with no missing lines.
3.  **No Regressions**: The changes do not cause content to be lost from files that were previously working correctly.
4.  **Clear Diagnostics**: If content is intentionally skipped (e.g., a table termination character), it is logged with a clear reason.
