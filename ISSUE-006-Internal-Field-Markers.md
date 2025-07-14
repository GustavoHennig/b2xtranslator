# Issue #006: Internal Field Markers in Output Text

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator includes internal Word field markers and formatting codes in the converted text output, such as HYPERLINK, MERGEFORMAT, DOCPROPERTY, PAGEREF, and SHAPE markers. These internal markers should be filtered out or converted to appropriate text representations.

## Severity
**MEDIUM** - Affects output quality and readability but doesn't prevent conversion

## Impact
- Poor text output quality with technical markers visible
- Reduced readability of converted documents
- Confusion for end users seeing internal Word codes
- Inconsistent output format across different documents
- Professional appearance degradation in converted text
- Potential parsing issues in downstream text processing

## Detailed Resolution Plan

Word documents represent fields using special control characters. A field consists of a start character (`0x13`), the field code (e.g., `HYPERLINK "http://example.com"`), a separator character (`0x15`), the most recently calculated result of the field (e.g., "Click here"), and an end character (`0x14`). The current text extractor is simply printing all of these parts as plain text.

The plan is to properly parse these fields and intelligently decide what to output.

### Phase 1: Robust Field Character Parsing

1.  **Identify Field Control Characters:**
    The core of the issue lies in how character runs are processed. The logic that iterates through characters needs to be aware of the special field control characters.

    -   `0x13` (Field Begin)
    -   `0x14` (Field End)
    -   `0x15` (Field Separator)

2.  **Create a State Machine for Text Extraction:**
    In `Text/TextMapping/TextMapping.cs`, the character processing logic should be managed by a simple state machine.

    ```csharp
    private enum ParsingState { Normal, InFieldCode, InFieldResult }
    private ParsingState _state = ParsingState.Normal;

    // In the character processing loop
    foreach (char c in run.Text)
    {
        if (c == 0x13) { _state = ParsingState.InFieldCode; continue; }
        if (c == 0x14) { _state = ParsingState.Normal; continue; }
        if (c == 0x15) { _state = ParsingState.InFieldResult; continue; }

        if (_state == ParsingState.Normal || _state == ParsingState.InFieldResult)
        {
            // Only write characters if they are in the normal stream or in the field result.
            // This automatically filters out the field codes.
            _writer.Write(c);
        }
        // If state is InFieldCode, do nothing, effectively skipping it.
    }
    ```

### Phase 2: Advanced Field Evaluation (for Hyperlinks)

The simple state machine from Phase 1 will hide the field codes but will not correctly handle cases like hyperlinks where the URL is part of the code. For this, a more advanced approach is needed.

1.  **Model the Field Structure:**
    When the parser encounters a `0x13` character, instead of just changing state, it should parse the entire field into an object.

    ```csharp
    // In Doc/DocFileFormat/ or similar
    public class Field
    {
        public string FieldCode { get; set; }
        public string ResultText { get; set; }
    }
    ```

2.  **Implement a `Field` Parser:**
    This parser will be triggered by `0x13` and will read characters until it has populated both the `FieldCode` and `ResultText` properties by respecting the `0x15` separator and `0x14` end marker.

3.  **Create a `FieldEvaluator`:**
    This class will be responsible for converting a `Field` object into a meaningful string for text output.

    ```csharp
    // In Text/TextMapping/
    public static class FieldEvaluator
    {
        public static string ToText(Field field)
        {
            if (field.FieldCode.StartsWith(" HYPERLINK"))
            {
                // Try to extract the URL from the field code
                var match = Regex.Match(field.FieldCode, @"\"(.*?)\"");
                if (match.Success)
                {
                    var url = match.Groups[1].Value;
                    // For text output, format it nicely
                    return $"{field.ResultText} [{url}]";
                }
            }

            // For all other field types (MERGEFORMAT, DOCPROPERTY, etc.),
            // the last calculated result is usually what the user wants to see.
            return field.ResultText;
        }
    }
    ```

4.  **Update `TextMapping` to Use the Evaluator:**
    The `TextMapping` will now collect characters into `Field` objects. When a complete field is parsed, it will pass it to the `FieldEvaluator` and write the returned string to the output.

## Files to Investigate
- `Text/TextMapping/TextMapping.cs`: This is the primary location for the changes.
- `Doc/DocFileFormat/Data/CharX.cs` (or wherever raw character runs are parsed): This may need to be modified to expose the field control characters rather than stripping them.
- `Doc/DocFileFormat/Structures/ParagraphProperties.cs`: To see how fields are associated with paragraphs.

## Success Criteria
1.  **No Field Markers**: The output text contains no instances of `HYPERLINK`, `MERGEFORMAT`, `DOCPROPERTY`, `PAGEREF`, or `SHAPE`.
2.  **Hyperlinks Preserved**: For hyperlinks, both the display text and the URL are present in a readable format (e.g., `Example Website [http://example.com]`).
3.  **Field Results Kept**: For fields like `DOCPROPERTY`, the calculated result (e.g., the author's name) is present in the output.
4.  **No Content Loss**: The process of removing field codes does not accidentally remove legitimate user-written text.
