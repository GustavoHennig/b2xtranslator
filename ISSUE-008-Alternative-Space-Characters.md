# Issue #008: Alternative Space Character Handling

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator fails to properly handle alternative space characters (non-standard space characters with different Unicode values) during text conversion, resulting in missing or incorrect spacing in the output text. This affects document readability and text formatting preservation.

## Severity
**LOW** - Affects text formatting and spacing but doesn't prevent basic conversion

## Impact
- Incorrect spacing in converted text documents
- Poor text formatting and readability
- Inconsistent space handling across different documents
- Potential text parsing issues in downstream applications
- Loss of original document spacing intentions
- Problems with documents containing non-ASCII space characters

## Analysis of Proposed Solution

### Code Review

- **Converter.cs**: No whitespace normalization or post-processing is currently performed. Text conversion delegates to mapping classes, which handle character extraction and writing.
- **DocumentMapping.cs**: The `writeText` method writes characters directly, only converting Windows-1252 control characters to Unicode. There is no normalization of alternative Unicode space characters, nor removal of zero-width spaces. Whitespace sequences are not consolidated.

### Appropriateness of Proposed Solution

The suggested approach—creating a normalization utility, integrating it into character processing, and consolidating whitespace—is appropriate and necessary. It will:
- Ensure consistent spacing by converting all alternative space characters to standard ASCII space.
- Remove zero-width spaces, preventing invisible artifacts.
- Improve readability by consolidating multiple consecutive spaces.

### Recommendation

Proceed with the proposed solution. It follows best practices for Unicode text normalization and will resolve the issue as described. No changes to the code have been made yet.

## Detailed Resolution Plan

The most robust way to solve this issue is to normalize all whitespace characters encountered during text extraction into a standard ASCII space (`U+0020`). This ensures consistent and predictable spacing in the final output.

### Phase 1: Create a Character Normalization Utility

1.  **Implement a `CharacterNormalizer`:**
    In the `Common/Tools/` project, create a new static class for character utility functions.

    ```csharp
    // In Common/Tools/CharUtils.cs
    using System.Collections.Generic;

    public static class CharUtils
    {
        private static readonly HashSet<char> WhitespaceChars = new HashSet<char>
        {
            '\u0020', // Space
            '\u00A0', // No-break space
            '\u2000', // En quad
            '\u2001', // Em quad
            '\u2002', // En space
            '\u2003', // Em space
            '\u2004', // Three-per-em space
            '\u2005', // Four-per-em space
            '\u2006', // Six-per-em space
            '\u2007', // Figure space
            '\u2008', // Punctuation space
            '\u2009', // Thin space
            '\u200A', // Hair space
            '\u3000'  // Ideographic space
        };

        private static readonly char ZeroWidthSpace = '\u200B';

        /// <summary>
        /// Normalizes various Unicode whitespace characters into a standard space.
        /// Zero-width spaces are removed entirely.
        /// </summary>
        /// <param name="c">The character to normalize.</param>
        /// <param name="normalizedChar">The normalized character.</param>
        /// <returns>True if the character is a non-zero-width space; false otherwise.</returns>
        public static bool TryNormalizeWhitespace(char c, out char normalizedChar)
        {
            if (c == ZeroWidthSpace)
            {
                normalizedChar = default;
                return false; // Remove this character
            }

            if (WhitespaceChars.Contains(c))
            {
                normalizedChar = ' '; // Normalize to standard space
                return true;
            }

            normalizedChar = c; // Not a whitespace character
            return true;
        }
    }
    ```

### Phase 2: Integrate Normalization into Text Mapping

1.  **Update the Character Processing Loop:**
    In `Text/TextMapping/TextMapping.cs`, modify the code that writes characters to the output stream to use the new normalization utility.

    ```csharp
    // In the character processing loop within TextMapping.cs
    foreach (char c in run.Text)
    {
        // ... handle field codes first (see ISSUE-006) ...

        if (CharUtils.TryNormalizeWhitespace(c, out char normalizedChar))
        {
            _writer.Write(normalizedChar);
        }
        // If TryNormalizeWhitespace returns false, it's a zero-width space and we do nothing.
    }
    ```

### Phase 3: Consolidate Whitespace (Optional but Recommended)

After normalization, the output might contain sequences of multiple spaces. A final cleanup step will improve readability.

1.  **Use a `StringBuilder` and Post-Process:**
    Instead of writing directly to the `TextWriter`, append the normalized characters to a `StringBuilder`. At the end of the conversion, get the final string and use a regular expression to consolidate whitespace before writing to the output stream.

    ```csharp
    // In TextConverter.cs
    public static void Convert(WordDocument doc, TextWriter writer)
    {
        var stringBuilder = new StringBuilder();
        var tempWriter = new StringWriter(stringBuilder);
        
        // Perform the mapping using the temp writer
        var mapping = new TextMapping(doc, tempWriter);
        mapping.Convert();

        // Get the raw text and clean it up
        string rawText = stringBuilder.ToString();
        string cleanText = Regex.Replace(rawText, @" {2,}", " "); // Consolidate spaces

        writer.Write(cleanText);
    }
    ```

## Files to Investigate
- `Text/TextMapping/TextMapping.cs`: The primary location for integrating the normalization logic.
- `Common/Tools/`: A good location for the new `CharUtils.cs` file.
- `Text/TextMapping/TextConverter.cs`: A potential place to implement the final whitespace consolidation.

## Success Criteria
1.  **Consistent Spacing**: The output document `Bug47742.doc` has consistent, single-space word separation.
2.  **No Missing Spaces**: Words are not run together due to unhandled space characters.
3.  **Readability**: The final text output is clean and readable, without excessive or irregular whitespace.
4.  **Zero-Width Spaces Removed**: Zero-width spaces are completely removed from the output and do not introduce any artifacts.
