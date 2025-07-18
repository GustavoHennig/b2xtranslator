# Issue #015: Text Extraction Mismatch for `other.doc`

## Current Status
**New issue, requires investigation.**

## Problem Description
The text extraction for the sample file `samples/other.doc` does not match the expected output in `samples/other.expected.txt`. This indicates a potential flaw in the text extraction logic for this specific document or a class of similar documents.
(complex_sections.doc has a similar issue)
## Severity
**MEDIUM** - Affects the quality and reliability of text extraction.

## Expected output

Line breaks or spaces are not important.

```txt
Big line
Bold simple line
Italic line
```

## Impact
- Incorrect text output for certain documents.
- Reduced trust in the accuracy of the conversion process.
- Potential for downstream processing errors if the extracted text is used in automated workflows.

## How to Reproduce

1. **Run the `doc2text` converter on the sample file:**
   ```bash
   dotnet run --project Shell/doc2text/doc2text.csproj -- samples/other.doc output.txt
   ```

2. **Compare the output with the expected text:**
   After running the command, compare the contents of the generated `output.txt` with `samples/other.expected.txt`.

   ```bash
   # On Windows, you can use fc (File Compare)
   fc output.txt samples/other.expected.txt

   # On Linux or macOS, you can use diff
   diff output.txt samples/other.expected.txt
   ```

   A discrepancy between the two files confirms the issue.

## Important

- It is possible to find a partial solution in `FASTSAVE-working-example.md`.
- There is a complete functional C code showing how to hanlde fastsaved documents pasted here `REFERENCE_FOR_FULL_WORD_PARSING.md` (the file is HUGE).
- Perform regression tests using the file `47950_normal.doc` to verify whether other files have been affected.

## Initial Analysis (Hypothesis)
The cause could be related to several factors, including but not limited to:
- Incorrect handling of specific formatting instructions in the `.doc` file.
- Issues with character encoding or symbol interpretation.
- Problems with extracting text from tables, headers, footers, or text boxes.
- A bug in the paragraph or character run processing logic.

Further investigation is needed to pinpoint the exact cause.
