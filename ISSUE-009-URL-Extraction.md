# Issue #009: URL Extraction Issues (Duplicate of #006)

## Current Status
**This issue is a duplicate of [Issue #006: Internal Field Markers in Output Text](./ISSUE-006-Internal-Field-Markers.md).**

The root cause and the resolution for this problem are identical to those described in `ISSUE-006`. The core of the problem is the incorrect handling of Word's `HYPERLINK` field codes.

## Problem Description
The b2xtranslator fails to properly extract and display URLs from Word documents during text conversion, resulting in missing hyperlink information in the converted text output. This affects the accessibility and completeness of hyperlinked content.

## Resolution Plan
The detailed resolution plan in `ISSUE-006` covers the necessary steps:

1.  **Implement a state machine** to distinguish between field codes and field results.
2.  **Create a `FieldEvaluator`** that can parse the `HYPERLINK` field code to extract the URL.
3.  **Format the output** to include both the hyperlink's display text and its URL in a readable format, such as `DisplayText [URL]`.

By implementing the solution for `ISSUE-006`, this issue will be resolved simultaneously.
