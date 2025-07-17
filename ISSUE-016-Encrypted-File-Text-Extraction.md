# ISSUE-016: Encrypted File Text Extraction Should Throw Specific Exception

## Problem Description

Currently, when attempting to perform text extraction on encrypted Microsoft Office binary files (e.g., `.doc`), the `b2xtranslator` library does not provide a clear indication that the file is encrypted and cannot be processed. This can lead to ambiguous errors, unexpected behavior, or incomplete output without a proper explanation for the user.

## Expected Behavior

When `b2xtranslator` encounters an encrypted binary Office file during text extraction, it should:

1.  **Detect Encryption:** Reliably identify that the input file is encrypted.
2.  **Throw a Specific Exception:** Instead of failing silently or with a generic error, a dedicated exception (e.g., `EncryptedFileException` or `UnsupportedFileFormatException` with a specific message) should be thrown.
3.  **Provide Clear Message:** The exception message should clearly state that the file is encrypted and therefore cannot be processed for text extraction.

## Rationale

Providing a specific exception for encrypted files improves the robustness and user-friendliness of the library. It allows calling applications to:

-   Gracefully handle encrypted files.
-   Inform users precisely why a file cannot be processed.
-   Differentiate between actual parsing errors and expected limitations due to encryption.

## Proposed Solution

-   Implement logic within the file format parsers (`Doc/`) to detect encryption headers or flags.
-   Define a new custom exception type (e.g., `B2xTranslator.Common.Exceptions.EncryptedFileException`).
-   Throw this exception with an informative message when an encrypted file is detected during text extraction or conversion attempts.
