# Issue #004: Parsing Stability and Crash Issues

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator crashes with exceptions and errors during document parsing when encountering malformed binary data, corrupted files, or unsupported document structures. These crashes result in complete application failure and loss of processing capability.

## Severity
**HIGH** - Application crashes prevent any document processing

## Impact
- Complete application failure with unhandled exceptions
- Loss of document processing capability
- Poor reliability and user confidence
- System instability in automated processing environments
- Data loss in batch processing scenarios
- Service interruption in server environments

## Detailed Resolution Plan

### Phase 1: Centralized Exception Handling and Logging

1.  **Create a Custom Exception Hierarchy:**
    Introduce specific exception types to represent common parsing errors. This provides more context than generic exceptions like `IndexOutOfRangeException`.

    ```csharp
    // In Common.Abstractions/Exceptions/
    public class ParsingException : Exception { ... }
    public class MalformedFileException : ParsingException { ... }
    public class UnsupportedFileVersionException : ParsingException { ... }
    ```

2.  **Implement a Centralized Error Handling Strategy:**
    In the main `Program.cs` of the shell applications, wrap the entire conversion process in a `try-catch` block that handles these specific exceptions gracefully.

    ```csharp
    // In Shell/doc2text/Program.cs
    public static void Main(string[] args)
    {
        try
        {
            // ... conversion logic ...
        }
        catch (MalformedFileException ex)
        {
            Console.Error.WriteLine($"Error: The input file appears to be corrupted. Details: {ex.Message}");
            Environment.Exit(1);
        }
        catch (UnsupportedFileVersionException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            Environment.Exit(1);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex}");
            Environment.Exit(1);
        }
    }
    ```

### Phase 2: Defensive Parsing Techniques

1.  **Add Pre-Parsing Validation:**
    Before attempting to parse the file, perform some basic validation checks in the `StructuredStorageReader`.

    ```csharp
    // In Common.CompoundFileBinary/StructuredStorage/StructuredStorageReader.cs
    public StructuredStorageReader(string filePath)
    {
        // 1. Check if file exists and is readable
        if (!File.Exists(filePath) || new FileInfo(filePath).Length < 512)
        {
            throw new MalformedFileException("File is missing or too small.");
        }

        // 2. Check for the compound file signature
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            var signature = new byte[8];
            stream.Read(signature, 0, 8);
            if (!signature.SequenceEqual(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }))
            {
                throw new MalformedFileException("Invalid OLECF file signature.");
            }
        }
        // ...
    }
    ```

2.  **Implement Safe Reading with Bounds Checking:**
    Create a helper class or extension methods for reading from byte arrays and streams that includes robust bounds checking.

    ```csharp
    public static class SafeBinaryReader
    {
        public static int ReadInt32(byte[] buffer, int offset)
        {
            if (offset + 4 > buffer.Length)
            {
                throw new MalformedFileException("Attempted to read past the end of a buffer.");
            }
            return BitConverter.ToInt32(buffer, offset);
        }
    }
    ```
    This `SafeBinaryReader` should then be used throughout the parsing code instead of direct calls to `BitConverter` or `Stream.Read` where the size is assumed.

### Phase 3: Tolerant Parsing (Graceful Degradation)

1.  **Wrap Individual Record Parsing:**
    In loops that iterate over binary records, wrap the parsing of each individual record in a `try-catch` block. This allows the process to skip a single corrupted record without failing the entire conversion.

    ```csharp
    // In a record parsing loop
    for (int i = 0; i < recordCount; i++)
    {
        try
        {
            var record = new MyRecord(stream);
            records.Add(record);
        }
        catch (ParsingException ex)
        {
            // Log the error and continue to the next record
            _logger.LogWarning($"Skipping corrupted record at index {i}. Details: {ex.Message}");
        }
    }
    ```

2.  **Use `TryParse` Pattern:**
    For complex structures, implement a `TryParse` pattern that returns a boolean indicating success instead of throwing an exception.

    ```csharp
    // In a complex object parser
    public static bool TryParse(Stream stream, out MyComplexObject result)
    {
        try
        {
            result = new MyComplexObject(stream);
            return true;
        }
        catch (ParsingException)
        {
            result = null;
            return false;
        }
    }
    ```

## Files to Investigate
- `Common.CompoundFileBinary/StructuredStorage/StructuredStorageReader.cs` - For pre-parsing validation.
- All binary record parsing classes in `Doc/`, `Ppt/`, and `Xls/` - For safe reading and tolerant parsing.
- The `Main` methods of all `Shell/` projects - For centralized exception handling.

## Success Criteria
1. **Crash-Free Operation**: The application no longer crashes with unhandled exceptions on the provided sample files.
2. **Clear Error Reporting**: When a file is corrupt, the user is given a clear, understandable error message.
3. **Graceful Degradation**: The parsers can successfully skip over corrupted sections of a file to produce a partial, but still useful, result.
4. **No Regressions**: The changes do not negatively impact the conversion of valid files.
