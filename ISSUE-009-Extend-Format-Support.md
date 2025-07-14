# Issue #009: Extend File Format Support (Word 6.0/95)

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The translator currently has limited or no support for older Word file formats, such as Word 6.0 and Word 95. This prevents users from converting legacy documents.

## Severity
**MEDIUM** - Core functionality is affected for a subset of users with older files.

## Impact
- Users cannot process or migrate documents created in older versions of Microsoft Word.
- Inability to extract content from a significant body of legacy documents.

## Detailed Resolution Plan

Supporting older Word formats requires version-specific parsing logic, as the File Information Block (FIB) and other data structures are different from the Word 97-2003 format.

### Phase 1: FIB Version Detection and Dispatch

1.  **Read the FIB Version:**
    The first step is to reliably read the version number from the FIB. This is one of the very first fields in the structure.

    ```csharp
    // In Doc/DocFileFormat/FileInformationBlock.cs
    public class FileInformationBlock
    {
        public short nFib { get; private set; }

        public FileInformationBlock(byte[] fibBytes)
        {
            // The nFib field is at a specific offset
            this.nFib = BitConverter.ToInt16(fibBytes, 2);
        }
    }
    ```

2.  **Create a FIB Factory:**
    Instead of having a single `FileInformationBlock` class, use a factory pattern to create a version-specific FIB object. This will make the code much cleaner.

    ```csharp
    // In Doc/DocFileFormat/FileInformationBlock.cs
    public abstract class Fib
    {
        public short nFib { get; protected set; }
        public static Fib Create(byte[] fibBytes)
        {
            short version = BitConverter.ToInt16(fibBytes, 2);
            switch (version)
            {
                case 101: // Word 6.0
                case 102: // Word 95
                    return new FibWord6(fibBytes);
                case 103: // Word 97
                // ... other versions
                default:
                    return new FibWord97(fibBytes);
            }
        }
    }

    public class FibWord6 : Fib { /* ... Word 6 specific parsing ... */ }
    public class FibWord97 : Fib { /* ... Word 97 specific parsing ... */ }
    ```

### Phase 2: Implement Word 6.0/95 FIB Parsing

1.  **Map the Word 6.0/95 FIB Structure:**
    Using the `[MS-DOC]` specification and resources like the `WORD95_SUPPORT.md` file, map the fields of the older FIB format in the `FibWord6` class. The offsets and meanings of many fields will be different.

    ```csharp
    public class FibWord6 : Fib
    {
        // Word 6 has different offsets for these critical pointers
        public int fcMin { get; private set; }
        public int fcMac { get; private set; }

        public FibWord6(byte[] fibBytes)
        {
            // Read fields at their Word 6-specific offsets
            this.nFib = BitConverter.ToInt16(fibBytes, 2);
            this.fcMin = BitConverter.ToInt32(fibBytes, 24);
            this.fcMac = BitConverter.ToInt32(fibBytes, 28);
            // ... and so on for all other required fields
        }
    }
    ```

### Phase 3: Adapt the Rest of the Parser

With a version-specific FIB, the rest of the `WordDocument` parser needs to be adapted to handle the differences.

1.  **Abstract Away Version Differences:**
    The `WordDocument` class should not be filled with `if (fib.nFib == 101)` checks. Instead, the version-specific `Fib` object should provide the necessary information through a common interface or abstract properties.

    ```csharp
    public abstract class Fib
    {
        public abstract int GetTextStreamStart();
        public abstract int GetTextStreamLength();
    }

    public class FibWord6 : Fib
    {
        public override int GetTextStreamStart() { return this.fcMin; }
        // ...
    }

    public class FibWord97 : Fib
    {
        public override int GetTextStreamStart() { /* return Word 97 value */ }
        // ...
    }

    // In WordDocument.cs, the code becomes version-agnostic
    var textStart = _fib.GetTextStreamStart();
    ```

2.  **Handle Different Character Encodings:**
    Older Word versions may use different default character encodings (e.g., Windows-1252) instead of Unicode for the entire document. The text reading logic needs to be aware of this and use the correct `System.Text.Encoding`.

    ```csharp
    // In the text stream reader
    Encoding encoding = (_fib.IsUnicode) ? Encoding.Unicode : Encoding.GetEncoding(1252);
    var text = encoding.GetString(textBytes);
    ```

## Files to Investigate
- `Doc/DocFileFormat/FileInformationBlock.cs`: This will be the central point of the changes.
- `Doc/DocFileFormat/WordDocument.cs`: Will need to be adapted to use the new FIB factory and abstract methods.
- `Doc/DocFileFormat/Data/TextStream.cs`: May need to be updated to handle different encodings.
- `WORD95_SUPPORT.md`: A key source of information for the FIB layout.

## Success Criteria
1.  **Conversion Success**: The file `a-idade-media.word95.doc` can be converted to text without crashing.
2.  **Content Accuracy**: The extracted text from Word 6.0/95 documents is correct and complete.
3.  **No Regressions**: The changes do not break support for Word 97-2003 documents.
4.  **Clean Architecture**: The version-specific logic is well-encapsulated within the new `Fib` classes, avoiding `if/switch` statements scattered throughout the codebase.