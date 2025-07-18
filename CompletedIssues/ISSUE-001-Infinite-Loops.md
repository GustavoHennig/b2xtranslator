# Issue #001: Infinite Loops and Application Hanging (Partially fixed)

## Current Status
**Could not reproduce with the available files, but a detailed resolution plan is now available.**

## Problem Description
The b2xtranslator encounters infinite parsing loops that cause applications to hang indefinitely during document conversion. This is a critical issue that renders the application unusable for certain document types.

## Severity
**CRITICAL** - Application becomes unresponsive and must be forcefully terminated

## Impact
- Complete application freeze requiring process termination
- Loss of any unsaved work in consuming applications
- Poor user experience and reliability concerns
- Potential server/service instability in automated scenarios

## Root Cause Analysis
Infinite loops typically occur in recursive parsing operations within the binary file format readers. Common scenarios include:

1. **Circular Reference Loops**: Document structures with circular references that aren't properly detected
2. **Malformed Binary Data**: Corrupted or non-standard binary structures causing parsing logic to loop indefinitely
3. **Missing Loop Termination Conditions**: Parsing algorithms that lack proper exit conditions
4. **Recursive Depth Issues**: Deep recursion without stack overflow protection

## Known Affected Components
- **Doc Parser**: Word document binary parsing in `Doc/` project
- **Structured Storage Reader**: Compound document parsing in `Common.CompoundFileBinary/`
- **Record Parsing**: Binary record iteration logic across all format parsers

## Reproduction Steps

### Basic Reproduction
1. Build the solution:
   ```bash
   dotnet build b2xtranslator.sln
   ```

2. Attempt to convert a problematic document:
   ```bash
   dotnet run --project Shell/doc2text/doc2text.csproj -- samples/problematic.doc output.txt
   ```

3. Monitor application behavior - it will hang without progress indication

## Detailed Resolution Plan


### Phase 0: Do not execute the plan Phase 1, 2 and 3
 - Instead of workarounds, investigate each loop and recursion in the project, and identify the continuation criteria.
 - Handle all possible scenarios where a malformed input document has binary information that could cause an infinite loop.
 - Only when, for any reason, it is impossible to prevent the behavior, apply phases 1, 2, and 3.


### Phase 0.1: Analysis of Potential Infinite Loops

This phase focuses on identifying and documenting potential infinite loop vulnerabilities in the `b2xtranslator` codebase. The following sections detail the findings and recommended fixes.

#### `PieceTable.cs`

- **Vulnerability**: The `while(goon)` loop in the constructor is not robust against malformed input. If the `type` variable is not `1` or `2`, the loop will never terminate.
- **Fix**: Add a default case to the `switch` statement that sets `goon` to `false` and logs a warning. Additionally, add a bounds check to the `pos` variable to prevent `IndexOutOfRangeException`.

```csharp
// In Doc/DocFileFormat/PieceTable.cs

// ... existing code ...
while (goon)
{
    try
    {
        if (pos >= bytes.Length)
        {
            goon = false;
            break;
        }
        byte type = bytes[pos];

        //check if the type of the entry is a piece table
        if (type == 2)
        {
            // ... existing code ...
            goon = false;
        }
        //entry is no piecetable so goon
        else if (type == 1)
        {
            short cb = System.BitConverter.ToInt16(bytes, pos + 1);
            pos = pos + 1 + 2 + cb;
        }
        else
        {
            goon = false;
        }
    }
    catch(Exception)
    {
        goon = false;
    }
}
// ... existing code ...
```

#### `Plex.cs`

- **Vulnerability**: The calculation of `n` in the constructor can lead to an extremely large value if `lcb` is large and `structureLength` is 0, causing the application to hang.
- **Fix**: Add a sanity check to `n` to ensure it doesn't exceed a reasonable threshold.

```csharp
// In Doc/DocFileFormat/Plex.cs

// ... existing code ...
int n = 0;
if(structureLength > 0)
{
    //this PLEX contains CPs and Elements
    n = ((int)lcb - CP_LENGTH) / (structureLength + CP_LENGTH);
}
else
{
    //this PLEX only contains CPs
    n = ((int)lcb - CP_LENGTH) / CP_LENGTH;
}

if (n < 0 || n > 1000000) // Sanity check
{
    // Log a warning and return
    return;
}

//read the n + 1 CPs
// ... existing code ...
```

#### `ListTable.cs` and `ListFormatOverrideTable.cs`

- **Vulnerability**: The constructors in both classes read a `count` from the stream and then loop that many times. A malformed `count` can cause the application to hang.
- **Fix**: Add a sanity check to the `count` variable to prevent excessive looping.

```csharp
// In Doc/DocFileFormat/ListTable.cs

// ... existing code ...
short count = reader.ReadInt16();

if (count < 0 || count > 10000) // Sanity check
{
    return;
}

//read the LSTF structs
// ... existing code ...
```

```csharp
// In Doc/DocFileFormat/ListFormatOverrideTable.cs

// ... existing code ...
int count = reader.ReadInt32();

if (count < 0 || count > 10000) // Sanity check
{
    return;
}

//read the LFOs
// ... existing code ...
```

#### `FormattedDiskPagePAPX.cs` and `FormattedDiskPageCHPX.cs`

- **Vulnerability**: The `GetAll...` methods calculate `n` based on the `lcb...` value from the FIB. A malformed `lcb...` value can lead to a very large `n`, causing a hang.
- **Fix**: Add a sanity check to `n` to prevent excessive looping.

```csharp
// In Doc/DocFileFormat/FormattedDiskPagePAPX.cs

// ... existing code ...
int n = (((int)fib.lcbPlcfBtePapx - 4) / 8) + 1;

if (n < 0 || n > 1000000) // Sanity check
{
    return list;
}

//Get the indexed PAPX FKPs
// ... existing code ...
```

```csharp
// In Doc/DocFileFormat/FormattedDiskPageCHPX.cs

// ... existing code ...
int n = (((int)fib.lcbPlcfBteChpx - 4) / 8) + 1;

if (n < 0 || n > 1000000) // Sanity check
{
    return list;
}

//Get the indexed CHPX FKPs
// ... existing code ...
```

#### `WordDocument.cs`

- **Vulnerability**: Several loops in the `Parse` method are susceptible to hangs if the counts read from the file are malformed.
- **Fix**: Add sanity checks to the loops that build the `AllPapx` and `AllSepx` dictionaries.

```csharp
// In Doc/DocFileFormat/WordDocument.cs

// ... existing code ...
//build a dictionaries of all PAPX
this.AllPapx = new Dictionary<int, ParagraphPropertyExceptions>();
for (int i = 0; i < this.AllPapxFkps.Count; i++)
{
    if (i > 10000) break; // Sanity check
    for (int j = 0; j < this.AllPapxFkps[i].grppapx.Length; j++)
    {
        if (j > 10000) break; // Sanity check
        this.AllPapx.Add(this.AllPapxFkps[i].rgfc[j], this.AllPapxFkps[i].grppapx[j]);
    }
}

// ... existing code ...

//build a dictionary of all SEPX
this.AllSepx = new Dictionary<int, SectionPropertyExceptions>();
for (int i = 0; i < this.SectionPlex.Elements.Count; i++)
{
    if (i > 10000) break; // Sanity check
    // ... existing code ...
}
// ... existing code ...
```


### Phase 1: Introduce a Parsing Context

1.  **Create a `ParsingContext` Class:**
    In the `Common.Abstractions` project, create a new class `ParsingContext`. This class will act as a state manager for a single conversion operation.

    ```csharp
    namespace b2xtranslator.Common.Abstractions
    {
        public class ParsingContext
        {
            public int RecursionDepth { get; private set; }
            private const int MaxRecursionDepth = 100; // Configurable

            private readonly HashSet<object> _visitedObjects = new HashSet<object>();

            public void EnterRecursion()
            {
                RecursionDepth++;
                if (RecursionDepth > MaxRecursionDepth)
                {
                    throw new InvalidOperationException("Maximum recursion depth exceeded.");
                }
            }

            public void LeaveRecursion()
            {
                RecursionDepth--;
            }

            public bool HasVisited(object obj)
            {
                return _visitedObjects.Contains(obj);
            }

            public void AddVisited(object obj)
            {
                _visitedObjects.Add(obj);
            }
        }
    }
    ```

2.  **Integrate `ParsingContext` into the `WordDocument` Parser:**
    Modify the `WordDocument` constructor and other relevant parsing classes to accept a `ParsingContext` instance. This context will be passed down through the entire parsing chain.

    ```csharp
    // In Doc/DocFileFormat/WordDocument.cs
    public WordDocument(StructuredStorageReader storage, ParsingContext context)
    {
        _storage = storage;
        _context = context;
        // ... rest of the constructor
    }
    ```

### Phase 2: Implement Loop and Recursion Detection

1.  **Track Recursion Depth:**
    In every recursive method within the parsers (e.g., methods that parse nested structures), use the `ParsingContext` to track the recursion depth.

    ```csharp
    // Example in a hypothetical recursive parsing method
    private void ParseRecursiveStructure(SomeObject parent)
    {
        _context.EnterRecursion();
        try
        {
            // ... parsing logic ...
            foreach (var child in parent.Children)
            {
                ParseRecursiveStructure(child);
            }
        }
        finally
        {
            _context.LeaveRecursion();
        }
    }
    ```

2.  **Detect Circular References:**
    Before processing an object, check if it has already been visited using the `ParsingContext`. This is particularly important when parsing linked structures like lists or fields.

    ```csharp
    // Example in a list parsing method
    private void ParseList(ListObject list)
    {
        if (_context.HasVisited(list))
        {
            // Log a warning and skip this list to prevent a loop
            return;
        }
        _context.AddVisited(list);

        // ... rest of the parsing logic ...
    }
    ```

### Phase 3: Implement Timeout and Cancellation

1.  **Introduce `CancellationToken`:**
    Update the main conversion methods to accept a `CancellationToken`.

    ```csharp
    // In Shell/doc2text/Program.cs
    public static void Main(string[] args)
    {
        var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60)); // 60-second timeout
        try
        {
            // ...
            TextConverter.Convert(doc, writer, cts.Token);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Conversion timed out.");
        }
    }

    // In Text/TextMapping/TextConverter.cs
    public static void Convert(WordDocument doc, TextWriter writer, CancellationToken cancellationToken)
    {
        // ...
        // Periodically check the token in long-running loops
        cancellationToken.ThrowIfCancellationRequested();
    }
    ```

2.  **Propagate the `CancellationToken`:**
    Pass the `CancellationToken` down to all major parsing and mapping components. Inside loops that iterate over large collections (e.g., paragraphs, records, characters), check for cancellation requests.

### Phase 4: Add Robust Error Handling

1.  **Wrap Critical Parsing Logic:**
    In low-level parsing code that reads from byte arrays or streams, wrap the logic in `try-catch` blocks to handle unexpected data gracefully.

    ```csharp
    // In a binary record parser
    public override void Read(byte[] data, int offset)
    {
        try
        {
            // ... read from data array ...
        }
        catch (IndexOutOfRangeException ex)
        {
            // Log the error and attempt to recover or skip the record
            _logger.LogWarning(ex, "Malformed record encountered. Skipping.");
        }
    }
    ```

## Files to Investigate
- `Doc/DocFileFormat/WordDocument.cs`
- `Common.CompoundFileBinary/StructuredStorage/StructuredStorageReader.cs`
- `Doc/DocFileFormat/Data/Plex.cs` (often involved in loops)
- `Doc/DocFileFormat/Data/FileInformationBlock.cs`
- All classes that perform recursive parsing.

## Success Criteria
1. **No Infinite Loops**: All conversion operations complete within reasonable time limits.
2. **Timeout Protection**: Operations automatically terminate after a configurable timeout period.
3. **Graceful Failure**: The application reports a clear error instead of hanging on problematic files.
4. **No Regressions**: The changes do not negatively impact the conversion of valid files.
