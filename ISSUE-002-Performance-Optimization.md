# Issue #002: Performance Optimization - Slow Conversion Scenarios

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator experiences extremely slow conversion times for certain document types, taking minutes or hours to complete operations that should finish in seconds. This significantly impacts user experience and system scalability.

## Severity
**HIGH** - Severely impacts usability and system performance

## Impact
- Unacceptable conversion times for end users
- Server timeout issues in automated processing
- Resource starvation in multi-user environments  
- Poor scalability for batch processing operations
- Increased infrastructure costs due to inefficient resource usage

## Root Cause Analysis
Performance issues typically stem from:

1. **Inefficient Parsing Algorithms**: O(n²) or worse complexity in document traversal
2. **Excessive Memory Allocations**: Frequent object creation and garbage collection pressure
3. **Redundant Processing**: Repeated parsing of the same document sections
4. **Blocking I/O Operations**: Synchronous file operations without optimization
5. **Large Document Handling**: Poor algorithms for processing large or complex documents

## Detailed Resolution Plan

### Phase 0: doc2text Pipeline Optimization

1.  **Record Type Filtering:** During the initial read, inspect record headers and process only text-related records (e.g., character, paragraph, and styling records). Skip images, drawings, and metadata not relevant to doc2text.  
2.  **Indexed Record Parsing:** Pre-scan record offsets to build an index of text-containing records. Iterate this index directly to reduce full document traversals.  
3.  **Delegate Caching:** Map record types to pre-compiled parsing delegates to avoid expensive type lookups or reflection at runtime.  
4.  **Batch Text Flushing:** Accumulate extracted text in a buffer or `StringBuilder`, flushing to the output writer in large chunks to minimize I/O calls.  
5.  **Compiled Regex Cache:** For post-processing (e.g., whitespace normalization, hyphen removal), use pre-compiled `Regex` instances stored in a static cache.  
6.  **Zero-Copy Span Access:** When slicing record data, use `ReadOnlySpan<byte>` and `MemoryMarshal` to parse text directly without additional allocations.  
7.  **Lightweight Visitor Pattern:** Implement a visitor interface for record processing to inline the dispatch logic and reduce virtual call overhead.  


### Phase 1: Benchmarking and Profiling


 ## Profiling insight 1:

The main reason your code is running slower than expected in GetCharacterPropertyExceptions is due to the nested iteration over all elements in AllChpxFkps and then over each grpchpx array. For every FKP, you loop through its character property exceptions and perform range checks, which can be costly if the data set is large. This approach results in a time complexity proportional to the total number of FKPs multiplied by the number of property exceptions per FKP. If AllChpxFkps or the grpchpx arrays are large, this can quickly become a bottleneck. Additionally, the method creates a new list and adds references to property exceptions, which is necessary, but if the same exceptions are repeatedly added or if the range checks are not optimal, this can lead to redundant work.
There is no obvious redundancy in storing data, but you could benefit from using a more efficient data structure if you know the ranges are sorted or if you can index into the FKPs more directly. For example, if FKPs are sorted by their range, a binary search could help you quickly locate the relevant FKPs and reduce the number of iterations. Also, if you are always searching for exceptions within a contiguous range, you might consider caching or indexing the FKPs for faster lookup.
Enumerable calls are not used in this method, but if you are using LINQ or similar constructs elsewhere, replacing them with for-loops can sometimes improve performance, especially in tight loops. Converting data to a different type is not directly applicable here, but if you can represent the ranges as intervals and use an interval tree or similar structure, you could speed up the search for relevant property exceptions.
Looking at the parent method, writeParagraph, it calls GetCharacterPropertyExceptions and then iterates over the results to write runs. If you can reduce the number of property exceptions returned, or if you can avoid unnecessary splitting of runs (for example, by merging adjacent runs with identical formatting), you could further improve performance. Also, if the parent method repeatedly calls GetCharacterPropertyExceptions for overlapping ranges, consider caching results or reusing data where possible.
Optimize b2xtranslator.DocFileFormat.WordDocument.GetCharacterPropertyExceptions(int, int):
•	Use binary search for FKPs: If FKPs are sorted, use binary search to quickly locate relevant FKPs and reduce unnecessary iterations.
•	Merge adjacent identical exceptions: Avoid adding multiple property exceptions for contiguous ranges with identical formatting to reduce list size.
•	Index FKPs for fast lookup: Build an index or cache for FKPs to allow direct access to relevant property exceptions, minimizing iteration.
Optimize b2xtranslator.txt.TextMapping.DocumentMapping.writeParagraph(int, int, bool):
•	Cache property exceptions: If paragraphs overlap or are processed sequentially, cache results from GetCharacterPropertyExceptions to avoid redundant computation.


## Gather more info:

1.  **Establish a Benchmark Suite:**
    Create a dedicated `Benchmark` project using a library like `BenchmarkDotNet`. This will provide a stable, repeatable way to measure performance.

    ```csharp
    // In a new project Benchmarks/Benchmark.csproj
    [MemoryDiagnoser]
    public class ConversionBenchmarks
    {
        [Benchmark]
        [Arguments("samples/large.doc")]
        public void ConvertLargeDoc(string path)
        {
            // ... conversion logic ...
        }
    }
    ```

2.  **Profile the Application:**
    Use `dotnet-trace` to profile the application and identify the hottest code paths. The focus should be on both CPU usage and memory allocations.

    ```bash
    # Profile CPU
    dotnet-trace collect --process-id <pid> --format speedscope

    # Profile Memory
    dotnet-trace collect --process-id <pid> --providers Microsoft-DotNETCore-GC --format speedscope
    ```

### Phase 2: Low-Hanging Fruit - Memory Allocations

Based on typical performance issues in file parsing, the first area to tackle is memory allocation.

1.  **Introduce `ArrayPool<T>` for Buffers:**
    Many parts of the code likely use `new byte[]` for temporary buffers. Replace these with `ArrayPool<T>.Shared` to reduce GC pressure.

    ```csharp
    // Before
    var buffer = new byte[4096];

    // After
    var buffer = ArrayPool<byte>.Shared.Rent(4096);
    try
    {
        // ... use buffer ...
    }
    finally
    {
        ArrayPool<byte>.Shared.Return(buffer);
    }
    ```
    **Files to investigate:** `StructuredStorageReader.cs`, `VirtualStream.cs`, and any other class that performs manual stream reading.

2.  **Use `Span<T>` and `Memory<T>`:**
    For operations that involve slicing or viewing parts of byte arrays, use `Span<T>` to avoid creating new array copies.

    ```csharp
    // Before
    var subArray = new byte[10];
    Array.Copy(original, 5, subArray, 0, 10);

    // After
    var span = new Span<byte>(original, 5, 10);
    ```
    **Files to investigate:** All binary record parsing classes in `Doc/`, `Ppt/`, and `Xls/`.

### Phase 3: Algorithmic Optimization

1.  **Optimize Stream Reading:**
    Ensure that file streams are read in large, efficient chunks rather than byte by byte. Review the `VirtualStream` implementation for inefficiencies.

2.  **Implement Lazy Loading:**
    Many document parts may not be needed for all conversion types. For example, `doc2text` does not need image data. Modify the `WordDocument` and other parser classes to only load data from the underlying storage when it is explicitly requested.

    ```csharp
    // Before
    public class WordDocument
    {
        public byte[] ImageData { get; private set; }
        public WordDocument(StructuredStorageReader storage)
        {
            // Eagerly loads image data
            ImageData = storage.ReadStream("Pictures");
        }
    }

    // After
    public class WordDocument
    {
        private readonly Lazy<byte[]> _imageData;
        public byte[] ImageData => _imageData.Value;

        public WordDocument(StructuredStorageReader storage)
        {
            _imageData = new Lazy<byte[]>(() => storage.ReadStream("Pictures"));
        }
    }
    ```

### Phase 4: I/O and Concurrency

1.  **Introduce Async Methods:**
    For library consumers who may be using the translator in a server environment, provide asynchronous versions of the main conversion methods.

    ```csharp
    // In TextConverter.cs
    public static async Task ConvertAsync(WordDocument doc, TextWriter writer, CancellationToken cancellationToken)
    {
        // ...
        await writer.WriteAsync(...);
    }
    ```

## Files to Investigate
- `Common.CompoundFileBinary/StructuredStorage/StructuredStorageReader.cs` - File I/O operations
- `Common/OpenXmlLib/OpenXmlWriter.cs` - XML generation performance
- `*Mapping/` - Conversion mapping performance
- Binary parsing loops in `Doc/`, `Ppt/`, `Xls/` projects

## Success Criteria
1. **Response Time**: 90% improvement in conversion times for slow scenarios.
2. **Memory Efficiency**: 50% reduction in peak memory usage.
3. **Regression Testing**: Automated performance monitoring prevents regressions.
4. **Scalability**: Linear performance scaling with document size.
