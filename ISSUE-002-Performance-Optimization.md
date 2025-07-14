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

### Phase 1: Benchmarking and Profiling

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
