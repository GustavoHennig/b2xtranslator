# Issue #003: Memory Management and Leak Issues

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator experiences memory leaks and excessive memory consumption during large file processing, leading to system instability, out-of-memory exceptions, and resource exhaustion in production environments.

## Severity
**HIGH** - Can cause system-wide resource exhaustion and application crashes

## Impact
- Out-of-memory exceptions during large file processing
- System instability and potential crashes
- Resource starvation affecting other applications
- Server instability in multi-user environments
- Poor scalability for batch processing operations
- Increased infrastructure costs due to excessive memory usage

## Detailed Resolution Plan

### Phase 1: Diagnostics and Tooling

1.  **Memory Leak Test Suite:**
    Create a new test project (`MemoryTests`) dedicated to detecting memory leaks. This test will repeatedly convert a document and assert that memory usage does not grow indefinitely.

    ```csharp
    // In MemoryTests/MemoryLeakTests.cs
    [Test]
    public void ConversionLoop_ShouldNotLeakMemory()
    {
        // Force a GC to get a clean baseline
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long initialMemory = GC.GetTotalMemory(false);

        for (int i = 0; i < 100; i++)
        {
            // Each conversion should be isolated
            using (var storage = new StructuredStorageReader("samples/test.doc"))
            using (var doc = new WordDocument(storage))
            using (var writer = new StringWriter())
            {
                TextConverter.Convert(doc, writer);
            }
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        long finalMemory = GC.GetTotalMemory(false);

        // Allow for a small, acceptable amount of memory growth
        long growth = finalMemory - initialMemory;
        Assert.That(growth, Is.LessThan(5 * 1024 * 1024)); // 5MB threshold
    }
    ```

2.  **Analyze Memory Dumps:**
    Use `dotnet-dump` to capture and analyze memory dumps of the application while it is processing a large file. The `dumpheap -stat` command is crucial for identifying which object types are consuming the most memory.

    ```bash
    dotnet-dump collect -p <pid>
    dotnet-dump analyze <dump_file> -c "dumpheap -stat"
    ```

### Phase 2: Implement IDisposable Correctly

1.  **Audit `IDisposable` Usage:**
    Review every class that implements `IDisposable` to ensure the pattern is implemented correctly. Pay close attention to finalizers and the disposal of both managed and unmanaged resources.

    ```csharp
    public class MyResourceHolder : IDisposable
    {
        private bool _disposed = false;
        private Stream _managedResource;
        private IntPtr _unmanagedResource;

        // ...

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Dispose managed resources
                _managedResource?.Dispose();
            }

            // Dispose unmanaged resources
            CloseHandle(_unmanagedResource);

            _disposed = true;
        }

        ~MyResourceHolder() { Dispose(false); }
    }
    ```

2.  **Ensure `using` Statements:**
    Everywhere an `IDisposable` object is instantiated, ensure it is wrapped in a `using` statement or that its `Dispose` method is called in a `finally` block.

    **Files to investigate:** All `Main` methods in `Shell/` projects, and any method that creates `Stream` or `*Reader`/`*Writer` objects.

### Phase 3: Optimize Memory-Intensive Operations

1.  **Replace `new byte[]` with `ArrayPool<T>`:**
    As detailed in `ISSUE-002`, use `ArrayPool<T>` for all temporary buffers to avoid frequent GC pressure from large array allocations.

2.  **Refactor Large Object Handling:**
    Identify areas where large objects (like the entire document model) are held in memory. Refactor these to use streaming or lazy-loading approaches. For example, instead of loading all pictures into a `List<byte[]>`, create a `List<PictureInfo>` that can load the picture data on demand.

    ```csharp
    // Before
    class WordDocument { public List<byte[]> Pictures; }

    // After
    class PictureInfo { public Func<byte[]> PictureBytes { get; set; } }
    class WordDocument { public List<PictureInfo> Pictures; }
    ```

### Phase 4: Break Object Reference Cycles

1.  **Identify and Break Circular References:**
    Memory leaks are often caused by circular references (e.g., a parent object references a child, and the child references the parent). When the `MemoryLeakTests` fail, use a memory profiler to inspect the object graph and find these cycles. They can often be broken by using `WeakReference<T>` or by clearing references upon disposal.

    ```csharp
    class Child
    {
        // Use WeakReference to avoid a strong cycle
        public WeakReference<Parent> Parent { get; set; }
    }
    ```

## Files to Investigate
- `Common.CompoundFileBinary/StructuredStorage/StructuredStorageReader.cs` - Stream and buffer management
- `Common/OpenXmlLib/OpenXmlWriter.cs` - XML object lifecycle
- `*Mapping/` - Conversion context and object management
- `Doc/DocFileFormat/WordDocument.cs` - Document object model lifecycle
- All `IDisposable` implementations across the solution.

## Success Criteria
1. **No Memory Leaks**: The `MemoryLeakTests` suite passes consistently.
2. **Bounded Memory Growth**: Memory usage does not grow unbounded over time in the shell applications.
3. **Efficient GC**: The number of Gen 2 garbage collections is minimized during typical conversions.
4. **Resource Cleanup**: All file handles and other unmanaged resources are properly disposed of, even when exceptions occur.
