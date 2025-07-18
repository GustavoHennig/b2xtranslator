# Issue #013: Improve Error Handling and Reliability

## Problem Description
The translator can be brittle when encountering corrupted, malformed, or non-standard binary files. Instead of gracefully reporting an error or attempting a partial conversion, it may crash with an unhandled exception (e.g., `NullReferenceException`, `IndexOutOfRangeException`).

## Severity
**HIGH** - Leads to application crashes and makes the library feel unreliable.

## Impact
- Unreliable conversions for any file that isn't perfectly formed.
- Consuming applications may crash, leading to data loss or service outages.
- Difficult to use in automated batch processing, as a single bad file can halt the entire process.

## Root Cause Analysis
The parsing logic often lacks sufficient validation and error-trapping. It proceeds with the assumption that the file structure conforms to the specification.

1.  **Insufficient Validation**: File headers, stream sizes, and internal data structure offsets are not always validated before being used.
2.  **Missing Error Boundaries**: Critical parsing sections are not wrapped in `try-catch` blocks, allowing low-level exceptions to bubble up and crash the application.
3.  **No Recovery Mechanism**: There is no concept of a "tolerant" or "best-effort" parsing mode that would attempt to skip over corrupted parts of a file.

## Known Affected Components
- All parsing libraries are potentially affected:
    - `Common.CompoundFileBinary/`
    - `Doc/`
    - `Ppt/`
    - `Xls/`

## Reproduction Steps

### Basic Reproduction
1.  Obtain or create a corrupted/truncated `.doc` file.
2.  Attempt to convert it:
    ```bash
    dotnet run --project Shell/doc2text/doc2text.csproj -- samples-local/corrupted.doc output.txt
    ```
3.  Observe the application crash with an unhandled exception.

## Testing Strategy

### Integration Tests
- Create a dedicated suite of tests for error handling.
- This suite should use a collection of known-bad files (truncated, corrupted, zero-byte, etc.).
- The tests should assert that converting these files does **not** throw an unhandled exception. Instead, the conversion should either complete (possibly with partial results) or throw a specific, expected exception (e.g., `InvalidFileFormatException`).

## Proposed Solutions

### 1. Add Pre-emptive Validation
- Before starting detailed parsing, add checks for file signatures, minimum file sizes, and the integrity of the compound file directory.

### 2. Implement Robust Error Boundaries
- Identify critical, high-risk parsing operations (e.g., reading from a stream based on a calculated offset).
- Wrap these operations in `try-catch` blocks.
- When an exception is caught, log it as a warning and, where possible, allow processing to continue (e.g., by skipping the problematic element).

### 3. Introduce Specific Exceptions
- Create custom exception types (e.g., `CorruptFileStructureException`, `UnsupportedFeatureException`) to provide more meaningful diagnostics than generic system exceptions.

### 4. Graceful Degradation
- For non-critical errors, the parser should be able to log the issue and continue, yielding a partial but still useful result.

## Files to Investigate
- `Common.CompoundFileBinary/StructuredStorage/StructuredStorageReader.cs`
- `Doc/DocFileFormat/WordDocument.cs`
- `Doc/DocFileFormat/FileInformationBlock.cs`
- Any code that involves reading from `byte[]` or `Stream` objects based on offsets read from the file itself.


## Current Status (Verified)
**PARTIALLY IMPLEMENTED** - Testing shows basic error handling exists but needs enhancement:
- Graceful handling for completely invalid files: "Input file test-corrupted.doc is not a valid Microsoft Word 97-2003 file."
- However, detailed resolution plan needed for edge cases and robustness improvements
- More sophisticated error handling and recovery mechanisms still required

## Enhanced Resolution Plan

### Phase 1: Comprehensive Error Classification (1-2 days)

#### 1.1 Expand Custom Exception Hierarchy
**File: `Common.Abstractions/Exceptions/ParsingExceptions.cs`** (create/enhance)

```csharp
namespace b2xtranslator.Common.Abstractions.Exceptions
{
    // Base exception for all parsing-related errors
    public abstract class ParsingException : Exception
    {
        public string DocumentPath { get; }
        public long? FilePosition { get; }
        
        protected ParsingException(string message, string documentPath = null, long? filePosition = null) 
            : base(message)
        {
            DocumentPath = documentPath;
            FilePosition = filePosition;
        }
        
        protected ParsingException(string message, Exception innerException, string documentPath = null, long? filePosition = null) 
            : base(message, innerException)
        {
            DocumentPath = documentPath;
            FilePosition = filePosition;
        }
    }
    
    // Specific exception types
    public class MalformedFileException : ParsingException
    {
        public MalformedFileException(string message, string documentPath = null) 
            : base($"File appears to be corrupted or malformed: {message}", documentPath) { }
    }
    
    public class UnsupportedFileVersionException : ParsingException
    {
        public string DetectedVersion { get; }
        
        public UnsupportedFileVersionException(string detectedVersion, string documentPath = null) 
            : base($"Unsupported file version: {detectedVersion}", documentPath)
        {
            DetectedVersion = detectedVersion;
        }
    }
    
    public class InsufficientDataException : ParsingException
    {
        public int RequiredBytes { get; }
        public int AvailableBytes { get; }
        
        public InsufficientDataException(int required, int available, long position, string documentPath = null)
            : base($"Insufficient data: required {required} bytes, only {available} available at position {position}", documentPath, position)
        {
            RequiredBytes = required;
            AvailableBytes = available;
        }
    }
    
    public class CorruptedStructureException : ParsingException
    {
        public string StructureName { get; }
        
        public CorruptedStructureException(string structureName, string details, string documentPath = null, long? position = null)
            : base($"Corrupted {structureName}: {details}", documentPath, position)
        {
            StructureName = structureName;
        }
    }
    
    public class RecursionLimitException : ParsingException
    {
        public int MaxDepth { get; }
        
        public RecursionLimitException(int maxDepth, string documentPath = null)
            : base($"Maximum recursion depth exceeded: {maxDepth}", documentPath)
        {
            MaxDepth = maxDepth;
        }
    }
}
```

#### 1.2 Enhanced Validation Framework
**File: `Common.Abstractions/Validation/DocumentValidator.cs`** (create)

```csharp
public static class DocumentValidator
{
    public static ValidationResult ValidateCompoundFile(string filePath)
    {
        var result = new ValidationResult();
        
        try
        {
            // File existence and basic checks
            if (!File.Exists(filePath))
            {
                result.AddError("File does not exist");
                return result;
            }
            
            var fileInfo = new FileInfo(filePath);
            if (fileInfo.Length < 512)
            {
                result.AddError($"File too small: {fileInfo.Length} bytes (minimum 512 required)");
                return result;
            }
            
            // Compound file signature validation
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var signature = new byte[8];
            if (stream.Read(signature, 0, 8) != 8)
            {
                result.AddError("Cannot read file signature");
                return result;
            }
            
            var expectedSignature = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            if (!signature.SequenceEqual(expectedSignature))
            {
                result.AddError($"Invalid compound file signature: {BitConverter.ToString(signature)}");
                return result;
            }
            
            // Additional header validation
            stream.Seek(22, SeekOrigin.Begin);
            var minorVersion = new byte[2];
            stream.Read(minorVersion, 0, 2);
            var majorVersion = new byte[2];
            stream.Read(majorVersion, 0, 2);
            
            result.AddInfo($"File version: {BitConverter.ToUInt16(majorVersion, 0)}.{BitConverter.ToUInt16(minorVersion, 0)}");
            
        }
        catch (Exception ex)
        {
            result.AddError($"Validation failed: {ex.Message}");
        }
        
        return result;
    }
    
    public static ValidationResult ValidateWordDocument(Stream documentStream)
    {
        var result = new ValidationResult();
        
        try
        {
            // Validate FIB (File Information Block)
            if (documentStream.Length < 1024)
            {
                result.AddError("Document stream too small for valid Word document");
                return result;
            }
            
            documentStream.Seek(0, SeekOrigin.Begin);
            var fibHeader = new byte[32];
            documentStream.Read(fibHeader, 0, 32);
            
            // Check FIB signature
            var fibSignature = BitConverter.ToUInt16(fibHeader, 0);
            if (fibSignature != 0xA5EC && fibSignature != 0xA5DB)
            {
                result.AddWarning($"Unexpected FIB signature: 0x{fibSignature:X4}");
            }
            
            // Check document version
            var nFib = BitConverter.ToUInt16(fibHeader, 2);
            result.AddInfo($"Document version (nFib): {nFib}");
            
            if (nFib < 101 || nFib > 1000)
            {
                result.AddWarning($"Unusual document version: {nFib}");
            }
            
        }
        catch (Exception ex)
        {
            result.AddError($"Document validation failed: {ex.Message}");
        }
        
        return result;
    }
}

public class ValidationResult
{
    public List<string> Errors { get; } = new List<string>();
    public List<string> Warnings { get; } = new List<string>();
    public List<string> Info { get; } = new List<string>();
    
    public bool IsValid => Errors.Count == 0;
    public bool HasWarnings => Warnings.Count > 0;
    
    public void AddError(string message) => Errors.Add(message);
    public void AddWarning(string message) => Warnings.Add(message);
    public void AddInfo(string message) => Info.Add(message);
}
```

### Phase 2: Safe Binary Reading Infrastructure (1-2 days)

#### 2.1 Enhanced Safe Reading Utilities
**File: `Common/Tools/SafeBinaryReader.cs`** (create)

```csharp
public static class SafeBinaryReader
{
    public static byte ReadByte(byte[] buffer, ref int offset)
    {
        ValidateBufferAccess(buffer, offset, 1);
        return buffer[offset++];
    }
    
    public static short ReadInt16(byte[] buffer, ref int offset)
    {
        ValidateBufferAccess(buffer, offset, 2);
        var result = BitConverter.ToInt16(buffer, offset);
        offset += 2;
        return result;
    }
    
    public static int ReadInt32(byte[] buffer, ref int offset)
    {
        ValidateBufferAccess(buffer, offset, 4);
        var result = BitConverter.ToInt32(buffer, offset);
        offset += 4;
        return result;
    }
    
    public static uint ReadUInt32(byte[] buffer, ref int offset)
    {
        ValidateBufferAccess(buffer, offset, 4);
        var result = BitConverter.ToUInt32(buffer, offset);
        offset += 4;
        return result;
    }
    
    public static byte[] ReadBytes(byte[] buffer, ref int offset, int count)
    {
        ValidateBufferAccess(buffer, offset, count);
        var result = new byte[count];
        Array.Copy(buffer, offset, result, 0, count);
        offset += count;
        return result;
    }
    
    public static string ReadString(byte[] buffer, ref int offset, int length, Encoding encoding)
    {
        ValidateBufferAccess(buffer, offset, length);
        var result = encoding.GetString(buffer, offset, length);
        offset += length;
        return result;
    }
    
    private static void ValidateBufferAccess(byte[] buffer, int offset, int count)
    {
        if (buffer == null)
            throw new ArgumentNullException(nameof(buffer));
            
        if (offset < 0)
            throw new InsufficientDataException(count, 0, offset);
            
        if (offset + count > buffer.Length)
            throw new InsufficientDataException(count, buffer.Length - offset, offset);
    }
    
    // Stream-based safe reading
    public static byte[] ReadBytesFromStream(Stream stream, int count)
    {
        if (stream.Position + count > stream.Length)
            throw new InsufficientDataException(count, (int)(stream.Length - stream.Position), stream.Position);
            
        var buffer = new byte[count];
        var bytesRead = stream.Read(buffer, 0, count);
        
        if (bytesRead != count)
            throw new InsufficientDataException(count, bytesRead, stream.Position - bytesRead);
            
        return buffer;
    }
}
```

#### 2.2 Error Recovery Context
**File: `Common.Abstractions/ErrorRecoveryContext.cs`** (create)

```csharp
public class ErrorRecoveryContext
{
    public string DocumentPath { get; }
    public bool AllowPartialRecovery { get; set; } = true;
    public int MaxRecoveryAttempts { get; set; } = 3;
    public TimeSpan MaxRecoveryTime { get; set; } = TimeSpan.FromSeconds(30);
    
    private readonly List<RecoveryAttempt> _recoveryLog = new List<RecoveryAttempt>();
    private readonly Stopwatch _recoveryTimer = new Stopwatch();
    
    public ErrorRecoveryContext(string documentPath)
    {
        DocumentPath = documentPath;
        _recoveryTimer.Start();
    }
    
    public bool CanAttemptRecovery()
    {
        return AllowPartialRecovery && 
               _recoveryLog.Count < MaxRecoveryAttempts && 
               _recoveryTimer.Elapsed < MaxRecoveryTime;
    }
    
    public void LogRecoveryAttempt(string operation, Exception error, bool successful)
    {
        _recoveryLog.Add(new RecoveryAttempt
        {
            Operation = operation,
            Error = error,
            Successful = successful,
            Timestamp = DateTime.UtcNow
        });
    }
    
    public RecoveryReport GenerateReport()
    {
        return new RecoveryReport
        {
            DocumentPath = DocumentPath,
            TotalAttempts = _recoveryLog.Count,
            SuccessfulAttempts = _recoveryLog.Count(r => r.Successful),
            TotalTime = _recoveryTimer.Elapsed,
            Attempts = _recoveryLog.ToList()
        };
    }
}

public class RecoveryAttempt
{
    public string Operation { get; set; }
    public Exception Error { get; set; }
    public bool Successful { get; set; }
    public DateTime Timestamp { get; set; }
}

public class RecoveryReport
{
    public string DocumentPath { get; set; }
    public int TotalAttempts { get; set; }
    public int SuccessfulAttempts { get; set; }
    public TimeSpan TotalTime { get; set; }
    public List<RecoveryAttempt> Attempts { get; set; }
}
```

### Phase 3: Intelligent Error Recovery (2-3 days)

#### 3.1 Tolerant Parsing Framework
**File: `Common.Abstractions/Parsing/TolerantParser.cs`** (create)

```csharp
public abstract class TolerantParser<T>
{
    protected ErrorRecoveryContext RecoveryContext { get; }
    protected ILogger Logger { get; }
    
    protected TolerantParser(ErrorRecoveryContext recoveryContext, ILogger logger)
    {
        RecoveryContext = recoveryContext;
        Logger = logger;
    }
    
    public ParseResult<T> TryParse(Stream stream)
    {
        var result = new ParseResult<T>();
        
        try
        {
            result.Value = ParseCore(stream);
            result.Success = true;
        }
        catch (ParsingException ex) when (RecoveryContext.CanAttemptRecovery())
        {
            Logger.LogWarning($"Parsing failed, attempting recovery: {ex.Message}");
            
            var recoveryResult = AttemptRecovery(stream, ex);
            RecoveryContext.LogRecoveryAttempt(typeof(T).Name, ex, recoveryResult.Success);
            
            if (recoveryResult.Success)
            {
                result.Value = recoveryResult.Value;
                result.Success = true;
                result.WasRecovered = true;
                result.RecoveryDetails = recoveryResult.RecoveryDetails;
            }
            else
            {
                result.Error = ex;
                result.Success = false;
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Unrecoverable parsing error: {ex.Message}");
            result.Error = ex;
            result.Success = false;
        }
        
        return result;
    }
    
    protected abstract T ParseCore(Stream stream);
    
    protected virtual ParseResult<T> AttemptRecovery(Stream stream, ParsingException originalError)
    {
        // Default implementation: try to skip to next known structure
        return new ParseResult<T> { Success = false };
    }
}

public class ParseResult<T>
{
    public T Value { get; set; }
    public bool Success { get; set; }
    public bool WasRecovered { get; set; }
    public Exception Error { get; set; }
    public string RecoveryDetails { get; set; }
}
```

#### 3.2 Specific Recovery Strategies
**File: `Doc/DocFileFormat/TolerantWordDocumentParser.cs`** (create)

```csharp
public class TolerantWordDocumentParser : TolerantParser<WordDocument>
{
    public TolerantWordDocumentParser(ErrorRecoveryContext context, ILogger logger) 
        : base(context, logger) { }
    
    protected override WordDocument ParseCore(Stream stream)
    {
        // Standard parsing logic
        return new WordDocument(stream);
    }
    
    protected override ParseResult<WordDocument> AttemptRecovery(Stream stream, ParsingException originalError)
    {
        Logger.LogInformation($"Attempting Word document recovery for: {originalError.Message}");
        
        // Strategy 1: Try parsing with minimal structure
        try
        {
            var partialDoc = ParseMinimalStructure(stream);
            return new ParseResult<WordDocument>
            {
                Value = partialDoc,
                Success = true,
                RecoveryDetails = "Recovered using minimal structure parsing"
            };
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Minimal structure recovery failed: {ex.Message}");
        }
        
        // Strategy 2: Try to extract just text runs
        try
        {
            var textOnlyDoc = ParseTextRunsOnly(stream);
            return new ParseResult<WordDocument>
            {
                Value = textOnlyDoc,
                Success = true,
                RecoveryDetails = "Recovered text content only"
            };
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Text-only recovery failed: {ex.Message}");
        }
        
        return new ParseResult<WordDocument> { Success = false };
    }
    
    private WordDocument ParseMinimalStructure(Stream stream)
    {
        // Implement minimal parsing that skips complex structures
        // Focus on essential document properties and main text
        throw new NotImplementedException("Minimal structure parsing not yet implemented");
    }
    
    private WordDocument ParseTextRunsOnly(Stream stream)
    {
        // Implement text-only extraction that bypasses complex formatting
        throw new NotImplementedException("Text-only parsing not yet implemented");
    }
}
```

### Phase 4: Enhanced Shell Application Error Handling (1 day)

#### 4.1 Comprehensive Error Handling in Shell Apps
**File: `Shell/doc2text/Program.cs`** (enhance)

```csharp
class Program
{
    private static readonly ILogger Logger = LoggerFactory.CreateLogger<Program>();
    
    static async Task<int> Main(string[] args)
    {
        try
        {
            var options = ParseCommandLineArguments(args);
            return await ConvertDocument(options);
        }
        catch (Exception ex)
        {
            return HandleGlobalError(ex);
        }
    }
    
    private static async Task<int> ConvertDocument(ConversionOptions options)
    {
        // Pre-flight validation
        var validationResult = DocumentValidator.ValidateCompoundFile(options.InputPath);
        if (!validationResult.IsValid)
        {
            Console.Error.WriteLine($"Input file validation failed:");
            foreach (var error in validationResult.Errors)
            {
                Console.Error.WriteLine($"  • {error}");
            }
            return ExitCodes.InvalidInput;
        }
        
        if (validationResult.HasWarnings)
        {
            Console.WriteLine("Warnings detected:");
            foreach (var warning in validationResult.Warnings)
            {
                Console.WriteLine($"  ⚠ {warning}");
            }
        }
        
        // Create recovery context
        var recoveryContext = new ErrorRecoveryContext(options.InputPath)
        {
            AllowPartialRecovery = options.AllowPartialRecovery,
            MaxRecoveryAttempts = options.MaxRecoveryAttempts
        };
        
        try
        {
            using var storage = new StructuredStorageReader(options.InputPath);
            
            // Try tolerant parsing first
            var parser = new TolerantWordDocumentParser(recoveryContext, Logger);
            var parseResult = parser.TryParse(storage.GetStream("WordDocument"));
            
            if (!parseResult.Success)
            {
                throw parseResult.Error ?? new ParsingException("Failed to parse document");
            }
            
            if (parseResult.WasRecovered)
            {
                Console.WriteLine($"Document was recovered: {parseResult.RecoveryDetails}");
            }
            
            // Convert to text
            using var writer = new StreamWriter(options.OutputPath);
            TextConverter.Convert(parseResult.Value, writer, options.ExtractionSettings);
            
            // Report recovery statistics if any
            var recoveryReport = recoveryContext.GenerateReport();
            if (recoveryReport.TotalAttempts > 0)
            {
                Console.WriteLine($"Recovery attempts: {recoveryReport.SuccessfulAttempts}/{recoveryReport.TotalAttempts}");
            }
            
            return ExitCodes.Success;
        }
        catch (MalformedFileException ex)
        {
            Console.Error.WriteLine($"Error: The file appears to be corrupted.");
            Console.Error.WriteLine($"Details: {ex.Message}");
            if (ex.FilePosition.HasValue)
            {
                Console.Error.WriteLine($"Error occurred at file position: {ex.FilePosition.Value}");
            }
            return ExitCodes.CorruptedFile;
        }
        catch (UnsupportedFileVersionException ex)
        {
            Console.Error.WriteLine($"Error: Unsupported file version.");
            Console.Error.WriteLine($"Detected version: {ex.DetectedVersion}");
            Console.Error.WriteLine("This tool supports Word 97-2003 (.doc) files.");
            return ExitCodes.UnsupportedVersion;
        }
        catch (InsufficientDataException ex)
        {
            Console.Error.WriteLine($"Error: Unexpected end of file.");
            Console.Error.WriteLine($"Required {ex.RequiredBytes} bytes, but only {ex.AvailableBytes} available.");
            return ExitCodes.TruncatedFile;
        }
        catch (RecursionLimitException ex)
        {
            Console.Error.WriteLine($"Error: Document structure too complex.");
            Console.Error.WriteLine($"Maximum nesting depth ({ex.MaxDepth}) exceeded.");
            return ExitCodes.ComplexityLimit;
        }
        catch (Exception ex)
        {
            Logger.LogError(ex, "Unexpected error during conversion");
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            
            if (options.Verbose)
            {
                Console.Error.WriteLine("Stack trace:");
                Console.Error.WriteLine(ex.ToString());
            }
            
            return ExitCodes.UnexpectedError;
        }
    }
    
    private static int HandleGlobalError(Exception ex)
    {
        Console.Error.WriteLine($"Fatal error: {ex.Message}");
        Logger.LogCritical(ex, "Unhandled exception in main");
        return ExitCodes.FatalError;
    }
}

public static class ExitCodes
{
    public const int Success = 0;
    public const int InvalidInput = 1;
    public const int CorruptedFile = 2;
    public const int UnsupportedVersion = 3;
    public const int TruncatedFile = 4;
    public const int ComplexityLimit = 5;
    public const int UnexpectedError = 10;
    public const int FatalError = 99;
}
```

### Phase 5: Testing and Validation Framework (1-2 days)

#### 5.1 Error Handling Test Suite
**File: `UnitTests/ErrorHandlingTests.cs`** (create)

```csharp
[TestFixture]
public class ErrorHandlingTests
{
    [Test]
    public void ShouldHandleEmptyFile()
    {
        var emptyFile = Path.GetTempFileName();
        try
        {
            File.WriteAllBytes(emptyFile, new byte[0]);
            
            var result = DocumentValidator.ValidateCompoundFile(emptyFile);
            
            Assert.That(result.IsValid, Is.False);
            Assert.That(result.Errors, Contains.Item("File too small: 0 bytes (minimum 512 required)"));
        }
        finally
        {
            File.Delete(emptyFile);
        }
    }
    
    [Test]
    public void ShouldHandleInvalidSignature()
    {
        var invalidFile = Path.GetTempFileName();
        try
        {
            var fakeData = new byte[1024];
            fakeData[0] = 0xFF; // Wrong signature
            File.WriteAllBytes(invalidFile, fakeData);
            
            var result = DocumentValidator.ValidateCompoundFile(invalidFile);
            
            Assert.That(result.IsValid, Is.False);
            Assert.That(result.Errors.Any(e => e.Contains("Invalid compound file signature")), Is.True);
        }
        finally
        {
            File.Delete(invalidFile);
        }
    }
    
    [Test]
    public void ShouldHandleTruncatedFile()
    {
        // Create a file with valid signature but truncated content
        var truncatedFile = Path.GetTempFileName();
        try
        {
            var data = new byte[256]; // Too small for valid compound file
            data[0] = 0xD0; data[1] = 0xCF; data[2] = 0x11; data[3] = 0xE0; // Valid signature start
            data[4] = 0xA1; data[5] = 0xB1; data[6] = 0x1A; data[7] = 0xE1;
            File.WriteAllBytes(truncatedFile, data);
            
            var result = DocumentValidator.ValidateCompoundFile(truncatedFile);
            
            Assert.That(result.IsValid, Is.False);
            Assert.That(result.Errors.Any(e => e.Contains("File too small")), Is.True);
        }
        finally
        {
            File.Delete(truncatedFile);
        }
    }
    
    [Test]
    public void SafeBinaryReaderShouldValidateBounds()
    {
        var buffer = new byte[10];
        var offset = 8;
        
        // Should throw when trying to read beyond buffer
        Assert.Throws<InsufficientDataException>(() => 
            SafeBinaryReader.ReadInt32(buffer, ref offset));
    }
    
    [Test]
    public void RecoveryContextShouldTrackAttempts()
    {
        var context = new ErrorRecoveryContext("test.doc")
        {
            MaxRecoveryAttempts = 2
        };
        
        context.LogRecoveryAttempt("Test1", new Exception("Error1"), false);
        Assert.That(context.CanAttemptRecovery(), Is.True);
        
        context.LogRecoveryAttempt("Test2", new Exception("Error2"), false);
        Assert.That(context.CanAttemptRecovery(), Is.False);
        
        var report = context.GenerateReport();
        Assert.That(report.TotalAttempts, Is.EqualTo(2));
        Assert.That(report.SuccessfulAttempts, Is.EqualTo(0));
    }
}
```

## Implementation Timeline
- **Day 1-2**: Exception hierarchy and validation framework
- **Day 3-4**: Safe binary reading and recovery context
- **Day 5-7**: Intelligent error recovery and tolerant parsing
- **Day 8**: Enhanced shell application error handling
- **Day 9-10**: Testing framework and validation

## Success Metrics
1. Zero unhandled exceptions in shell applications
2. Clear, actionable error messages for all failure scenarios
3. Successful partial recovery for at least 70% of corrupted files
4. Comprehensive error logging and diagnostics
5. Regression testing ensures no impact on valid file processing

## Testing Commands
```bash
# Build the solution
dotnet build b2xtranslator.sln

# Test comprehensive error handling
dotnet test UnitTests/UnitTests.csproj --filter "Category=ErrorHandling"

# Test with various error scenarios
echo "Invalid content" > test-invalid.doc
dotnet run --project Shell/doc2text/doc2text.csproj -- test-invalid.doc output.txt
# Should show: "Error: The file appears to be corrupted."

# Test with empty file
touch test-empty.doc
dotnet run --project Shell/doc2text/doc2text.csproj -- test-empty.doc output.txt
# Should show: "Input file validation failed: File too small"

# Test with truncated file
head -c 100 samples/simple.doc > test-truncated.doc
dotnet run --project Shell/doc2text/doc2text.csproj -- test-truncated.doc output.txt
# Should handle gracefully with specific error message

# Test error recovery with verbose output
dotnet run --project Shell/doc2text/doc2text.csproj -- test-corrupted.doc output.txt --verbose --allow-partial-recovery

# Integration tests with error scenarios
dotnet test IntegrationTests/IntegrationTests.csproj --filter "Category=ErrorRecovery"
```
