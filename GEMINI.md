# GEMINI.md


This file provides guidance to Gemini when working with code in this repository.

## Overview

b2xtranslator is a .NET Core library that converts Microsoft Office binary files to various formats. This fork focuses **exclusively on Word (.doc) files** and **plain text extraction**. The original library supported Excel (.xls) and PowerPoint (.ppt) formats, but these are now considered legacy features and will not be maintained or enhanced in this fork.

### Fork Goals and Enhancements

This fork aims to significantly expand the functionality and robustness for **Word document processing only**:

**Enhanced Text Extraction (PRIMARY FOCUS):**
- Improved plain text extraction capabilities from .doc files
- Better handling of complex document structures and formatting
- Support for extracting text from embedded objects and headers/footers
- Robust conversion from .doc to .txt format

**Extended Word File Format Support:**
- Support for older Word file formats (Word 6.0, Word 95, etc.)
- Better compatibility with various binary Word file versions
- Enhanced handling of corrupted or non-standard .doc files

**Critical Bug Fixes:**
- **Infinite Loop Issues:** Address parsing loops that cause applications to hang
- **Performance Problems:** Fix extremely slow conversion scenarios 
- **Memory Leaks:** Resolve memory management issues in large file processing
- **Parsing Errors:** Fix crashes and exceptions during document parsing

**Reliability Improvements:**
- Robust error handling and recovery mechanisms
- Better validation of input files before processing
- Graceful degradation when encountering unsupported features

## Common Development Commands

### Building the Solution
```bash
dotnet build b2xtranslator.sln
```

### Running Tests
```bash
# Unit tests (NUnit-based, older project)
dotnet test UnitTests/UnitTests.csproj

# Integration tests (xUnit-based, newer project)
dotnet test IntegrationTests/IntegrationTests.csproj
```

### Running Individual Applications
```bash
# Convert DOC to text (PRIMARY FOCUS)
dotnet run --project Shell/doc2text/doc2text.csproj -- input.doc output.txt

# Convert DOC to DOCX (LEGACY - NOT MAINTAINED)
dotnet run --project Shell/doc2x/doc2x.csproj -- input.doc output.docx

# Convert PPT to PPTX (LEGACY - NOT MAINTAINED)
dotnet run --project Shell/ppt2x/ppt2x.csproj -- input.ppt output.pptx

# Convert XLS to XLSX (LEGACY - NOT MAINTAINED)
dotnet run --project Shell/xls2x/xls2x.csproj -- input.xls output.xlsx
```

### Testing with Sample Files
The repository includes extensive sample files in the `samples/` directory with expected output files for validation:
```bash
# Sample files follow pattern: filename.doc, filename.expected.txt
# The correct way to test is to export to text and compare with expected output

# Test doc2text conversion with sample files from samples/ directory
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/simple-table.doc output.txt
# Then compare output.txt with samples/simple-table.expected.txt

dotnet run --project Shell/doc2text/doc2text.csproj -- samples/bug65255.doc output.txt
# Then compare output.txt with samples/bug65255.expected.txt

# Sample file locations:
# - samples/ - Main sample files directory with .doc files and corresponding .expected.txt files
# Files include various document types: tables, headers/footers, images, complex formatting
```

## Architecture Overview

### Project Structure
The solution follows a layered architecture with clear separation of concerns:

**Core Libraries:**
- `Common.Abstractions/` - Core interfaces and abstractions
- `Common.CompoundFileBinary/` - Structured storage reader for binary Office files
- `Common/` - Shared utilities, OpenXML library, and common components

**Format-Specific Parsers:**
- `Doc/` - Word (.doc) file format parser (PRIMARY FOCUS)
- `Ppt/` - PowerPoint (.ppt) file format parser (LEGACY - NOT MAINTAINED)
- `Xls/` - Excel (.xls) file format parser (LEGACY - NOT MAINTAINED)

**Mapping/Conversion Layers:**
- `Text/` - DOC to plain text conversion mappings (PRIMARY FOCUS)
- `Docx/` - DOC to DOCX conversion mappings (LEGACY - NOT MAINTAINED)
- Each format has its own `*Mapping/` subdirectory with conversion logic

**Command-Line Tools:**
- `Shell/doc2text/` - DOC to text converter (PRIMARY FOCUS)
- `Shell/doc2x/` - DOC to DOCX converter (LEGACY - NOT MAINTAINED)
- `Shell/ppt2x/` - PPT to PPTX converter (LEGACY - NOT MAINTAINED)
- `Shell/xls2x/` - XLS to XLSX converter (LEGACY - NOT MAINTAINED)

### Key Architectural Patterns

**Visitor Pattern:** Used extensively for traversing and converting document structures. Core interfaces include:
- `IVisitable` - Elements that can be visited during conversion
- `IMapping` - Conversion mappings that process visitables
- `AbstractOpenXmlMapping` - Base class for OpenXML output mappings

**Structured Storage:** Binary Office files use Microsoft's Compound Document format:
- `StructuredStorageReader` - Reads the compound document structure
- Format-specific classes parse the binary streams within

**Mapping Architecture:** Two-phase conversion process:
1. **Parse Phase:** Binary format → Internal object model (e.g., `WordDocument`, `PowerpointDocument`)
2. **Convert Phase:** Internal model → OpenXML via mapping classes

**Context Pattern:** Conversion context objects carry shared state:
- `ConversionContext` - Shared conversion state and utilities
- `ExcelContext` - Excel-specific conversion context

### File Format Handling

**Binary Format Support:**
- Word 97-2003 (.doc files) - PRIMARY FOCUS
- PowerPoint 97-2003 (.ppt files) - LEGACY, NOT MAINTAINED
- Excel 97-2003 (.xls files) - LEGACY, NOT MAINTAINED
- Handles complex Word features: macros, charts, embedded objects, formatting

**Output Formats:**
- Plain text extraction for Word documents (.txt) - PRIMARY FOCUS
- OpenXML: .docx, .pptx, .xlsx with proper document type detection - LEGACY, NOT MAINTAINED
- Automatic file extension correction based on document type (e.g., macro-enabled docs → .docm) - LEGACY, NOT MAINTAINED

### Key Components

**Document Parsing:**
- `WordDocument` - Main document class (PRIMARY FOCUS)
- `PowerpointDocument`, `XlsDocument` - Main document classes (LEGACY - NOT MAINTAINED)
- `BiffRecord` hierarchy for Excel binary records (LEGACY - NOT MAINTAINED)
- `Record` hierarchy for PowerPoint binary records (LEGACY - NOT MAINTAINED)
- File Information Block (FIB) parsing for Word documents (PRIMARY FOCUS)

**OpenXML Generation:**
- `OpenXmlPackage` and format-specific implementations
- `OpenXmlWriter` for efficient XML generation
- Part-based architecture matching OpenXML specification

## Development Guidelines

**IMPORTANT:** Do not run any write GIT operations. You are only authorized to perform READ operations with GIT. Any attempt to perform write operations (such as push, commit, merge, rebase, etc.) is strictly prohibited.


### Testing Strategy
- Unit tests cover core parsing and conversion logic
- Integration tests use sample files with expected outputs
- Test files include edge cases, corrupted files, and various Office versions
- Regression testing ensures sample file compatibility

### Error Handling
- Format-specific exceptions for different error conditions
- Graceful handling of unsupported features or corrupted files
- Detailed logging via `TraceLogger` for debugging conversion issues

### Performance Considerations
- Stream-based parsing for large files
- Lazy loading of document sections
- Memory-efficient XML writing for large OpenXML documents

### Critical Issues to Address

**Known Performance Problems:**
- Infinite loops in certain document parsing scenarios (priority: critical)
- Extremely slow conversion times for specific file types
- Memory consumption issues with large or complex documents
- Threading and async operation bottlenecks

**Parsing Reliability Issues:**
- Crashes when encountering malformed binary data
- Incomplete handling of older Word file format versions
- Missing error boundaries in recursive parsing operations
- Insufficient validation of file structure before processing

**Debugging and Monitoring:**
- Add performance profiling hooks for identifying bottlenecks
- Implement timeout mechanisms for long-running operations
- Enhanced logging for tracking parsing progress and identifying stuck operations
- Memory usage monitoring and cleanup verification

### Adding New Features
When extending Word document text extraction capabilities:
1. Add parsing logic to the Doc/ project (Word document parsing only)
2. Create or extend mapping classes in the Text/ directory for text extraction
3. Add test cases with sample .doc files and corresponding .expected.txt files in the samples/ directory
4. Update the doc2text converter to handle new features

**Note:** This fork does not accept enhancements to PowerPoint (.ppt) or Excel (.xls) functionality.

### Code Organization
- Format parsers are read-only - they parse binary data into object models
- Mapping classes handle all output generation and OpenXML specifics
- Shared utilities in Common/ avoid duplication across formats
- Shell applications are thin wrappers around core conversion logic
