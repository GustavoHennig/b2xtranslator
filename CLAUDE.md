# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

b2xtranslator is a .NET Core library that converts Microsoft Office binary files (doc, xls, ppt) to Open XML formats (docx, xlsx, pptx). It's a .NET 8 port of an original .NET 2 Mono implementation, designed to handle legacy Office file formats.

### Fork Goals and Enhancements

This fork aims to significantly expand the functionality and robustness of the original implementation:

**Enhanced Text Extraction:**
- Improved plain text extraction capabilities
- Better handling of complex document structures and formatting
- Support for extracting text from embedded objects and headers/footers

**Extended File Format Support:**
- Support for older Word file formats (Word 6.0, Word 95, etc.)
- Better compatibility with various binary Office file versions
- Enhanced handling of corrupted or non-standard files

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
# Convert DOC to DOCX
dotnet run --project Shell/doc2x/doc2x.csproj -- input.doc output.docx

# Convert PPT to PPTX  
dotnet run --project Shell/ppt2x/ppt2x.csproj -- input.ppt output.pptx

# Convert XLS to XLSX
dotnet run --project Shell/xls2x/xls2x.csproj -- input.xls output.xlsx

# Convert DOC to text
dotnet run --project Shell/doc2text/doc2text.csproj -- input.doc output.txt
```

### Testing with Sample Files
The repository includes extensive sample files in `samples/` and `samples-local/` directories with expected output files for validation:
```bash
# Sample files follow pattern: filename.doc, filename.expected.txt, filename.actual.txt
# Integration tests automatically compare actual vs expected output

# Test doc2text conversion with sample files
dotnet run --project Shell/doc2text/doc2text.csproj -- samples-local/simple-table.doc output.txt
dotnet run --project Shell/doc2text/doc2text.csproj -- samples-local/bug65255.doc output.txt

# Sample file locations:
# - samples/ - Main sample files directory
# - samples-local/ - Additional test files with expected outputs
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
- `Doc/` - Word (.doc) file format parser
- `Ppt/` - PowerPoint (.ppt) file format parser  
- `Xls/` - Excel (.xls) file format parser

**Mapping/Conversion Layers:**
- `Docx/` - DOC to DOCX conversion mappings
- `Text/` - DOC to plain text conversion mappings
- Each format has its own `*Mapping/` subdirectory with conversion logic

**Command-Line Tools:**
- `Shell/doc2x/` - DOC to DOCX converter
- `Shell/ppt2x/` - PPT to PPTX converter
- `Shell/xls2x/` - XLS to XLSX converter
- `Shell/doc2text/` - DOC to text converter

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
- Word 97-2003 (.doc files)
- PowerPoint 97-2003 (.ppt files) 
- Excel 97-2003 (.xls files)
- Handles complex features: macros, charts, embedded objects, formatting

**Output Formats:**
- OpenXML: .docx, .pptx, .xlsx with proper document type detection
- Plain text extraction for Word documents
- Automatic file extension correction based on document type (e.g., macro-enabled docs → .docm)

### Key Components

**Document Parsing:**
- `WordDocument`, `PowerpointDocument`, `XlsDocument` - Main document classes
- `BiffRecord` hierarchy for Excel binary records
- `Record` hierarchy for PowerPoint binary records
- File Information Block (FIB) parsing for Word documents

**OpenXML Generation:**
- `OpenXmlPackage` and format-specific implementations
- `OpenXmlWriter` for efficient XML generation
- Part-based architecture matching OpenXML specification

## Development Guidelines

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
When extending conversion capabilities:
1. Add parsing logic to appropriate format-specific project (Doc/, Ppt/, Xls/)
2. Create or extend mapping classes in corresponding *Mapping/ directories
3. Add test cases with sample files and expected outputs
4. Update the main converter classes to handle new features

### Code Organization
- Format parsers are read-only - they parse binary data into object models
- Mapping classes handle all output generation and OpenXML specifics
- Shared utilities in Common/ avoid duplication across formats
- Shell applications are thin wrappers around core conversion logic
