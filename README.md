# b2xtranslator

.NET Core library to convert Microsoft Office binary files to various formats. This fork focuses **exclusively on Word (.doc) files** and **plain text extraction** from legacy Microsoft Word documents (Word 97-2003, Word 95, and Word 6.0).

You can also use the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) to manipulate OpenXML files.

Forked from a [.NET 2 Mono implementation](https://sourceforge.net/projects/b2xtranslator/) under the BSD license.

## Key Features

### Text Extraction (Primary Focus)
- **DOC to Plain Text Conversion**: Robust extraction from Word 97-2003, Word 95, and Word 6.0 formats
- **Enhanced Compatibility**: Handles tables, headers/footers, embedded objects, and complex document structures
- **Clean Output**: Produces readable text while preserving document flow
- **Edge Case Handling**: Robust processing of corrupted or non-standard .doc files

### Legacy Format Support (Not Maintained)
- PowerPoint (.ppt) to PPTX conversion
- Excel (.xls) to XLSX conversion
- Word (.doc) to DOCX conversion

*Note: This fork maintains these legacy features but does not actively enhance them.*

## Roadmap

### Planned Enhancements
- Enhanced formatting support for lists (numbers, bullet points, indents) and tables
- Configurable text extraction options (--no-headers-footers, --no-textboxes, --no-comments, --no-bullets)
- Performance optimizations for large document processing
- Additional error handling and recovery mechanisms

This project is inspired by and informed by several existing open-source implementations of the Word Binary Format:

| Name            | Language | Description                                                        | Link                                                                  |
|-----------------|----------|--------------------------------------------------------------------|-----------------------------------------------------------------------|
| **wvWare**      | C        | Original GPL Word97 `.doc` text extractor                          | [SourceForge](https://sourceforge.net/projects/wvware/)               |
| **OnlyOffice**  | C++      | Proprietary editor with open-source core, includes DOC parsing     | [GitHub](https://github.com/ONLYOFFICE/core/tree/master/MsBinaryFile) |
| **Antiword**    | C        | Lightweight Word `.doc` to text/postscript converter               | [GitHub Mirror](https://github.com/grobian/antiword)                  |
| **Apache POI**  | Java     | Java API for Microsoft Documents, includes Word97 support via HWPF | [Apache POI - HWPF](https://poi.apache.org/hwpf/index.html)           |
| **LibreOffice** | C++      | Full office suite with robust support for legacy DOC files         | [GitHub](https://github.com/LibreOffice/core)                         |
| **Catdoc**      | C        | Lightweight Word `.doc` to text converter                          | [GitHub Mirror](https://github.com/petewarden/catdoc)                 |
| **DocToText**   | C++      | Lightweight any document file to text converter                    | [GitHub](https://github.com/tokgolich/doctotext)                      |

## References

* [Microsoft Office binary files documentation](https://msdn.microsoft.com/en-us/library/cc313105.aspx)
* [Open XML Standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
* [Microsoft article on this implementation](https://blogs.msdn.microsoft.com/interoperability/2009/05/11/binary-to-open-xml-b2x-translator-interoperability-for-the-office-binary-file-formats/)
* [.NET 2 Mono implementation architecture](http://b2xtranslator.sourceforge.net/architecture.html)

All code retained from that version ©2009 DI<sup><u>a</u></sup>LOGIK<sup><u>a</u></sup> http://www.dialogika.de/  
.NET core port work and move to `System.IO.Compression` ©2017 Evolution https://www.evolutionjobs.com/
