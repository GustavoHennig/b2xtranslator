# Binary(doc,xls,ppt) to OpenXMLTranslator

.NET Core library to convert Microsoft Office binary files (`doc`, `xls` and `ppt`) to Open XML (`docx`, `xlsx` and `pptx`).
You can use the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) to manipulate those.

Forked from a [.NET 2 Mono implementation](https://sourceforge.net/projects/b2xtranslator/) under the BSD license.

## Fork Goals and Roadmap

This fork aims to significantly expand the capabilities of the original implementation. The following sections outline planned enhancements and ongoing development goals:

### Roadmap
- **Improve Text Extraction**: Better plain text extraction with support for complex document structures
- **Extend Format Support**: Support for older Word file formats (Word 6.0, Word 95, etc.)
- **Improve Error Handling**: Graceful handling of corrupted or non-standard files
- Make it handle lists (numbers, bullet points, indents)
- Handle TextBox content: "news-example.doc"
- Symbol handling: "Bug49908.doc"

### 🔧 Pending Issues
- Address parsing loops that cause applications to hang
- **Performance Optimization**: Fix extremely slow conversion scenarios
- **Memory Management**: Resolve memory leaks in large file processing
- **Parsing Stability**: Fix crashes and exceptions during document parsing
- Lists are not extracted with proper bullet symbols during text conversion
- Internal markers are being written to the output: HYPERLINK, MERGEFORMAT, DOCPROPERTY, PAGEREF... See EL_TechnicalTemplateHandling.doc, ProblemExtracting.doc, Bug51686.doc
- testPictures.doc: testPictures.doc SHAPE  \* MERGEFORMAT
- Revisit: samplehtmlfieldandlist.doc
- fastsavedmix.doc - 1 line is missing
- Entire paragraphs are missing: "Bug50936_1.doc", "ESTAT Article comparing RU-LFS-22 12 05_EN.doc", "Bug47958.doc", "Bug53380_3.doc"
- Alternative space chars need to be handled: "Bug47742.doc" (space != 20)
- URL not being shown: "pad.doc"
 
### 📊 Reliability Improvements (Planned)
- Enhanced validation of input files before processing
- Better error recovery mechanisms
- Comprehensive logging for debugging conversion issues
- Support for timeout mechanisms in long-running operations 

## Reference Projects and Implementations

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
