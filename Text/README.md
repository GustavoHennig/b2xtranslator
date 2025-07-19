# GHSoftware.WordDocTextExtractor

**GHSoftware.WordDocTextExtractor** is a .NET library for **extracting text from legacy Microsoft Word `.doc` files (Word 97-2003, Word 95 and Word 6.0)**.

This project is based on and extends the functionality of the original `b2xtranslator` project, with a focus on **robust plain text extraction**. It handles complex document structures and edge cases, producing clean, readable text output from binary `.doc` files.

---

## **Key Features**

* **DOC to Plain Text Conversion**
  Converts `.doc` files (Word 97, 6.0) into plain text while preserving document flow as much as possible.

* **Enhanced Compatibility**
  Handles tables, headers/footers, embedded objects, and other Word-specific structures.

* **Active Maintenance**
  This library is actively maintained and modernized for .NET 6+.

---

## **Installation**

```
dotnet add package GHSoftware.WordDocTextExtractor
```

---

## **Usage Example**

```csharp
using b2xtranslator.txt;

string path = "legacy-document.doc";
string extractedText = DocTextExtractor.ExtractTextFromFile(docPath);

Console.WriteLine(extractedText);
```

---

## **More Information**

For advanced usage, CLI tools, or additional formats, see the main [`b2xtranslator`](https://github.com/GustavoHennig/b2xtranslator) project.

---

## **License**

This project is open source and distributed under the same license as the original `b2xtranslator`.
See [LICENSE](https://github.com/GustavoHennig/b2xtranslator/blob/master/LICENSE).

---

## **Credits**

Originally developed by **DIaLOGIKa** (2008-2009) and **Evolution Recruitment Solutions** (2017).
Maintained and extended by **Gustavo Hennig (2025)**.

