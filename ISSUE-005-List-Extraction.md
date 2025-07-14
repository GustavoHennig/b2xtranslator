# Issue #005: List Extraction Issues - Missing Bullet Points and Numbering

## Current Status
**A detailed resolution plan is now available.**

## Problem Description
The b2xtranslator fails to properly extract list formatting during text conversion, resulting in plain text output that lacks bullet points, numbering, and proper indentation. This significantly impacts document readability and structure preservation.

## Severity
**MEDIUM** - Affects document structure and readability but doesn't prevent conversion

## Impact
- Loss of document structure and organization
- Poor readability of converted text documents
- Missing visual hierarchy from lists
- Inability to distinguish between regular paragraphs and list items
- Poor user experience when converting documents with extensive lists
- Information organization is lost in the conversion process

## Detailed Resolution Plan

### Phase 1: Enhance List Data Parsing

The foundation of fixing list extraction is to correctly parse the data structures that define lists in the Word binary format. This involves the List Table (`LST`) and the List Format Override Table (`LFO`).

1.  **Fully Parse `LSTF` and `LVL` Structures:**
    Ensure the `Doc/DocFileFormat/Data/ListTable.cs` and related classes can completely parse the `LSTF` (List Format) and `LVL` (Level) structures from the document's Table Stream. Key fields are:
    - `lsid`: The unique ID for the list.
    - `tplc`: A template code that defines the list's basic properties.
    - `rgistd`: An array that maps paragraph styles to list levels.
    - In the `LVL` structure: `nfc` (number format code), `ixch` (character index for bullets), and indentation properties (`dxaLeft`, `dxaIndent`).

2.  **Parse `PAPX` for List Properties:**
    In `Doc/DocFileFormat/Structures/ParagraphProperties.cs`, the `PAPX` structure contains two crucial fields that link a paragraph to a list:
    - `ilfo`: The index into the `LFO` table. This identifies which list override applies.
    - `ilvl`: The indentation level of the paragraph within the list (0-8).
    These fields must be reliably extracted for every paragraph.

### Phase 2: Augment the Internal Document Model

Once the data is parsed, it needs to be stored in an accessible way in the in-memory document model.

1.  **Create List Information Classes:**
    Introduce new classes to represent the parsed list data.

    ```csharp
    // In a new file, e.g., Doc/DocFileFormat/ListInfo.cs
    public class ListLevelInfo
    {
        public int NumberFormatCode { get; set; } // The nfc code
        public char BulletCharacter { get; set; } // The actual character to use
        public string NumberFormatString { get; set; } // e.g., "%1."
        public int IndentDxa { get; set; }
    }

    public class ListInfo
    {
        public int ListId { get; set; }
        public List<ListLevelInfo> Levels { get; set; } = new List<ListLevelInfo>();
    }
    ```

2.  **Store Parsed Lists in `WordDocument`:**
    In `WordDocument.cs`, add a property to store all the lists defined in the document.

    ```csharp
    // In Doc/DocFileFormat/WordDocument.cs
    public Dictionary<int, ListInfo> AllLists { get; private set; }
    ```
    This dictionary will be populated by the `ListTable` parser, mapping a `ListId` to its full definition.

3.  **Link Paragraphs to Lists:**
    Add a property to the paragraph representation to hold its specific list formatting.

    ```csharp
    // In the class representing a parsed paragraph
    public class Paragraph
    {
        // ... other properties
        public int ListId { get; set; } = -1; // Default to no list
        public int ListLevel { get; set; } = -1;
    }
    ```
    When parsing the `PAPX` for each paragraph, populate these two fields.

### Phase 3: Implement Text Mapping for Lists

With the list information now available in the model, the `TextMapping` can be updated to generate the correct output.

1.  **Create a List State Manager:**
    In `Text/TextMapping/TextMapping.cs`, create a mechanism to track the current number for each active list.

    ```csharp
    // In TextMapping.cs
    private Dictionary<int, int> _listCounters = new Dictionary<int, int>();
    ```

2.  **Update Paragraph Mapping Logic:**
    In the method that processes paragraphs (likely in `TextMapping.cs` or a `ParagraphMapping.cs`), modify the logic to be list-aware.

    ```csharp
    // In the paragraph processing method
    protected override void HandleParagraph(Paragraph p)
    {
        if (p.ListId != -1)
        {
            // This is a list item
            var list = _wordDocument.AllLists[p.ListId];
            var level = list.Levels[p.ListLevel];

            // 1. Calculate indentation
            var indent = new string(' ', p.ListLevel * 2);
            _writer.Write(indent);

            // 2. Get bullet or number
            string bullet;
            if (level.NumberFormatCode == 23) // Simple bullet
            {
                bullet = level.BulletCharacter + " ";
            }
            else // Numbered list
            {
                if (!_listCounters.ContainsKey(p.ListId))
                {
                    _listCounters[p.ListId] = 0;
                }
                _listCounters[p.ListId]++;
                bullet = string.Format(level.NumberFormatString, _listCounters[p.ListId]) + " ";
            }
            _writer.Write(bullet);
        }
        else
        {
            // Not a list item, so reset counters for any list that might have just ended.
            // (This logic needs to be robust to handle list restarts)
            _listCounters.Clear(); 
        }

        // 3. Write the actual paragraph text
        _writer.WriteLine(p.Text);
    }
    ```

## Files to Investigate
- `Text/TextMapping/TextMapping.cs`
- `Doc/DocFileFormat/Data/ListTable.cs`
- `Doc/DocFileFormat/Structures/ParagraphProperties.cs` (`PAPX` parsing)
- `Doc/DocFileFormat/Data/ListFormatOverride.cs` (`LFO` parsing)

## Success Criteria
1. **Bullet Preservation**: Simple bulleted lists are prefixed with a bullet character (e.g., `•`).
2. **Number Preservation**: Simple numbered lists are prefixed with the correct, incrementing number (e.g., `1.`, `2.`).
3. **Indentation**: Nested list items are indented with leading spaces.
4. **No Regressions**: Paragraphs that are not part of a list are rendered correctly without any extra formatting.
