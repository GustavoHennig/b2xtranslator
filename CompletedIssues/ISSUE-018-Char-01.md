The character 0x001 is being returned in the text output, it is common between links.
Instead of only filtering, can you identify why?

If not fiter it out.

Example: `samples\Fuzzed.doc`, next char after: `http://www.nasa.gov/multimedia/imagegallery/index.html ` in the extracted text.

---

### Analysis and Solution

**Problem:**
The character `0x01` (a non-printable control character) is incorrectly included in the final text output, typically appearing at the end of a URL extracted from a hyperlink.

**Root Cause:**
The issue originates from the parsing of hyperlink field codes within `Text/TextWriter.cs`. The `0x01` character is present in the raw field data of some `.doc` files, immediately following the URL. The existing URL extraction logic uses simple string replacements, which are not precise enough to exclude this trailing control character. As a result, the `0x01` is treated as part of the URL and written to the output.

The specific faulty code is in the `WriteEndElement` method of the `TextWriter` class:
```csharp
// in Text/TextWriter.cs

// ...
else if ("instrText".Equals(element.LocalName))
{
    string content = element.Content.ToString();
    string trimmedContent = content.Trim();

    if (trimmedContent.StartsWith("HYPERLINK "))
    {
        string url;
        if (trimmedContent.StartsWith("HYPERLINK ""))
        {
            // This line is the cause of the issue. 
            // It includes the trailing 0x01 character as part of the URL.
            url = trimmedContent.Replace("HYPERLINK "", "").Replace(""", "").Trim();
        }
        else
        {
            url = trimmedContent.Replace("HYPERLINK ", "").Trim();
        }
        _pendingHyperlinkUrl = url;
        //...
    }
//...
```

**Proposed Fix:**
To resolve this, the URL parsing logic should be improved to be more robust. Instead of relying on broad string replacements, it should specifically extract only the content within the quotes of the `HYPERLINK` field code. Using a regular expression is the recommended approach.

**Example of a corrected implementation:**
```csharp
// In Text/TextWriter.cs, inside the WriteEndElement method

// ...
else if ("instrText".Equals(element.LocalName))
{
    // ...
    if (trimmedContent.StartsWith("HYPERLINK ""))
    {
        // Use a regular expression to find the first quoted string, which is the URL.
        // This ensures that any characters outside the quotes are ignored.
        var match = System.Text.RegularExpressions.Regex.Match(trimmedContent, @"""([^""]+)""");
        if (match.Success)
        {
            url = match.Groups[1].Value;
        }
        else
        {
            // Fallback to the old logic if the regex fails for some reason.
            url = trimmedContent.Replace("HYPERLINK "", "").Replace(""", "").Trim();
        }
    }
    // ...
}
```
This change ensures that only the valid URL is extracted, effectively discarding extraneous characters like `0x01` and preventing them from appearing in the final text output.