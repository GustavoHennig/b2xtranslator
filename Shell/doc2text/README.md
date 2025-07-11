# doc2text

This tool converts a Word document (.doc) to a plain text file (.txt).

## Usage

To run the conversion, use the following command from the root of the repository:

```bash
dotnet run --project Shell/doc2text/doc2text.csproj -- <input.doc> <output.txt>
```

### Arguments

-   `<input.doc>`: The path to the input Word document file.
-   `<output.txt>`: The path where the output text file will be created.

### Example

```bash
dotnet run --project Shell/doc2text/doc2text.csproj -- samples/sample1.doc output.txt
```
