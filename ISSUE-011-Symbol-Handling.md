# Issue #011: Incorrect Symbol Handling (Partially fixed)

## Problem Description
Special symbols, such as those from the "Symbol" or "Wingdings" fonts, are not being handled correctly during text extraction.
This can result in missing characters, incorrect characters (e.g., showing a letter instead of the symbol), or garbled text in the output. The `README.md` mentions `Bug49908.doc` as a sample case.

## Reproduction Steps

### Basic Reproduction
1.  Use the sample file `Bug49908.doc`.
2.  Run the text extraction tool:
    ```bash
    dotnet run --project Shell/doc2text/doc2text.csproj -- samples/Bug49908.doc output.txt
    ```
3.  Examine `output.txt` and compare it to the original document to see where symbols are rendered incorrectly.

### Scenario 2 `symbol.doc`
See the image `ISSUE-011-Symbol-Handling.png`
The output should show the exact corresponding UTF-8 code of the symbol presented.


Expected result (it is the last working version result, this result is satisfactory)
```
!"#$%&'()∗±,−./0123456789:;<=>?@ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩ
ΞΨΖ[\]^_`αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}~¡¢≤¤∞¦§¨©ª←↑­®¯°±
²≥´μ∂·¸¹º→¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐ∑ÒÓÔ∏√×∝ÙÚ↔∠↕Þß◊áâã
		
ä∅∈∉∋∌∩∪⊂⊃⊆⊇⊥∴∵óôõö÷øùúûüýþÿ
``` 

### Scenario 3 `61586.doc`

The only issue here is the symbol: μ

```txt


TEST? 
111 μg.h/mL (AUC) and 15 μg/mL (Cmax).  
TEST? 
Greek muμ
(


```
