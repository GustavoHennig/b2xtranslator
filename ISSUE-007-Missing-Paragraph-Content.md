# Issue #007: Missing Paragraph Content Issues



## Problem Description
The b2xtranslator fails to extract some paragraphs from certain Word documents.

## Scenarios, affected files:

- 53446.doc
Missing fragment
```txt
 satisfactory to Lakeland and ENA, the Contract Price hereunder with respect to the discount provided to Lakeland herein, would be adjusted downward to reflect the reduction in the winter quantities delivered to Lakeland via FGT and the increase in the summer quantities delivered to Lakeland via FGT. 
```

- "Bug50936_1.doc"
```txt
BAR also performed gas cap pressure testing as part of the visual and functional test during the Roadside program.  During this portion of the test, BAR personnel consulted look-up charts to determine if an adapter was available to test the gas cap.  (Because adapters are not available for all vehicle models, not all gas caps are tested.)  This is the same procedure used in Smog Check stations.  If the vehicle was subject to testing, the fuel cap was removed from the vehicle and attached to a portable fuel cap testing unit.  The cap is subjected to a pressure of 30 inches of water.  The cap must hold this pressure with a leak rate of no more than 60 cubic centimeters per minute.  In addition to not being on the look-up chart, gas caps were not tested if the fuel cap tester did not pass the calibration test (in this situation, the test team was not able to test caps until the equipment was repaired).  If the vehicle had an incorrect fuel cap or no fuel cap at all, the vehicle failed the gas cap test.
```

- Bug51944.doc (extracting thing that does not exist, missing things)



- EL_TechnicalTemplateHandling.doc

This paragraph is missing from the file: `\samples\EL_TechnicalTemplateHandling.doc`

```
Updating the value of any property within this dialog will be reflected in the document after closing the dialogue. This is important since all the properties displayed in the document are updated, independently of their number of occurrences in the document. The Properties Dialog can thus be called any time you need to update the document properties. For all properties, the last setting will be memorized by the dialog. 
NOTE : the options present in this dialog rely on specific bookmarks and fields inside the document in order to be able to add the requested information at the defined spot. If these bookmarks and/or fields are removed or modified, it is possible that the properties dialog will not be able to achieve the desired result. 
Properties Dialog: Front Information Table
```
