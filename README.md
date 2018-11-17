# DHCHeroCalculator

Required software
This calculator requires a program that can open an .xlsm file. To make full use of the calculator's automated features,
this would be a copy of Microsoft Excel, but the calculator can also be easily modified to work without access to VBA.

To use this calculator without VBA:
1. Sort both the Hero Capabilities and Hero Contributions to Team sheets by Color Ascending, Heroes Ascending
2. Ensure that both have the same heroes listed in the same order
3. Use the example formulas in the Example Formulas file to replace the automation-filled values with the appropriate formulas

Frequent issues:
* The object libraries used in the code were the ones bundled with Microsoft Office 2010. These libraries will not exactly
match the libraries of other users. To resolve this issue, push Alt-F11 in Microsoft Excel to open the VBE Editor, Tools menu,
References, uncheck any libraries prefixed with "MISSING: ", and enable the appropriate matching libraries of the same name
but different version in the list of available libraries.
