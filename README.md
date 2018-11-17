# DHCHeroCalculator

I'm getting warnings about macros. Is this calculator trying to execute malicious code on my computer?:
No, but malicious macros can be attached to Microsoft Office documents, so I'm glad you're suspicious. To alleviate concerns, all of the code included in the calculator is saved here separately in project files for your review. In addition, you can simply choose not to enable macros and review the code directly in the program by pushing Alt-F11 to bring up the VBE Editor. Or, finally, you can simply convert the file to .xlsx format, stripping it of the code, and use the Example Formulas to update the calculator by hand.  

Required software:
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
