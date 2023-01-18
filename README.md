# Orange Label Creator
Dynamic document for creating one or more 'Orange Labels' to be used when commissioning hardware.



### Summary
The Orange Label Creator allows the quick and easy creation of multiple 'Orange Labels' while mitigating administrative errors.



### Compatibility:
Microsoft Office 2013 and newer. *(Untested, but might work with older versions)*



### Required Files
The Orange Label Creator does not currently require any additional files to be functional.



### Developer Mode
To improve versatility, I have implemented a **Developer Mode** into this Macro Enabled Excel document.

Developer Mode allows a user to safely make certain edits to the document such as changing the options in a dropdown list, the picture associated with a worksheet, the password used to lock and protect the document, and more. To enter Developer Mode, you must start by opening the document, enabling Macros, and clicking on any unlocked/editable cell. Once you have done that simply press *Ctrl*+*Shift*+*Alt*+*D*.

If done correctly, a window will appear informing you that you are about to enter Developer Mode. While in this mode, you will be able to safely make changed to certain aspects of the document. If further or more in-depth changes need to be made, please contact me. See my bio for the appropriate contact information.



### Making edits beyond what can be made in Developer Mode

Edits beyond what is possible while in Developer Mode can be made by unlocking the document entirely with the password given in the files Developer Mode sheet, or by entering Developer Mode, and while on the **DevMode** sheet pressing *Ctrl*+*Shift*+*~*.
**I STRONGLY URGE YOU TO USE CAUTION IF YOU MAKE THESE KINDS OF CHANGES. ONLY DO SO IF YOU KNOW EXACTLY WHAT YOU ARE DOING.**

It is important to note that changes beyond what can be made while in Developer Mode are very likely to render the file broken as there are hard coded references to cells, sheets, and other aspects of the document that can be broken if changes are made and the VBA code is not updated to reflect those changes. That being said, if you understand how to read and write VBA macros, and you are proficient in doing so, then you are more than welcome to attempt to do this on your own, but at your own risk of breaking the file (and possibly other associated files) by doing so.
