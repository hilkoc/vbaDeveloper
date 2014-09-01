VbaDeveloper
============

VbaDeveloper is an excel addin for easy version control of all your vba code.
It will automatically export all your classes and modules to plain text, whenever you save your vba project. In this way changes can easily be committed using git or svn or any other source control system.

VbaDeveloper can also import the code again into your excel workbook. This is particularly useful after reverting an earlier commit or after merging branches. When you open an excel workbook it will ask if you want to import the code for that project.

Building the addin
-----------------------

This repository does not contain the addin itself, only the files needed to build it.  In short it come downs to these steps:

 - Manually import the Build module into a new excel workbook.
 - Add the required vba references.
 - Save the workbook as an excel add-in.
 - Close it, then open it again and import the other modules.

Detailed instructions can be found in *src/vbaDeveloper.xlam/Build.bas*.
