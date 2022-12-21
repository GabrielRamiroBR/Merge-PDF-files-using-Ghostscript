# Merge-PDF-files-using-Ghostscript
This repository contains a code in VBScript that merge pdf files in a folder using GhostScript API

Please follow the steps in https://ghostscript.readthedocs.io/en/latest/Install.html to install the ghostScript API in your operating system.

Also remember to change the conde line #9 *commandLine = "**gswin64c.exe** -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile=Merge.pdf "* 

Replacing **gswin64c.exe** for the right command depending on your operational system:

System | Invocation name|
--- | --- |
Unix | gs |
VMS | gs |
MS Windows 95 and later | gswin32.exe , gswin32c.exe, gswin64.exe , gswin64c.exe
OS2 | gsos2

