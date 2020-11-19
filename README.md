# VBAExportForGit
A VBA Macro to export all modules, classes and forms in a file, open a command prompt in the folder where they are exported, and set up a .gitignore there.

# Requirements:
"Trust access to the VBA project object model" must be checked in the Excel Trust Center. (Under "Macro Settings")
Microsoft Scripting Runtime and Microsoft Visual Basic for Applications Extensibility references must be enabled.

# How to use:
-Create a new ".xlsm" file, name it something convenient, and import ExportModule.bas
-In the same folder, create an "ignorelist.txt" file, and add to it the filenames/extenstions that you'd like to be included in your default .gitignore files. *(mine is just "\*.tmp" at the moment*
-Add the macro to Excel's Quick Access Toolbar (via "Customize the Quick Access Toolbar).
-Make Changes to your vba code.
-Run the macro.
-Run your desired git commands in the command prompt window that is opened.
-That's it!

# What it does:
1: Loops through all components in the active workbook
2: Creates a "source" folder in the same directory as the workbook file
3: Creates a subdirectory in the "source" directory based on the workbook's name
  *This is done in case you have multiple workbooks that are all part of the same project, so that you can keep them separate from eachother*
4: Copies ignorelist.txt to a .gitignore file inside the "source" folder
5: Opens a command prompt in the source folder

I hope this contributes to improving your VBA coding workflow!
