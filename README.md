# TLDR Overview
This project contains 3 main subs which are designed to make using git with vba significantly less painless than it otherwise is.

* VBAExportForGit -> Exports all modules, classes and forms into a 'source' folder, placed in the same directory as the active workbook, and then opens a cmd prompt there (so you can git add, commit, push, etc.). 
* OpenCMDOnly -> Just opens a command prompt in the 'source' folder, if it exists (so you can git pull, branch, checkout, etc.).
* VBAImportForGit -> Checks whether there is a 'source' folder in the active workbook's folder, and if there is, imports modules, classes, and forms there. (If components with the same names exist, user will be asked whether they want to replace them.)

## Requirements:
* "Trust access to the VBA project object model" must be checked in the Excel Trust Center. *(Under "Macro Settings")*
* Microsoft Scripting Runtime and Microsoft Visual Basic for Applications Extensibility references must be enabled.

## VBAExportForGit
A VBA Module to export all modules, classes and forms in a file, open a command prompt in the folder where they are exported, and set up a .gitignore there.

### How to use:
* Create a new ".xlsm" file, name it something convenient, and import .bas and .cls files from this repo
* In the same folder, create an "ignorelist.txt" file, and add to it the filenames/extenstions that you'd like to be included in your default .gitignore files. *(mine is just "\*.tmp" at the moment)*
* Add the macro to Excel's Quick Access Toolbar (via "Customize the Quick Access Toolbar).
* Make Changes to your vba code.
* Run the macro.
* Run your desired git commands in the command prompt window that is opened.

## VBAImportForGit
A VBA Module to import all modules, classes and forms from a 'source' folder.

### How to use:
* Create a new ".xlsm" file, name it something convenient, and import .bas and .cls files from this repo
* In the same folder, create an "ignorelist.txt" file, and add to it the filenames/extenstions that you'd like to be included in your default .gitignore files. *(mine is just "\*.tmp" at the moment*
* Add the macro to Excel's Quick Access Toolbar (via "Customize the Quick Access Toolbar).
* Save/close the workbook
* Create a new .xlsm file, naming it *exactly the same* as the file that you want to import from. (so, if you were importing from this repo, you would name it "VBAExportForGit.xlsm")
* Clone/Pull your repo into the same folder as your .xlsm file
* Rename the repo folder 'source'.
* Run the macro from the excel file you want to import to.
* All vba objects from the source folder will be imported.

## OpenCMDOnly
A VBA Module to open a CMD window in the 'source' folder, or the current excel file's folder, if no 'source' folder exists there.

### How to use:
* Create a new ".xlsm" file, name it something convenient, and import .bas and .cls files from this repo
* Add the macro to Excel's Quick Access Toolbar (via "Customize the Quick Access Toolbar).
* Run the macro from the file that you'd like to do verson control for. A CMD window will be opened in the 'source' folder if it exists in that location, or the option will be given to open in the file's location if no 'source' folder is found.

## Basic Workflow for Exporting
* Edit an excel file, make changes to its vba code.
* Run VBAExportForGitMacro
* CMD window will open
* Run git commands in open window (add/commit/push)
* Exit CMD window
* Done!

## Basic Workflow for Importing
* Run OpenCMDOnly macro from file you want to import to
* (Create 'source' folder if it doesn't exist)
* Run git commands in source folder (clone/pull)
* Exit CMD window
* Run VBAImportForGit from file you want to import to
* All objects in 'source' folder will be imported (You will be prompted to overwrite if they already exist.)
* Done!

## I hope this contributes to improving your VBA coding workflow!
