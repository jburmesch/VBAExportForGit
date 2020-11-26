Attribute VB_Name = "MainSubsModule"
Option Explicit


Const EXPORT_FOLDER_NAME As String = "source"
    

'Exports all components from active workbook
Sub VBAExportForGit()
    Dim wb As Workbook
    Dim project As VBProject
    Dim projectFolderPath As String
    Dim workbookName As String
    
    Set wb = ActiveWorkbook
    Set project = wb.VBProject
    
    workbookName = Split(wb.Name, ".")(0)
    projectFolderPath = wb.path & "\" & EXPORT_FOLDER_NAME
    
    Call ExportComponents(project.VBComponents, _
                                        projectFolderPath, workbookName)
    Call CopyGitIgnore(projectFolderPath)
    Call OpenCommandPrompt(projectFolderPath)
    ThisWorkbook.Close
    
End Sub


'Imports all components in active workbook's source folder if it exists.
'Otherwise, opens file dialogue to select components to import.
Sub VBAImportForGit()
    Dim importFiles As Variant
    Dim fileTypes As Variant
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    fileTypes = Array(".bas", ".cls", ".frm")
    
    importFiles = GetImportFiles(wb, fileTypes)
    Call Import(importFiles, wb)
    ThisWorkbook.Close
    
End Sub


'Open a command line in the source folder
Sub OpenCMDOnly()
    Dim wb As Workbook
    Dim projectFolderPath As String
    
    Set wb = ActiveWorkbook
    projectFolderPath = wb.path & "\" & EXPORT_FOLDER_NAME
    Call OpenCommandPrompt(projectFolderPath)
    
End Sub
