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


Sub VBAImportForGit()
    Dim importFiles() As String
    
    importFiles = GetImportFiles
    
End Sub


'Open a command line in the source folder
Sub OpenCMDOnly()
    Dim wb As Workbook
    Dim projectFolderPath As String
    
    Set wb = ActiveWorkbook
    projectFolderPath = wb.path & "\" & EXPORT_FOLDER_NAME
    Call OpenCommandPrompt(projectFolderPath)
    
End Sub
