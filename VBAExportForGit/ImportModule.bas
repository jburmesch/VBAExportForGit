Attribute VB_Name = "ImportModule"
Option Explicit


Const IMPORT_FOLDER_NAME As String = "source"
Const ALWAYS_PICK_FILES As Boolean = False


Sub VBAImportForGit()
    Dim importFiles() As String
    
    importFiles = GetImportFiles
    
End Sub


Function GetImportFiles() As String()
    Dim wb As Workbook
    Dim workbookName As String
    Dim importFolder As String
    Dim files() As String
    
    Set wb = ActiveWorkbook
    workbookName = Split(wb.Name, ".")(0)
    importFolder = wb.path & "\" & IMPORT_FOLDER_NAME & "\" & workbookName
    
    If Not FolderExists(importFolder) Then
        files = PickFiles(IMPORT_FOLDER_NAME _
                                            & " folder not found. Please select files to import.")
    Else
        If ALWAYS_PICK_FILES Then
            files = PickFiles("Please pick files to import", importFolder)
        
        Else
            files = AllImportableFiles(importFolder)
            
        End If
        
    End If
End Function


Function PickFiles(prompt As String) As String()
    Dim filePicker As FileDialog

    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    With filePicker
        .Title = prompt
        .AllowMultiSelect = True
        .Show
            If .SelectedItems.Count = 1 Then
                'save the path of whatever the user picked
                PickFile = .SelectedItems(1)
            Else
                End
            End If
    End With
End Function
