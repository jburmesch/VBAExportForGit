Attribute VB_Name = "ImportModule"
Option Explicit


Const IMPORT_FOLDER_NAME As String = "source"
Const ALWAYS_PICK_FILES As Boolean = True


Function GetImportFiles() As String()
    Dim wb As Workbook
    Dim workbookName As String
    Dim importFolder As String
    Dim files As Variant
    Dim fileTypes As Variant

    fileTypes = Array("*.bas", "*.cls", "*.frm")
    
    Set wb = ActiveWorkbook
    workbookName = Split(wb.Name, ".")(0)
    importFolder = wb.path & "\" & IMPORT_FOLDER_NAME & "\" & workbookName
    
    If Not FolderExists(importFolder) Then
        files = PickFiles(IMPORT_FOLDER_NAME _
                                            & " folder not found. Please select files to import.", , fileTypes)
    Else
        If ALWAYS_PICK_FILES Then
            files = PickFiles("Please pick files to import", importFolder, fileTypes)
        
        Else
            files = AllImportableFiles(importFolder)
            
        End If
        
    End If
End Function


Function PickFiles(Optional prompt As String, Optional folderPath As String, _
                            Optional fileTypes As Variant) As Variant
    Dim filePicker As FileDialog
    Dim results() As Variant
    Dim item As Variant
    Dim i As Integer
    
    With Application.FileDialog(msoFileDialogFilePicker)
        If Not IsMissing(prompt) Then .Title = prompt
        If Not IsMissing(folderPath) Then .InitialFileName = folderPath
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "File type", Join(fileTypes, ", ")
        If .Show <> 0 And .SelectedItems.Count <> 1 Then
            ReDim results(1 To .SelectedItems.Count)
            For Each item In .SelectedItems
                i = i + 1
                results(i) = item
            
            Next
        
        ElseIf .Show <> 0 And .SelectedItems.Count = 1 Then
            results(1) = .SelectedItems(1)
        
        End If
        
    End With
    PickFiles = results
    
End Function

Function AllImportableFiles(importFolder As String) As String()

End Function
