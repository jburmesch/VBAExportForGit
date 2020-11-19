Attribute VB_Name = "ExportModule"
Option Explicit


Const EXPORT_FOLDER_NAME As String = "source"
Const IGNORE_LIST As String = "ignorelist.txt"


Sub ExportComponents()
    Dim wb As Workbook
    Dim project As VBProject
    Dim projectFolderPath As String
    Dim workbookName As String
    
    Set wb = ActiveWorkbook
    workbookName = Split(wb.Name, ".")(0)
    Set project = wb.VBProject
    projectFolderPath = wb.path & "\" & EXPORT_FOLDER_NAME
    Call Export(project.VBComponents, projectFolderPath, workbookName)
    Call OpenCommandPrompt(projectFolderPath)
    Call CopyGitIgnore(projectFolderPath)
    
End Sub


Sub CopyGitIgnore(folderPath As String)
    Shell "cmd /c copy /a /v /y " & IGNORE_LIST & " " _
            & Chr(34) & folderPath & "\.gitignore" & Chr(34)
    
End Sub


Sub Export(components As VBComponents, _
                  projectFolderPath As String, workbookName As String)
    Dim component As VBComponent
    Dim wbFolderPath As String
    Dim filePath As String
    
    For Each component In components
        If IsExportable(component) Then
            wbFolderPath = CreateFolders(projectFolderPath, workbookName)
            filePath = SetFilePath(wbFolderPath, component)
            component.Export filePath
            
        End If
        
    Next
End Sub


Sub OpenCommandPrompt(folderPath As String)
    Shell "cmd /K cd " & Chr(34) & folderPath & Chr(34), vbNormalFocus
    
End Sub


Function CreateFolders(projectFolderPath As String, subFolderName As String) As String
    
    Dim path As String
    
    Call EnsureFolderExists(projectFolderPath)
    path = projectFolderPath & "\" & subFolderName
    Call EnsureFolderExists(path)
    CreateFolders = path
    
End Function


Function SetFilePath(folderPath As String, component As VBComponent) As String
    Dim extension As String
    
    extension = GetExtension(component)
    SetFilePath = folderPath & "\" & component.Name & extension
    
End Function


'adapted from: https://stackoverflow.com/questions/
'10803834/create-a-folder-and-sub-folder-in-excel-vba
Sub EnsureFolderExists(path As String)
    Dim fso As New FileSystemObject
    
    If Not fso.FolderExists(path) Then Call CreateFolder(path)
    
End Sub


'adapted from: https://stackoverflow.com/questions/
'10803834/create-a-folder-and-sub-folder-in-excel-vba
Sub CreateFolder(path As String)
    Dim fso As New FileSystemObject

    On Error GoTo Oops
    fso.CreateFolder path
    Exit Sub
    
Oops:
    MsgBox "There was an error creating the export folder"
    Exit Sub
    
End Sub


Function IsExportable(component As VBComponent) As Boolean
    If component.Type = vbext_ct_ClassModule _
        Or component.Type = vbext_ct_MSForm _
        Or component.Type = vbext_ct_StdModule _
    Then
        IsExportable = True
    
    Else
        IsExportable = False
    
    End If
    
End Function


Function GetExtension(component As VBComponent) As String
    Select Case component.Type
        
    Case vbext_ct_ClassModule
        GetExtension = ".cls"
    
    Case vbext_ct_StdModule
        GetExtension = ".bas"
    
    Case vbext_ct_MSForm
        GetExtension = ".frm"
    
    Case Else
        MsgBox "Component extension not found."
        Exit Function
        
    End Select
    
End Function
