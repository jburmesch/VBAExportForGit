Attribute VB_Name = "ExportModule"
Option Explicit


Const EXPORT_FOLDER_NAME As String = "source"
Const IGNORE_LIST As String = "ignorelist.txt"


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
    Call OpenCommandPrompt(projectFolderPath)
    Call CopyGitIgnore(projectFolderPath)
    ThisWorkbook.Close
    
End Sub


'Creates export folder and subfolder if they don't exist, and exports components there.
Sub ExportComponents(components As VBComponents, _
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


'checks if folders exist and creates them if they don't
Function CreateFolders(projectFolderPath As String, subFolderName As String) As String
    
    Dim subFolderPath As String
    
    Call EnsureFolderExists(projectFolderPath)
    subFolderPath = projectFolderPath & "\" & subFolderName
    Call EnsureFolderExists(subFolderPath)
    CreateFolders = subFolderPath
    
End Function


'checks component types to see if they should be exported
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


'copys IGNORE_LIST file into a .gitignore inside folderPath
Sub CopyGitIgnore(folderPath As String)
    'chr(34) = double quotes, to ensure that any
    'spaces in the file path don't cause problems.
    Shell "cmd /c copy /a /v /y " & IGNORE_LIST & " " _
            & Chr(34) & folderPath & "\.gitignore" & Chr(34)
    
End Sub

'opens a command prompt window at folderPath
Sub OpenCommandPrompt(folderPath As String)
    Shell "cmd /K cd " & Chr(34) & folderPath & Chr(34), vbNormalFocus
    
End Sub

'concatenates folder path, component name and appropriate file extension
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
    MsgBox "Folder creation error."
    Exit Sub
    
End Sub


'returns the appropriate extension based on the component type
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
