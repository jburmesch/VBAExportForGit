Attribute VB_Name = "ExportModule"
Option Explicit


Const EXPORT_FOLDER_NAME As String = "source"
Const IGNORE_LIST As String = "ignorelist.txt"


'Creates export folder and subfolder if they don't exist, and exports components there.
Sub ExportComponents(components As VBComponents, _
                                    projectFolderPath As String, workbookName As String)
    Dim component As VBComponent
    Dim wbFolderPath As String
    Dim filePath As String
    
    For Each component In components
        wbFolderPath = CreateFolders(projectFolderPath, workbookName)
        filePath = SetFilePath(wbFolderPath, component)
        component.Export filePath
        
    Next
    
End Sub


'checks if folders exist and creates them if they don't
Function CreateFolders(projectFolderPath As String, subFolderName As String) As String
    
    Dim subFolderPath As String
    
    If Not FolderExists(projectFolderPath) Then Call CreateFolder(projectFolderPath)
    subFolderPath = projectFolderPath & "\" & subFolderName
    If Not FolderExists(subFolderPath) Then Call CreateFolder(subFolderPath)
    CreateFolders = subFolderPath
    
End Function


'copys IGNORE_LIST file into a .gitignore inside folderPath
Sub CopyGitIgnore(folderPath As String)
    If Not FileExists(folderPath & "\.gitignore") Then
        'chr(34) = double quotes, to ensure that any
        'spaces in the file path don't cause problems.
        Shell "cmd /c copy /a /v /y " _
                 & Chr(34) & ThisWorkbook.path & "\" & IGNORE_LIST & Chr(34) _
                 & " " & Chr(34) & folderPath & "\.gitignore" & Chr(34)
    
    End If
    
End Sub


Function FileExists(filePath As String) As Boolean
    Dim fso As New FileSystemObject
    FileExists = False
    
    If fso.FileExists(filePath) Then FileExists = True
        
End Function

'opens a command prompt window at folderPath
Sub OpenCommandPrompt(folderPath As String)
    Shell "cmd /k cd /d " & Chr(34) & folderPath & Chr(34), vbNormalFocus
    
End Sub

'concatenates folder path, component name and appropriate file extension
Function SetFilePath(folderPath As String, component As VBComponent) As String
    Dim extension As String
    
    extension = GetExtension(component)
    SetFilePath = folderPath & "\" & component.Name & extension
    
End Function


'adapted from: https://stackoverflow.com/questions/
'10803834/create-a-folder-and-sub-folder-in-excel-vba
Function FolderExists(path As String)
    Dim fso As New FileSystemObject
    FolderExists = False
    
    If fso.FolderExists(path) Then FolderExists = True
    
End Function


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
    
    Case vbext_ct_Document
        GetExtension = ".cls"
        
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
