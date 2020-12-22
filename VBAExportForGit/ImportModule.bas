Attribute VB_Name = "ImportModule"
Option Explicit


Const IMPORT_FOLDER_NAME As String = "source"
Const ALWAYS_PICK_FILES As Boolean = True

'Create and return a list of file paths to import
Function GetImportFiles(wb As Workbook, fileTypes As Variant) As Variant
    Dim workbookName As String
    Dim importFolder As String
    Dim files As Variant
    
    'remove the extension from the workbook name
    workbookName = Split(wb.Name, ".")(0)
    'put together the default export folder name for the active workbook
    importFolder = wb.path & "\" & IMPORT_FOLDER_NAME & "\" & workbookName
    
    'if it doesn't exist, allow the user to browse and find the files they'd like to import
    If Not FolderExists(importFolder) Then
        files = PickFiles(IMPORT_FOLDER_NAME _
                                & " folder not found. Please select files to import.", , fileTypes)
    'either open a file picker in the default folder, or automatically import all files,
    'depending on ALWAYS_PICK_FILES setting.
    Else
        If ALWAYS_PICK_FILES Then
            files = PickFiles("Please pick files to import", importFolder, fileTypes)
        
        Else
            files = AllImportableFiles(importFolder, fileTypes)
            
        End If
        
    End If
    GetImportFiles = files
    
End Function

'open a file picker and return an array of paths to selected files
Function PickFiles(Optional prompt As String, Optional folderPath As String, _
                            Optional fileTypes As Variant) As Variant
    Dim results() As Variant
    Dim item As Variant
    Dim exitCode As Integer
    Dim i As Integer
    
    With Application.FileDialog(msoFileDialogFilePicker)
        'set title if provided
        If Not IsMissing(prompt) Then .Title = prompt
        'set starting folder path if provided
        If Not IsMissing(folderPath) Then .InitialFileName = folderPath
        'set file types if provided
        If Not IsMissing(fileTypes) Then
            .Filters.Clear
            .Filters.Add "File type", "*" & Join(fileTypes, ", *")
        
        End If
        .AllowMultiSelect = True
        'show the picker and save its exit code (0 = cancel).
        exitCode = .Show
        'if multiple files were selected, add them all to results array
        If exitCode <> 0 Then
            ReDim results(1 To .SelectedItems.Count)
            For Each item In .SelectedItems
                i = i + 1
                results(i) = item
            
            Next
            
        'if cancel was clicked, end.
        Else
            End
        
        End If
        
    End With
    PickFiles = results
    
End Function


'return array of paths of all files of given fileTypes(extensions) in importFolder
Function AllImportableFiles(importFolder As String, _
                                          fileTypes As Variant) As Variant
    Dim fso As New FileSystemObject
    Dim fol As Variant
    Dim fil As Variant
    Dim files As Variant
    Dim result() As Variant
    ReDim result(1 To 1)
    
    Set fol = fso.GetFolder(importFolder)
    'check all files in import folder
    For Each fil In fol.files
        'see if the file names end with one of the desired extensions
        If ExtensionInList(fil.Name, fileTypes) Then
            'add it to result if it is
            If IsEmpty(result(1)) Then
                'just add it if the array is empty
                result(1) = importFolder & "\" & fil.Name
                
            Else
                'if not empty, expand array and add
                ReDim Preserve result(1 To UBound(result) + 1)
                result(UBound(result)) = importFolder & "\" & fil.Name
            
            End If
        
        End If
    
    Next
    AllImportableFiles = result
    
End Function


'check if a file's extension is on the list
Function ExtensionInList(fileName As Variant, extensionList As Variant) As Boolean
    Dim extension As Variant
    
    ExtensionInArray = False
    For Each extension In extensionList
        'see if the rightmost (extension length) characters of the
        'file name match the extension
        If Right(fileName, Len(extension)) = extension Then
            ExtensionInList = True
            Exit For
        
        End If
    
    Next

End Function


'see if a component of the same name already exists in the given components list
Function ComponentExists(componentName As String, _
                                        components As VBComponents) As Boolean
    Dim c As VBComponent
    
    ComponentExists = False
    For Each c In components
        If c.Name = componentName Then
            ComponentExists = True
            Exit For
        
        End If
    
    Next
    
End Function


'get the component name from a component path
Function GetComponentName(componentPath As Variant) As String
    Dim folderSplit As Variant
    Dim fileName As String
    
    folderSplit = Split(componentPath, "\")
    fileName = Split(folderSplit(UBound(folderSplit)), ".")(0)
    GetComponentName = fileName
    
End Function


'import an array of paths to vb components
Sub Import(importFiles As Variant, wb As Workbook)
    Dim i As Integer
    Dim components As VBComponents
    Dim componentName As String
    
    Set components = wb.VBProject.VBComponents
    'loop through all import files
    For i = 1 To UBound(importFiles)
        'get the component name of the file
        componentName = GetComponentName(importFiles(i))
        'see if it already exists in the workbook's VBComponents
        If ComponentExists(componentName, components) Then
            'if it does, ask the user if they'd like to overwrite it
            If PromptForOverwrite(componentName) Then
                If componentName = "ThisWorkbook" Then
                    components.Import importFiles(i)
                
                Else
                    components.Remove components(componentName)
                    components.Import importFiles(i)
                
                End If
            Else
                components.Import importFiles(i)
                
            End If
            
        'if it doesn't already exist, just import it
        Else
            components.Import importFiles(i)
            
         End If
         
    Next
    
End Sub


'ask the user if they'd like to overwrite an existing component
Function PromptForOverwrite(componentName As String) As Boolean
    Dim ans As Variant
    
    PromptForOverwrite = False
    ans = MsgBox(componentName & " already exists in this workbook. " _
                          & "Would you like to overwrite it?", vbYesNoCancel)
    If ans = vbCancel Then
        End
        
    ElseIf ans = vbYes Then
        PromptForOverwrite = True
        
    End If
    
End Function

'Notify a user that the default export folder doesn't exist, and give them the option to open
'a cmd window anyway.
Function NoExportFolderPrompt(folderName As String) As Boolean
    Dim ans As Variant
    
    NoExportFolderPrompt = False
    ans = MsgBox("No " & folderName & " folder exists for this workbook. " _
                          & "Would you like to open a command prompt in the workbbook " _
                          & "directory?", vbYesNo)
    If ans = vbYes Then NoExportFolderPrompt = True
    
End Function

