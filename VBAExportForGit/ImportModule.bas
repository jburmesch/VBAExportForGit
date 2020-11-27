Attribute VB_Name = "ImportModule"
Option Explicit


Const IMPORT_FOLDER_NAME As String = "source"
Const ALWAYS_PICK_FILES As Boolean = False


Function GetImportFiles(wb As Workbook, fileTypes As Variant) As Variant
    Dim workbookName As String
    Dim importFolder As String
    Dim files As Variant
    
    workbookName = Split(wb.Name, ".")(0)
    importFolder = wb.path & "\" & IMPORT_FOLDER_NAME & "\" & workbookName
    
    If Not FolderExists(importFolder) Then
        files = PickFiles(IMPORT_FOLDER_NAME _
                                            & " folder not found. Please select files to import.", , fileTypes)
    Else
        If ALWAYS_PICK_FILES Then
            files = PickFiles("Please pick files to import", importFolder, fileTypes)
        
        Else
            files = AllImportableFiles(importFolder, fileTypes)
            
        End If
        
    End If
    GetImportFiles = files
    
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
        If Not IsMissing(fileTypes) Then
            .Filters.Clear
            .Filters.Add "File type", "*" & Join(fileTypes, ", *")
        
        End If
        .AllowMultiSelect = True
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


Function AllImportableFiles(importFolder As String, Optional fileTypes As Variant) As Variant
    Dim fso As New FileSystemObject
    Dim fol As Variant
    Dim fil As Variant
    Dim files As Variant
    Dim result() As Variant
    ReDim result(0)
    
    Set fol = fso.GetFolder(importFolder)
    For Each fil In fol.files
        If ExtensionInArray(fil.Name, fileTypes) Then
            If IsEmpty(result(0)) Then
                result(0) = fil.Name
                
            Else
                ReDim Preserve result(UBound(result) + 1)
                result(UBound(result)) = fil.Name
            
            End If
        
        End If
    
    Next
    AllImportableFiles = result
    
End Function


Function ExtensionInArray(entry As Variant, theArray As Variant) As Boolean
    Dim item As Variant
    
    ExtensionInArray = False
    For Each item In theArray
        If Right(entry, Len(item)) = item Then
            ExtensionInArray = True
            Exit For
        
        End If
    
    Next

End Function


Function ComponentExists(componentName As String, components As VBComponents) As Boolean
    Dim c As VBComponent
    
    ComponentExists = False
    For Each c In components
        If c.Name = componentName Then
            ComponentExists = True
            Exit For
        
        End If
    
    Next
    
End Function


Function GetComponentName(componentPath As Variant) As String
    Dim folderSplit As Variant
    Dim fileName As String
    
    folderSplit = Split(componentPath, "\")
    fileName = Split(folderSplit(UBound(folderSplit)), ".")(0)
    GetComponentName = fileName
    
End Function


Sub Import(importFiles As Variant, wb As Workbook)
    Dim i As Integer
    Dim components As VBComponents
    Dim componentName As String
    
    Set components = wb.VBProject.VBComponents
    For i = 1 To UBound(importFiles)
        componentName = GetComponentName(importFiles(i))
        If ComponentExists(componentName, components) Then
            If PromptForOverwrite(componentName) Then
                components.Remove components(componentName)
                components.Import importFiles(i)
                
            Else
                components.Import importFiles(i)
                
            End If
            
        Else
            components.Import importFiles(i)
            
         End If
         
    Next
    
End Sub


Function PromptForOverwrite(componentName As String) As Boolean
    Dim ans As Variant
    
    PromptForOverwrite = False
    ans = MsgBox(componentName & " already exists in this workbook.  Would you like to overwrite it?", vbYesNoCancel)
    If ans = vbCancel Then
        End
        
    ElseIf ans = vbYes Then
        PromptForOverwrite = True
        
    End If
    
End Function
