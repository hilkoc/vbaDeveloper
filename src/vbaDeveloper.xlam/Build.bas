Attribute VB_Name = "Build"
Option Explicit

Sub ExportCodeModules()

    'This code Exports all VBA modules
    Dim componentIndex
    Dim comName
    'TODO lookup the current path and use naming convention for the export dir
    
    Dim sExportPath As String
    'sExportPath = fso.GetParentFolderName(proj_filename) & "\src"
    sExportPath = ThisWorkbook.Path & "\src\" '"C:\dev\local\ExcelBuild\src\"
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(sExportPath) Then
        fso.CreateFolder sExportPath
    End If
    
    With ThisWorkbook.VBProject
        For componentIndex = 1 To .VBComponents.Count
            If .VBComponents(componentIndex).CodeModule.CountOfLines > 0 Then
                comName = .VBComponents(componentIndex).CodeModule.name
                .VBComponents(componentIndex).Export sExportPath & comName & ".vba"
            End If
        Next componentIndex
    End With

End Sub

Sub ImportCodeModules()
    Dim componentIndex As Integer
    Dim ModuleName As String
    With ThisWorkbook.VBProject
        For componentIndex = 1 To .VBComponents.Count
    
            ModuleName = .VBComponents(componentIndex).CodeModule.name
    
            If ModuleName <> "Build" Then
                If Right(ModuleName, 6) = "Macros" Then
                    .VBComponents.Remove .VBComponents(ModuleName)
                    .VBComponents.Import "X:\Data\MySheet\" & ModuleName & ".vba"
               End If
            End If
        Next componentIndex
    End With

End Sub

'''
' First time instructions:
' 1. Open VB editor (Alt+F11)  and from the menu File->Import, import these files:
'     * src/vbaDeveloper.xlam/Build.bas
'     * src/vbaDeveloper.xlam/CExcelEvents.cls
' 2. From tools references add
'     * Microsoft Visual Basic for Applications Extensibility 5.3
'     * Microsoft Scripting Runtime
' 3. Rename the project to 'vbaDeveloper'
' 4. Save as vbaDeveloper.xlam in the same directory as 'src'
' 5. Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'
' 6. Let vbaDeveloper import its own code. Put the cursor in the function 'testImport' and press F5.



' Usually called after the given workbook is saved
Public Sub exportVbaCode(vbaProject As VBProject)
    'locate and create the export directory if necessary
    Dim vbProjectFileName As String
    vbProjectFileName = vbaProject.fileName
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Exit Sub
    End If
    
    Dim fso As New Scripting.FileSystemObject
    Dim projDir As String
    projDir = fso.GetParentFolderName(vbProjectFileName)
    Dim proj_root As String
    proj_root = projDir & "\src\"
    Dim export_path As String
    export_path = proj_root & fso.GetFileName(vbProjectFileName)
    
    If Not fso.FolderExists(proj_root) Then
        fso.CreateFolder proj_root
        debugPrint "Created Folder " & proj_root
    End If
    If Not fso.FolderExists(export_path) Then
        fso.CreateFolder export_path
        debugPrint "Created Folder " & export_path
    End If
    
    debugPrint "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In vbaProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        If hasCodeToExport(component) Then
            'DebugPrint "exporting type is " & component.Type
            Select Case component.Type
                Case vbext_ct_ClassModule
                    exportComponent export_path, component
                Case vbext_ct_StdModule
                    exportComponent export_path, component, ".bas"
                Case vbext_ct_MSForm
                    exportComponent export_path, component, ".frm"
                Case vbext_ct_Document
                    exportLines export_path, component
                Case Else
                    'Raise "Unkown component type"
            End Select
        End If
    Next component
End Sub

Private Function hasCodeToExport(component As VBComponent) As Boolean
    hasCodeToExport = True
    If component.CodeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.CodeModule.Lines(1, 1))
        'DebugPrint firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function

'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = ".cls")
    debugPrint "exporting " & component.name & extension
    component.Export exportPath & "\" & component.name & extension
End Sub

'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = ".sheet.cls"
    Dim fileName As String
    fileName = exportPath & "\" & component.name & extension
    debugPrint "exporting " & component.name & extension
    'component.Export exportPath & "\" & component.name & extension
    Dim fso As New Scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = fso.CreateTextFile(fileName, True, False)
    outStream.Write (component.CodeModule.Lines(1, component.CodeModule.CountOfLines))
    outStream.Close
End Sub

' Usually called after the given workbook is opened
Public Sub importVbaCode(vbaProject As VBProject)
    'find project files
    Dim vbProjectFileName As String
    vbProjectFileName = vbaProject.fileName
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        debugPrint "No file name for project " & vbaProject.name & ", skipping"
        Exit Sub
    End If
    
    
    Dim fso As New Scripting.FileSystemObject
    Dim projDir As String
    projDir = fso.GetParentFolderName(vbProjectFileName)
    Dim proj_root As String
    proj_root = projDir & "\src\"
    Dim export_path As String
    export_path = proj_root & fso.GetFileName(vbProjectFileName)
    
    If Not fso.FolderExists(proj_root) Then
        debugPrint "Could not find folder " & proj_root
        Exit Sub
    End If
    If Not fso.FolderExists(export_path) Then
        debugPrint "Could not find folder " & export_path
        Exit Sub
    End If
    
    'for each file found:
    If fso.FolderExists(export_path) Then
        Dim proj_contents As Folder
        Set proj_contents = fso.GetFolder(export_path)
        
        Dim file As Object
        For Each file In proj_contents.Files()
            
            Dim fileName As String
            fileName = file.name
            'check if and how to import the file
            If Len(fileName) > 4 Then
                Dim lastPart As String
                lastPart = Right(fileName, 4)
                Select Case lastPart
                    Case ".cls" ' 10 == Len(".sheet.cls")
                        If Len(fileName) > 10 And Right(fileName, 10) = ".sheet.cls" Then
                            'import lines into sheet
                            importLines vbaProject, file
                        Else
                            'import component
                            importComponent vbaProject, file
                        End If
                    Case ".bas", ".frm"
                       'import component
                       importComponent vbaProject, file
                    Case Else
                        'do nothing
                        debugPrint "Skipping file " & fileName
                End Select
            End If
        Next
    End If

    debugPrint "imported code for " & vbaProject.name
End Sub

'Not used anymore
Private Function wantToImport(fileName As String) As Boolean
    wantToImport = False
    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = Right(fileName, 4)
        Select Case lastPart
            Case ".bas", ".frm"
               wantToImport = True
            Case ".cls" ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And Right(fileName, 10) = ".sheet.cls" Then
                    wantToImport = False 'For now we don't import these
                Else
                    wantToImport = True
                End If
            Case Else
                wantToImport = False
        End Select
    End If
End Function


Private Sub importComponent(vbaProject As VBProject, file As Object)
    Dim component_name As String
    component_name = Left(file.name, InStr(file.name, ".") - 1)
    
    If component_exists(vbaProject, component_name) Then
        'Remove it. (Sheets cannot be removed!)
        Dim c As VBComponent
        Set c = vbaProject.VBComponents(component_name)
        debugPrint "removing " & component_name & "  " & c.name
        vbaProject.VBComponents.Remove c
    End If
    debugPrint "Importing component " & component_name & " from  " & file.Path
    ' If we get duplicate modules, like MyClass1, try
    ' Application.OnTime (Now + TimeValue("00:00:01")), "function_name" vbaProject.VBComponents.Import file.Path
    vbaProject.VBComponents.Import file.Path
End Sub

Private Sub importLines(vbaProject As VBProject, file As Object)
    Dim component_name As String
    component_name = Left(file.name, InStr(file.name, ".") - 1)
                
    If Not component_exists(vbaProject, component_name) Then
        'Create a sheet and component to import this into
        '...skipping that for now
        Exit Sub
    End If
    Dim c As VBComponent
    Set c = vbaProject.VBComponents(component_name)
    debugPrint "Importing lines from " & component_name & " into component " & c.name
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    c.CodeModule.AddFromFile file.Path
End Sub

Private Function component_exists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    
    component_exists = True
    Exit Function
doesnt:
    component_exists = False
End Function



''''''''''''''''''

Public Function Hello() As String
    Hello = "hello it works"
End Function

