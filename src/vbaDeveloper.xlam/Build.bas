Attribute VB_Name = "Build"
'''
' Build instructions:
' 1. Open a new workbook in excel, then open the VB editor (Alt+F11)  and from the menu File->Import, import these files:
'     * src/vbaDeveloper.xlam/Build.bas
' 2. From tools references... add
'     * Microsoft Visual Basic for Applications Extensibility 5.3
'     * Microsoft Scripting Runtime
' 3. Rename the project to 'vbaDeveloper'
' 5. Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'
' 6. In VB Editor, press F4, then under Microsfoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
' 7. In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'
' 8. Close excel. Open excel with a new workbook, then open the just saved vbaDeveloper.xlam
' 9. Let vbaDeveloper import its own code. Put the cursor in the function 'testImport' and press F5
' 10.If necessary rename module 'Build1' to Build. Menu File-->Save vbaDeveloper.xlam
'''



Public Sub testImport()
    Dim proj_name As String
    proj_name = "vbaDeveloper"
    
    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode vbaProject
End Sub


Public Sub testExport()
    Dim proj_name As String
    proj_name = "vbaDeveloper"

    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.exportVbaCode vbaProject
End Sub

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
        Debug.Print "Created Folder " & proj_root
    End If
    If Not fso.FolderExists(export_path) Then
        fso.CreateFolder export_path
        Debug.Print "Created Folder " & export_path
    End If
    
    Debug.Print "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In vbaProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        If hasCodeToExport(component) Then
            'Debug.Print "exporting type is " & component.Type
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
        'Debug.Print firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function

'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = ".cls")
    Debug.Print "exporting " & component.name & extension
    component.Export exportPath & "\" & component.name & extension
End Sub

'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = ".sheet.cls"
    Dim fileName As String
    fileName = exportPath & "\" & component.name & extension
    Debug.Print "exporting " & component.name & extension
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
        Debug.Print "No file name for project " & vbaProject.name & ", skipping"
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
        Debug.Print "Could not find folder " & proj_root
        Exit Sub
    End If
    If Not fso.FolderExists(export_path) Then
        Debug.Print "Could not find folder " & export_path
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
                        Debug.Print "Skipping file " & fileName
                End Select
            End If
        Next
    End If

    Debug.Print "imported code for " & vbaProject.name
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
        Debug.Print "removing " & component_name & "  " & c.name
        vbaProject.VBComponents.Remove c
    End If
    Debug.Print "Importing component " & component_name & " from  " & file.Path
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
    Debug.Print "Importing lines from " & component_name & " into component " & c.name
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

