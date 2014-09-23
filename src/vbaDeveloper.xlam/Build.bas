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

Option Explicit

Private Const IMPORT_DELAY As String = "00:00:03"

'We need to make these variables public such that they can be read by application.ontime
Public componentsToImport As Dictionary 'Key = componentName, Value = componentFilePath
Public sheetsToImport As Dictionary 'Key = componentName, Value = File object
Public vbaProjectToImport As VBProject

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
    If component.codeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.codeModule.Lines(1, 1))
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
    outStream.Write (component.codeModule.Lines(1, component.codeModule.CountOfLines))
    outStream.Close
End Sub


' Usually called after the given workbook is opened
Public Sub importVbaCode(vbaProject As VBProject)
    'find project files
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
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

    'initialize globals for Application.OnTime
    Set componentsToImport = New Dictionary
    Set sheetsToImport = New Dictionary
    Set vbaProjectToImport = vbaProject

    Dim projContents As Folder
    Set projContents = fso.GetFolder(export_path)
    Dim file As Object
    For Each file In projContents.Files()
        'check if and how to import the file
        checkHowToImport file
    Next

    Dim componentName As String
    Dim vComponentName As Variant
    'Remove all the modules and class modules
    For Each vComponentName In componentsToImport.Keys
        componentName = vComponentName
        removeComponent vbaProject, componentName
    Next
    'Then import them
    Debug.Print "Invoking 'Build.importComponents'with Application.Ontime with delay " & IMPORT_DELAY
    ' to prevent duplicate modules, like MyClass1 etc.
    Application.OnTime Now() + TimeValue(IMPORT_DELAY), "'Build.importComponents'"
    Debug.Print "almost finished importing code for " & vbaProject.name
End Sub

Private Sub checkHowToImport(file As Object)
    Dim fileName As String
    fileName = file.name
    Dim componentName As String
    componentName = left(fileName, InStr(fileName, ".") - 1)
    If componentName = "Build" Then '"don't remove or import ourself
        Exit Sub
    End If

    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = right(fileName, 4)
        Select Case lastPart
            Case ".cls" ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And right(fileName, 10) = ".sheet.cls" Then
                    'import lines into sheet: importLines vbaProjectToImport, file
                    sheetsToImport.Add componentName, file
                Else
                    'importComponent vbaProject, file
                    componentsToImport.Add componentName, file.Path
                End If
            Case ".bas", ".frm"
                'importComponent vbaProject, file
                componentsToImport.Add componentName, file.Path
            Case Else
                'do nothing
                Debug.Print "Skipping file " & fileName
        End Select
    End If
End Sub

' Only removes the vba component if it exists
Private Sub removeComponent(vbaProject As VBProject, componentName As String)
    If componentExists(vbaProject, componentName) Then
        Dim c As VBComponent
        Set c = vbaProject.VBComponents(componentName)
        Debug.Print "removing " & c.name
        vbaProject.VBComponents.Remove c
    End If
End Sub

Public Sub importComponents()
    If componentsToImport Is Nothing Then
        Debug.Print "Failed to import! 'Dictionary 'componentsToImport' was not initialized."
        Exit Sub
    End If
    Dim componentName As String
    Dim vComponentName As Variant
    For Each vComponentName In componentsToImport.Keys
        componentName = vComponentName
        importComponent vbaProjectToImport, componentsToImport(componentName)
    Next

    'Import the sheets
    For Each vComponentName In sheetsToImport.Keys
        componentName = vComponentName
        importLines vbaProjectToImport, sheetsToImport(componentName)
    Next

    Debug.Print "Finished importing code for " & vbaProjectToImport.name
    'We're done, clear globals explicitly to free memory
    Set componentsToImport = Nothing
    Set vbaProjectToImport = Nothing
End Sub

' Assumes any component with same name has already been removed
Private Sub importComponent(vbaProject As VBProject, filePath As String)
    Debug.Print "Importing component from  " & filePath
    vbaProject.VBComponents.Import filePath
End Sub


Private Sub importLines(vbaProject As VBProject, file As Object)
    Dim componentName As String
    componentName = left(file.name, InStr(file.name, ".") - 1)
    Dim c As VBComponent
    If Not componentExists(vbaProject, componentName) Then
        'Create a sheet to import this code into. We cannot set the ws.codeName property which is read-only,
        ' instead we set its vbComponent.name which leads to the same result.
        Dim addedSheetCodeName As String
        addedSheetCodeName = addSheetToWorkbook(componentName, vbaProject.fileName)
        Set c = vbaProject.VBComponents(addedSheetCodeName)
        c.name = componentName
    End If
    Set c = vbaProject.VBComponents(componentName)
    Debug.Print "Importing lines from " & componentName & " into component " & c.name

    ' At this point compilation errors may cause a crash, so we ignore those
    On Error Resume Next
    c.codeModule.DeleteLines 1, c.codeModule.CountOfLines
    c.codeModule.AddFromFile file.Path
    On Error GoTo 0
End Sub


Public Function componentExists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    componentExists = True
    Exit Function
doesnt:
    componentExists = False
End Function


' Returns a reference to the workbook. Opens it if it is not already opened.
' Raises error if the file cannot be found.
Public Function openWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath) 'can raise error
    End If
    Set openWorkbook = wb
End Function

' Returns the CodeName of the added sheet or an empty String if the workbook could not be opened.
Public Function addSheetToWorkbook(sheetName As String, workbookFilePath As String) As String
    Dim wb As Workbook
    On Error Resume Next 'can throw if given path does not exist
    Set wb = openWorkbook(workbookFilePath)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Dim ws As Worksheet
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = sheetName
        'ws.CodeName = sheetName: cannot assign to read only property
        Debug.Print "Sheet added " & sheetName
        addSheetToWorkbook = ws.CodeName
    Else
        Debug.Print "Skipping file " & sheetName & ". Could not open workbook " & workbookFilePath
        addSheetToWorkbook = ""
    End If
End Function
