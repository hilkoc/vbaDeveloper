Attribute VB_Name = "Installer"

'1) Create an Excel file called Installer.xlsm in the folder:
'   *\GIT\vbaDeveloper-master\srcvbaDeveloper.xlam\

'2) Open the VB Editor (Alt+F11) right click on the active project and choose Import a file and chose:
'    *\GIT\vbaDeveloper-master\srcvbaDeveloper.xlam\Installer.bas

'3a) Go in Tools--> References and activate:
'   - Microsoft Scripting Runtime
'   - Microsoft Visual Basic for Application Extensibility X.X

'3b) Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'  (In excel 2010: 'Trust access to the vba project object model')

'4) Run the Sub AutoInstaller in the module Installer

'5) Crate a new excel file and also open the file vbaDeveloper.xlam located in the folder: *\GIT\vbaDeveloper-master\

'6) Make step 3a and 3b again for this file and run the sub testImport located in the modul "Build".

Sub AutoInstaller()

'Prepare variable
Dim CurrentWB As Workbook
Dim NewWB As Workbook
Dim WBKvbaDeveloper As VBIDE.VBProject
Dim WBKvbaDeveloperModule As VBIDE.VBComponent
Dim strPathOfBuild As String

'Set the variables
Set CurrentWB = ActiveWorkbook
Set NewWB = Workbooks.Add
Set WBKvbaDeveloper = NewWB.VBProject
Set WBKvbaDeveloperModule = WBKvbaDeveloper.VBComponents.Add(vbext_ct_StdModule)
Set Modulecode = WBKvbaDeveloperModule.CodeModule


'Move code form Build.bas  to the new workbook
strPathOfBuild = CurrentWB.Path & "/Build.bas"
Dim LineNb As Double
LineNb = 1
Open strPathOfBuild For Input As #1

Do Until EOF(1)
    'Copy content of Build.bas in the module of the Excel File
    Line Input #1, textline
    Modulecode.InsertLines LineNb, textline
    LineNb = LineNb + 1
    'ModuleCode.InsertLines 1, "sub test()"
Loop

Modulecode.DeleteLines 1

Close #1
    'Rename the project (in the VBA) to vbaDeveloper
    WBKvbaDeveloper.Name = "vbaDeveloper"
    WBKvbaDeveloperModule.Name = "Build"
    'WBKvbaDeveloper.VBComponents(1).Name
    
    'In VB Editor, press F4, then under Microsoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
    NewWB.IsAddin = True
    'In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'
    strLocationXLAM = Replace(CurrentWB.Path, "src\vbaDeveloper.xlam", "") 
    NewWB.SaveAs strLocationXLAM & "vbaDeveloper.xlam", xlOpenXMLAddIn
	
    'Close excel. Open excel with a new workbook, then open the just saved vbaDeveloper.xlam
    NewWB.Close savechanges:=False

    
End Sub

