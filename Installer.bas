Attribute VB_Name = "Installer"

Option Explicit

'1) Create an Excel file called Installer.xlsm in same folder than Installer.bas:
'   *\GIT\vbaDeveloper-master\

'2) Open the VB Editor (Alt+F11) right click on the active project and choose Import a file and chose:
'    *\GIT\vbaDeveloper-master\Installer.bas


'3a) Go in Tools--> References and activate:
'   - Microsoft Scripting Runtime
'   - Microsoft Visual Basic for Application Extensibility X.X

'3b) Enable programatic access to VBA:
'       File -> Options -> Trust Center, Trust Center Settings, -> Macros,
'       tick the box: 'Enable programatic access to VBA'  (In excel 2010: 'Trust access to the vba project object model')

'4) Run the Sub AutoInstaller in the module Installer


'5) Create a new excel file and also open the file vbaDeveloper.xlam located in the folder: *\GIT\vbaDeveloper-master\

'6) Make step 3a and 3b again for this file and run the sub testImport located in the module "Build".


Sub AutoInstaller()

'Prepare variable
Dim CurrentWB As Workbook

Dim textline As String, strPathOfBuild As String, strLocationXLAM As String

'Set the variables
Set CurrentWB = ThisWorkbook
Set NewWB = Workbooks.Add

'Import code form Build.bas  to the new workbook
strPathOfBuild = CurrentWB.Path & "\src\vbaDeveloper.xlam\Build.bas"
NewWB.VBProject.VBComponents.Import strPathOfBuild

    'Rename the project (in the VBA) to vbaDeveloper
    NewWB.VBProject.Name = "vbaDeveloper"

    
    'In VB Editor, press F4, then under Microsoft Excel Objects, select ThisWorkbook.Set the property 'IsAddin' to TRUE
    NewWB.IsAddin = True
    'In VB Editor, menu File-->Save Book1; Save as vbaDeveloper.xlam in the same directory as 'src'

    strLocationXLAM = CurrentWB.Path
    NewWB.SaveAs strLocationXLAM & "\vbaDeveloper.xlam", xlOpenXMLAddIn
        
    'Close excel. Open excel with a new workbook, then open the just saved vbaDeveloper.xlam
    NewWB.Close savechanges:=False

End Sub

