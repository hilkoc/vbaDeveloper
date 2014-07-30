Attribute VB_Name = "Test"
Option Explicit

Public Sub testMyCustomActions_Open()
    Dim myCustomAction As CustomActions
    Set myCustomAction = New MyCustomActions
    myCustomAction.afterOpen
End Sub



Public Sub testImport()
    Dim proj_name As String
    'proj_name = "vbaDeveloper"
    proj_name = "CCCvbaDeveloper"
    'proj_name = "testBuildAddin"
    
    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode vbaProject
End Sub


Public Sub testExport()
    Dim proj_name As String
    proj_name = "vbaDeveloper"
    'proj_name = "CCCvbaDeveloper"
    'proj_name = "testBuildAddin"

    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.exportVbaCode vbaProject
End Sub

'    Dim pr As Object
'    For Each pr In Application.VBE.VBProjects
'        DebugPrint pr.name
'    Next pr

'    Dim wkb As Workbook
'    For Each wkb In Application.Workbooks
'        wbName = wkb.Name
'        DebugPrint wbName
'    Next wkb
