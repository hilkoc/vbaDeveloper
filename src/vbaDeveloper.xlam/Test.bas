Attribute VB_Name = "Test"
Option Explicit

Public Sub testMyCustomActions_Open()
    Dim myCustomAction As CustomActions
    Set myCustomAction = New MyCustomActions
    myCustomAction.afterOpen
End Sub


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
