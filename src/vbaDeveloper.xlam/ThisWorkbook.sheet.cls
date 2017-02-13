Option Explicit

'' The classes EventListener, CustomActions, MyCustomActions are not imported automatically by the build module.
'' After they are imported manually, the comments below can be uncommented. This will enable automatic code exports on save
'' and automatic code imports on open.


' Private listener As EventListener
' Private customAction As CustomActions

'' Initialize member to listen to excel events
Private Sub Workbook_Open()
    Debug.Print "vbaDeveloper thisWorkbook_open()"
    ' Set listener = New EventListener
    ' Set customAction = New MyCustomActions
    menu.createMenu
    ' customAction.afterOpen
End Sub


'' Clean up our private members
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "vbaDeveloper thisWorkbook_BeforeClose()"
    menu.deleteMenu
    ' If Not customAction Is Nothing Then
    '     customAction.BeforeClose
    '     Set customAction = Nothing
    ' End If
    ' Set listener = Nothing
End Sub