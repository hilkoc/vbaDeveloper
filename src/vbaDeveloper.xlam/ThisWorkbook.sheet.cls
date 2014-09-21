Option Explicit

Private listener As EventListener
Private customAction As CustomActions

' Initialize member to listen to excel events
Private Sub Workbook_Open()
    Debug.Print "vbaDeveloper thisWorkbook_open()"
    Set listener = New EventListener
    Set customAction = New MyCustomActions
    menu.createMenu
    customAction.afterOpen
End Sub

' Clean up our private members
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "vbaDeveloper thisWorkbook_BeforeClose()"
    menu.deleteMenu
    customAction.beforeClose
    Set customAction = Nothing
    Set listener = Nothing
End Sub
    