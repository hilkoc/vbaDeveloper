Private eventListener As CExcelEvents

' Initialize member to listen to excel events
Private Sub Workbook_Open()
    Set eventListener = New CExcelEvents
End Sub

' Clean up our private members
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Set eventListener = Nothing
End Sub


