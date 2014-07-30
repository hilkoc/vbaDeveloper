Private eventListener As CExcelEvents
Private customAction As CustomActions

' Initialize member to listen to excel events
Private Sub Workbook_Open()
    debugPrint "vbaDeveloper thisWorkbook_open()"
    Set eventListener = New CExcelEvents
    Set customAction = New MyCustomActions
    customAction.afterOpen
End Sub

' Clean up our private members
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    debugPrint "vbaDeveloper thisWorkbook_BeforeClose()"
    customAction.beforeClose
    Set customAction = Nothing
    Set eventListener = Nothing
End Sub

    
    