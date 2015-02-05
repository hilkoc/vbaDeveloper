VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyCustomActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CustomActions
' This class serves as an example only.


Private Const MY_FAVORITE_WORKBOOK_PATH As String = "C:\path\to\myFavoriteWorkbook\"
Private Const MY_FAVORITE_WORKBOOK_NAME As String = "example.xlsm"


' Doc: See CustomActions
Private Sub CustomActions_afterOpen()
    On Error GoTo CustomActions_afterOpen_Error

    If Not IsWorkBookOpen(MY_FAVORITE_WORKBOOK_NAME) Then
        ' The next line usually raises an error, therefore it is commented out.
        'Application.Workbooks.Open (MY_FAVORITE_WORKBOOK_PATH & MY_FAVORITE_WORKBOOK_NAME)
    End If

    On Error GoTo 0
    Exit Sub
CustomActions_afterOpen_Error:
    ErrorHandling.handleError "vbaDeveloper.MyCustomActions afterOpen"
End Sub

' Doc: See CustomActions
Private Sub CustomActions_beforeClose()
End Sub

Function IsWorkBookOpen(wkbName As String) As Boolean
    Dim wBook As Workbook
    On Error Resume Next
    Set wBook = Workbooks(wkbName)
    IsWorkBookOpen = Not (wBook Is Nothing)
    On Error GoTo 0
End Function


