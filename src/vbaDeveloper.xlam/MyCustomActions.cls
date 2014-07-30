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

Private Const ROOT_FXS_PATH As String = "C:\dev\fx_spreadsheets\trunk\Pricing Tools\Loader\"
Private Const LOADER_NAME As String = "loader.xla"

Private Const gridAddin As String = "CFXGrid.xla"
Private Const gridUsername As String = "FoxGrid_FID"
Private Const gridPassword As String = "ha86dk3r"
Private Const gridGrid As String = "LNDEV"
Private Const gridApp As String = "WIN_CFX_3.27_patchVolVolCorrCurve"

Private Const foxAddin As String = "CFXDatabase.xla"
Private Const foxUsername As String = "hc15196"
Private Const foxPassword As String = "*******"


' Doc: See CustomActions
Private Sub CustomActions_afterOpen()
On Error GoTo CustomActions_afterOpen_Error
    
    If IsWorkBookOpen(LOADER_NAME) Then
        'connect to fox and to grid
        debugPrint "loader is open, connecting to fox and to grid"
        
        Application.Run gridAddin & "!modGrid.ConnectToGrid", gridUsername, gridPassword, gridGrid, gridApp
        
        Application.Run foxAddin & "!modAPI.FoxLogin", foxUsername, foxPassword 'connect to default server
    End If
    'Workbooks.Open ROOT_FXS_PATH & LOADER_NAME
    
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

