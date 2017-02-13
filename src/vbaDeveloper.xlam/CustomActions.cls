VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Interface with hooks for thisWorkbook open and close actions
'
' An implementation can for example open a number of workbooks, connect to a database, load data and initialize worksheets,
' or any other tasks that otherwise have to be done manually.

'Called after thisWorkbook is opened
Sub afterOpen()
End Sub

'Called before thisWorkbook is closed
Sub beforeClose()
End Sub

