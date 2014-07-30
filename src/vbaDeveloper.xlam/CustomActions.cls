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

' Interface class with hooks for thisWorkbook open and close actions

'Called after thisWorkbook is opened
Sub afterOpen()
End Sub

'Called before thisWorkbook is closed
Sub beforeClose()
End Sub
