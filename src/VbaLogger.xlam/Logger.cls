VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Interface Class Logger

' Returns (a description of) the location where the log messages are written to.
Public Function whereIsMyLog() As String
End Function


' Log the given messages at INFO level
Public Sub info(message As String, Optional msg2 As String, Optional msg3 As String)
End Sub


' Log the given messages at WARN level
Public Sub warn(message As String, Optional msg2 As String, Optional msg3 As String)
End Sub


' Log the given messages at FATAL level
Public Sub fatal(message As String, Optional msg2 As String, Optional msg3 As String)
End Sub
