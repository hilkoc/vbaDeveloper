VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class ConsoleLogger
Implements Logger
Implements LoggerPrototype
' We are required to implement both interfaces, even if LoggerPrototype also implements Logger.

' This logger appends to the 'immediate window' in the Excel VBA IDE.


Private name As String ' The name of this logger


Private Function Logger_whereIsMyLog() As String
    Logger_whereIsMyLog = "Log output is printed to the immediate window."
End Function


Private Sub Logger_info(message As String, Optional msg2 As String, Optional msg3 As String)
    logMessage LogFactory.info, message, msg2, msg3
End Sub


Private Sub Logger_warn(message As String, Optional msg2 As String, Optional msg3 As String)
    logMessage LogFactory.warn, message, msg2, msg3
End Sub


Private Sub Logger_fatal(message As String, Optional msg2 As String, Optional msg3 As String)
    logMessage LogFactory.fatal, message, msg2, msg3
End Sub


Public Sub LoggerPrototype_setName(loggerName As String)
    name = loggerName
End Sub


Private Function LoggerPrototype_clone() As LoggerPrototype
    Dim clone As ConsoleLogger
    Set clone = New ConsoleLogger
    ' the clone's name will be overwritten, so no need to: clone.LoggerPrototype_setName name
    Set LoggerPrototype_clone = clone
End Function


' The actual implementation of the 'info' , 'warn' and 'fatal' subs.
Private Sub logMessage(status As String, message As String, Optional msg2 As String, Optional msg3 As String)
    Dim formatted As String
    formatted = LogFactory.formatLogMessage(status & "|" & name, message, msg2, msg3)
    Debug.Print formatted
End Sub
