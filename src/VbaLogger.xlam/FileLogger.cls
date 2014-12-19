VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class FileLogger
Implements Logger
Implements LoggerPrototype
' We are required to implement both interfaces, even if LoggerPrototype also implements Logger.


' This logger appends to a file in the given directory

Private name As String ' The name of this logger
Private logDir As String 'The full absolute path to the directory for the logfile.


' Returns the full absolute path to the logfile.
Private Function logFilePath() As String
    logFilePath = logDir & name & ".log"
End Function


Private Function Logger_whereIsMyLog() As String
    Logger_whereIsMyLog = logFilePath()
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
    Dim clone As FileLogger
    Set clone = New FileLogger
    clone.setLogDir logDir
    ' the clone's name will be overwritten, so no need to: clone.LoggerPrototype_setName name
    Set LoggerPrototype_clone = clone
End Function


' The actual implementation of the 'info' , 'warn' and 'fatal' subs.
Private Sub logMessage(status As String, message As String, Optional msg2 As String, Optional msg3 As String)
    Dim formatted As String
    formatted = LogFactory.formatLogMessage(status, message, msg2, msg3)
    Debug.Print formatted
    
    Dim fso As FileSystemObject, ts As TextStream
    Set fso = New FileSystemObject
    Set ts = fso.OpenTextFile(logFilePath(), ForAppending, Create:=True)
    ts.WriteLine formatted
    ts.Close
End Sub


' The fullDirPath is expected to end with a '\' character.
Public Sub setLogDir(fullDirPath As String)
    logDir = fullDirPath
End Sub
