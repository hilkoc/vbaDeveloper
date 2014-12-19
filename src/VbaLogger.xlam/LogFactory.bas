Attribute VB_Name = "LogFactory"
Option Explicit

' Requires reference to Microsoft Scripting Runtime.
' The LogFactory holds common configuration for all loggers and
' creates and manages instances of loggers.



Public Const info As String = "INFO"
Public Const warn As String = "WARN"
Public Const fatal As String = "FATAL"

Private loggerMap As Dictionary ' maps loggerName to instance
Private thePrototype As LoggerPrototype ' the prototype from which to create new logger instances


' Common log configuration for all loggers
Private Const dateFormat As String = "YYMMDD hh:mm.ss"
Private Const SEP As String = "|" ' the separator between different parts on the line
Private Const logDirPath As String = "C:\Temp\"

' usage:
'
'Property Get log() As Logger
'    Set log = LogFactory.getLogger(ThisWorkbook.name)
'End Property


' The logger that is returned will depend on the configuration and prototype above.
Public Function getLogger(loggerName As String) As Logger
    If loggers.Exists(loggerName) Then
        Set getLogger = loggers.Item(loggerName)
        Exit Function
    End If
    Dim loggerInstance As LoggerPrototype
    Set loggerInstance = prototype.clone()
    loggerInstance.setName loggerName
    loggers.Add Key:=loggerName, Item:=loggerInstance
    Set getLogger = loggerInstance
End Function


' Configure the prototype to use for producing all logger instances.
' Expects a fully configured logger instance.
Public Sub configurePrototype(instance As LoggerPrototype)
    Set loggers = New Dictionary
    Set prototype = instance
End Sub


' Creates a new Logger instance that appends to immediate window.
Public Function getConsoleLogger(loggerName As String) As Logger
    Dim loggerInstance As ConsoleLogger
    Set loggerInstance = New ConsoleLogger
    loggerInstance.LoggerPrototype_setName loggerName
    Set getConsoleLogger = loggerInstance
End Function


' Creates a new Logger instance that appends to the file at the given full absolute path.
Public Function getFileLogger(loggerName As String) As Logger
    Dim fileLoggerInstance As FileLogger
    Set fileLoggerInstance = New FileLogger
    fileLoggerInstance.setLogDir logDirPath
    fileLoggerInstance.LoggerPrototype_setName loggerName
    Set getFileLogger = fileLoggerInstance
End Function


' Returns the loggerMap and initializes it if necessary.
Private Property Get loggers() As Dictionary
    If loggerMap Is Nothing Then
        Set loggerMap = New Dictionary
    End If
    Set loggers = loggerMap
End Property


' Returns the prototype and initializes it with a default logger if necessary.
Private Property Get prototype() As LoggerPrototype
    If thePrototype Is Nothing Then
        Set thePrototype = getConsoleLogger(ThisWorkbook.name)
    End If
    Set prototype = thePrototype
End Property


' Formats the given message parts into one string with date and status in front, all separated by the SEP character.
Function formatLogMessage(status As String, message As String, Optional msg2 As String, Optional msg3 As String)
    Dim formatted As String
    formatted = Format(Now(), dateFormat) & SEP & status & SEP & message
    If Not msg2 = "" Then
        formatted = formatted & SEP & msg2
    End If
    If Not msg3 = "" Then
        formatted = formatted & SEP & msg3
    End If
    formatLogMessage = formatted
End Function


' Release all logger instances.
Public Sub clear()
    Set loggers = Nothing
    Set prototype = Nothing
End Sub


Sub testFileLogger()
    clear
    
    Dim thePrototype As Logger
    Set thePrototype = getFileLogger(ThisWorkbook.name)
    configurePrototype thePrototype
    
    Dim log As Logger
    Set log = getLogger(ThisWorkbook.name)
    
    Debug.Print "logging to " & log.whereIsMyLog()
    
    log.info "hello it works"
    log.warn "hello it works2"
    log.fatal " a fatal message"
End Sub


