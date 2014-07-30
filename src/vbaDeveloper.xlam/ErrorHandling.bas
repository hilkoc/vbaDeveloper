Attribute VB_Name = "ErrorHandling"
Option Explicit

Public Sub RaiseError(errNumber As Integer, Optional errSource As String = "", Optional errMessage As String = "")
    If errSource = "" Then 'set default values
        errSource = Err.Source
        errMessage = Err.Description
    End If
    Err.Raise vbObjectError + errNumber, errSource, errMessage
End Sub

Public Sub handleError(Optional errLocation As String = "")
    Dim errorMessage As String
    errorMessage = "Error in " & errLocation & ", [" & Err.Source & "] : error number " & Err.Number & vbNewLine & Err.Description
    debugPrint errorMessage
    MsgBox errorMessage, vbCritical, "vbaDeveloper ErrorHandler"
End Sub


Public Sub debugPrint(message As String)
    Debug.Print message
End Sub
