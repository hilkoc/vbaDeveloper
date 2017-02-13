Attribute VB_Name = "BuildForm"
''
' BuildForm v1.0.0
' (c) Georges Kuenzli - https://github.com/gkuenzli/vbaDeveloper
'
' `BuildForm` exports a MSForm to 3 files :
'   - .frm : code of the component
'   - .frx : OLE ActiveX binary data => design data of the component
'   - .frd : JSON data => human-readable design data of the component
'
' @module FormSerializer
' @author gkuenzli
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const USERFORM_DATA_EXT As String = ".frd"
Private Const USERFORM_CODE_EXT As String = ".frm"
Private Const USERFORM_XOLE_EXT As String = ".frx"


''
' Export a MSForm to the specified path
' Do export component parts only when a change is detected
'
' @method exportMSForm
' @param {String} exportPath
' @param {VBComponent} component
''
Public Sub exportMSForm(exportPath As String, component As VBComponent)
    Dim FSO As New Scripting.FileSystemObject
    Dim frxChanged As Boolean
    Dim frmChanged As Boolean
    Dim storedFilePath As String
    Dim tempFilePath As String
    Dim tempFolder As String
    
    storedFilePath = JoinPath(exportPath, component.name)
    
    ' Create temporary folder
    tempFolder = storedFilePath & "~"
    If Not FSO.FolderExists(tempFolder) Then
        FSO.CreateFolder tempFolder
    End If
    tempFilePath = JoinPath(tempFolder, component.name)
    
    ' Export component to temporary files
    component.Export tempFilePath & USERFORM_CODE_EXT
    
    ' Comparing MSForm data (stored vs current)
    Dim storedData As String
    Dim currentData As String
    storedData = loadMSFormData(exportPath, component)
    currentData = FormSerializer.SerializeMSForm(component)
    frxChanged = getCleanCode(storedData) <> getCleanCode(currentData)
    
    ' Comparing MSForm code (stored vs current, hence temporary)
    Dim storedCode As String
    Dim currentCode As String
    storedCode = getCleanCode(loadTextFile(storedFilePath & USERFORM_CODE_EXT))
    currentCode = getCleanCode(getCleanFormHeader(loadTextFile(tempFilePath & USERFORM_CODE_EXT)))
    frmChanged = storedCode <> currentCode
    
    ' Persist changed elements
    If frxChanged Then
        Debug.Print "exporting " & component.name & USERFORM_XOLE_EXT
        DeleteFile storedFilePath & USERFORM_XOLE_EXT
        FSO.MoveFile tempFilePath & USERFORM_XOLE_EXT, storedFilePath & USERFORM_XOLE_EXT
        Debug.Print "exporting " & component.name & USERFORM_DATA_EXT
        saveTextFile storedFilePath & USERFORM_DATA_EXT, currentData
    End If
    If frmChanged Then
        Debug.Print "exporting " & component.name & USERFORM_CODE_EXT
        saveTextFile storedFilePath & USERFORM_CODE_EXT, currentCode
    End If
    
    ' Clean temporary files
    On Error Resume Next
    FSO.DeleteFile tempFilePath & ".*", True
    FSO.DeleteFolder tempFolder, True
    On Error GoTo 0
End Sub

Private Sub DeleteFile(ByVal fileName As String)
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(fileName) Then
        FSO.DeleteFile fileName
    End If
End Sub

Private Function loadMSFormData(ByVal exportPath As String, ByVal component As VBComponent) As String
    loadMSFormData = loadTextFile(getMSFormFileName(exportPath, component))
End Function

Public Function loadTextFile(ByVal fileName As String) As String
    Dim FSO As New Scripting.FileSystemObject
    Dim inStream As TextStream
    
    ' Check if data file does exist
    If Not FSO.FileExists(fileName) Then
        Debug.Print "loadTextFile skipped because " & fileName & " does not exist"
        Exit Function
    End If
    
    ' Read data file contents
    Set inStream = FSO.OpenTextFile(fileName, ForReading, False)
    loadTextFile = inStream.ReadAll
    inStream.Close
End Function

Public Sub saveTextFile(ByVal fileName As String, ByVal text As String)
    Dim FSO As New Scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = FSO.CreateTextFile(fileName, True, False)
    outStream.Write text
    outStream.Close
End Sub

Private Function getMSFormFileName(ByVal exportPath As String, ByVal component As VBComponent) As String
    getMSFormFileName = exportPath & "\" & component.name & USERFORM_DATA_EXT
End Function

Private Function isCodeIdentical(ByVal component As VBComponent, ByVal otherVersion As String) As Boolean
    Dim compVersion As String
    compVersion = getComponentCode(component)
    isCodeIdentical = getCleanCode(compVersion) = getCleanCode(otherVersion)
End Function

Private Function getCleanCode(ByVal code As String) As String
    getCleanCode = RemoveTrailingEmptyLines(RemoveLeadingEmptyLines(code))
End Function

Private Function getComponentCode(ByVal component As VBComponent) As String
    getComponentCode = component.codeModule.lines(1, component.codeModule.CountOfLines)
End Function

Public Function RemoveLeadingEmptyLines(ByVal text As String) As String
    Do
        text = LTrim(text)
        If Left(text, 2) = vbCrLf Then
            text = Mid(text, 3)
        Else
            RemoveLeadingEmptyLines = text
            Exit Function
        End If
    Loop
End Function

Public Function RemoveTrailingEmptyLines(ByVal text As String) As String
    Do
        text = LTrim(text)
        If Right(text, 2) = vbCrLf Then
            text = Left(text, Len(text) - 2)
        Else
            RemoveTrailingEmptyLines = text & vbCrLf
            Exit Function
        End If
    Loop
End Function

Public Function getCleanFormHeader(ByVal userFormCode As String) As String
    Dim lns
    Dim i As Long
    Dim startLn As Long
    Dim removeLns As Long
    Dim seenAttribute As Boolean
    Dim inCode As Boolean
    lns = Split(userFormCode, vbCrLf)
    For i = LBound(lns) To UBound(lns)
        ' Found end of header ?
        If Not seenAttribute Then
            If InStr(lns(i), "Attribute") = 1 Then
                seenAttribute = True
            End If
        ElseIf startLn = 0 Then
            If InStr(lns(i), "Attribute") <> 1 Then
                startLn = i - 1
            End If
        End If
        If startLn > 0 And Not inCode Then
            If Trim(lns(i)) = "" Then
                removeLns = removeLns + 1
            Else
                If removeLns = 0 Then
                    getCleanFormHeader = userFormCode
                    Exit Function
                End If
                inCode = True
            End If
        End If
        If inCode Then
            lns(i - removeLns) = lns(i)
        End If
    Next i
    ReDim Preserve lns(UBound(lns) - removeLns)
    getCleanFormHeader = Join(lns, vbCrLf)
End Function


''
' Join Path with \
'
' @example
' ```VB.net
' Debug.Print JoinPath("a/", "/b")
' Debug.Print JoinPath("a", "b")
' Debug.Print JoinPath("a/", "b")
' Debug.Print JoinPath("a", "/b")
' -> a/b
' ```
'
' @param {String} LeftSide
' @param {String} RightSide
' @return {String} Joined path
''
Public Function JoinPath(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "\" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "\" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If

    If LeftSide <> "" And RightSide <> "" Then
        JoinPath = LeftSide & "\" & RightSide
    Else
        JoinPath = LeftSide & RightSide
    End If
End Function


