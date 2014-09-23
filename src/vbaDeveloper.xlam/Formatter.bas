Attribute VB_Name = "Formatter"
Option Explicit

Private Const BEG_SUB = "Sub "
Private Const END_SUB = "End Sub"
Private Const BEG_PB_SUB = "Public Sub "
Private Const BEG_PV_SUB = "Private Sub "

Private Const BEG_FUN = "Function "
Private Const END_FUN = "End Function"
Private Const BEG_PB_FUN = "Public Function "
Private Const BEG_PV_FUN = "Private Function "

Private Const BEG_PROP = "Property "
Private Const END_PROP = "End Property"
Private Const BEG_PB_PROP = "Public Property "
Private Const BEG_PV_PROP = "Private Property "

Private Const BEG_ENUM = "Enum "
Private Const END_ENUM = "End Enum"
Private Const BEG_PB_ENUM = "Public Enum "
Private Const BEG_PV_ENUM = "Private Enum "

Private Const BEG_IF = "If "
Private Const END_IF = "End If"
Private Const BEG_WITH = "With "
Private Const END_WITH = "End With"

Private Const BEG_SELECT = "Select "
Private Const END_SELECT = "End Select"

Private Const BEG_FOR = "For "
Private Const END_FOR = "Next "
Private Const BEG_DOWHILE = "Do While "
Private Const BEG_DOUNTIL = "Do Until "

Private Const BEG_TYPE = "Type "
Private Const END_TYPE = "End Type"
Private Const BEG_PB_TYPE = "Public Type "
Private Const BEG_PV_TYPE = "Private Type "

' Single words that need to be handled separately
Private Const ONEWORD_END_FOR = "Next"
Private Const ONEWORD_END_LOOP = "Loop"
Private Const ONEWORD_ELSE = "Else"
Private Const BEG_END_ELSEIF = "ElseIf"
Private Const BEG_END_CASE = "Case "


Private Const INDENT = "    "

Private words As Dictionary 'Keys are Strings, Value is an Integer indicating change in indentation
Private indentation(0 To 20) As Variant ' Prevent repeatedly building the same strings by looking them up in here

Private Sub initialize()
    initializeWords
    initializeIndentation
End Sub

Private Sub initializeIndentation()
    Dim indentString As String
    indentString = ""
    Dim i As Integer
    For i = 0 To UBound(indentation)
        indentation(i) = indentString
        indentString = indentString & INDENT
    Next
End Sub

Private Sub initializeWords()
    Dim w As Dictionary
    Set w = New Dictionary

    w.Add BEG_SUB, 1
    w.Add END_SUB, -1
    w.Add BEG_PB_SUB, 1
    w.Add BEG_PV_SUB, 1

    w.Add BEG_FUN, 1
    w.Add END_FUN, -1
    w.Add BEG_PB_FUN, 1
    w.Add BEG_PV_FUN, 1

    w.Add BEG_PROP, 1
    w.Add END_PROP, -1
    w.Add BEG_PB_PROP, 1
    w.Add BEG_PV_PROP, 1

    w.Add BEG_ENUM, 1
    w.Add END_ENUM, -1
    w.Add BEG_PB_ENUM, 1
    w.Add BEG_PV_ENUM, 1

    w.Add BEG_IF, 1
    w.Add END_IF, -1
    w.Add BEG_SELECT, 2 'because any following 'Case' indents to the left
    w.Add END_SELECT, -2
    w.Add BEG_WITH, 1
    w.Add END_WITH, -1

    w.Add BEG_FOR, 1
    w.Add END_FOR, -1
    w.Add BEG_DOWHILE, 1
    w.Add BEG_DOUNTIL, 1

    w.Add BEG_TYPE, 1
    w.Add END_TYPE, -1
    w.Add BEG_PB_TYPE, 1
    w.Add BEG_PV_TYPE, 1

    Set words = w
End Sub


Private Property Get vbaWords() As Dictionary
    If words Is Nothing Then
        initialize
    End If
    Set vbaWords = words
End Property

Public Sub testFormatting()
    If words Is Nothing Then
        initialize
    End If
    'Debug.Print Application.VBE.ActiveCodePane.codePane.Parent.Name
    'Debug.Print Application.VBE.ActiveWindow.caption
    
    Dim projName As String, moduleName As String
    projName = "vbaDeveloper"
    moduleName = "Test"
    Dim vbaProject As VBProject
    Set vbaProject = Application.VBE.VBProjects(projName)
    Dim code As codeModule
    Set code = vbaProject.VBComponents(moduleName).codeModule

    'removeIndentation code
    'formatCode code
    formatProject vbaProject
End Sub

Public Sub formatProject(vbaProject As VBProject)
    Dim codePane As codeModule
    
    Dim code As Variant
    For Each code In vbaProject.VBComponents.Items
        Set codePane = code
        Debug.Print "Formatting " & codePane.Parent.name
        formatCode codePane
    Next
End Sub

Public Sub format()
    formatCode Application.VBE.ActiveCodePane.codeModule
End Sub


Public Sub formatCode(codePane As codeModule)
    On Error GoTo formatCodeError
    Dim lineCount As Integer
    lineCount = codePane.CountOfLines

    Dim indentLevel As Integer, nextLevel As Integer, levelChange As Integer
    indentLevel = 0
    Dim lineNr As Integer
    For lineNr = 1 To lineCount
        Dim line As String
        line = Trim(codePane.Lines(lineNr, 1))
        If Not line = "" Then
            If isEqual(ONEWORD_ELSE, line) _
                Or lineStartsWith(BEG_END_ELSEIF, line) _
                Or lineStartsWith(BEG_END_CASE, line) Then
                levelChange = 1
                indentLevel = -1 + indentLevel
            ElseIf isLabel(line) Then
                levelChange = indentLevel
                indentLevel = 0
            Else
                levelChange = indentChange(line)
            End If

            nextLevel = indentLevel + levelChange
            If levelChange <= -1 Then
                indentLevel = nextLevel
            End If

            line = indentation(indentLevel) + line
            indentLevel = nextLevel
        End If
        Call codePane.ReplaceLine(lineNr, line)
    Next
    Exit Sub
formatCodeError:
    Debug.Print "Error while formatting " & codePane.Parent.name
    Debug.Print Err.Number & " " & Err.Description
    Debug.Print " on line " & lineNr & ": " & line
    Debug.Print "indentLevel: " & indentLevel & " , levelChange: " & levelChange
End Sub


Public Sub removeIndentation(codePane As codeModule)
    Dim lineCount As Integer
    lineCount = codePane.CountOfLines

    Dim lineNr As Integer
    For lineNr = 1 To lineCount
        Dim line As String
        line = codePane.Lines(lineNr, 1)
        line = Trim(line)
        Call codePane.ReplaceLine(lineNr, line)
    Next
End Sub

Private Function indentChange(ByVal line As String) As Integer
    indentChange = 0
    Dim w As Dictionary
    Set w = vbaWords

    If isEqual(line, ONEWORD_END_FOR) Or _
        isEqual(line, ONEWORD_END_LOOP) Then
        indentChange = -1
        GoTo hell
    End If
    Dim word As String
    Dim vord As Variant
    For Each vord In w.Keys
        word = vord 'Cast the Variant to a String
        If lineStartsWith(word, line) Then
            indentChange = vbaWords(word)
            GoTo hell
        End If
    Next
hell:
End Function

' Returns true if both strings are equal, ignoring case
Private Function isEqual(first As String, second As String) As Boolean
    isEqual = (StrComp(first, second, vbTextCompare) = 0)
End Function

' Returns True if strToCheck begins with begin, ignoring case
Private Function lineStartsWith(begin As String, strToCheck As String) As Boolean
    lineStartsWith = False
    Dim beginLength As Integer
    beginLength = Len(begin)
    If Len(strToCheck) >= beginLength Then
        lineStartsWith = isEqual(begin, left(strToCheck, beginLength))
    End If
End Function


Private Function isLabel(line As String) As Boolean
    isLabel = (right(line, 1) = ":")
End Function
