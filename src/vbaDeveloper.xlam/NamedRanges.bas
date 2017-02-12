Attribute VB_Name = "NamedRanges"
Option Explicit

Private Const NAMED_RANGES_FILE_NAME As String = "NamedRanges.csv"

Private Enum columns
    name = 0
    RefersTo
    Comments
End Enum


' Import named ranges from csv file
' Existing ranges with the same identifier will be replaced.
Public Sub importNamedRanges(wb As Workbook)
    Dim importDir As String
    importDir = Build.getSourceDir(wb.FullName, createIfNotExists:=False)
    If importDir = "" Then
        Debug.Print "No import directory for workbook " & wb.name & ", skipping"
        Exit Sub
    End If

    Dim fileName As String
    fileName = importDir & NAMED_RANGES_FILE_NAME
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(fileName) Then
        Dim inStream As TextStream
        Set inStream = FSO.OpenTextFile(fileName, ForReading, Create:=False)
        Dim line As String
        Do Until inStream.AtEndOfStream
            line = inStream.ReadLine
            importName wb, line
        Loop
        inStream.Close
    End If
End Sub


Private Sub importName(wb As Workbook, line As String)
    Dim parts As Variant
    parts = Split(line, ",")
    Dim rangeName As String, rangeAddress As String, comment As String
    rangeName = parts(columns.name)
    rangeAddress = parts(columns.RefersTo)
    comment = parts(columns.Comments)

    ' Existing namedRanges don't need to be removed first.
    ' wb.Names.Add will automatically replace or add the given namedRange.
    wb.Names.Add(rangeName, rangeAddress).comment = comment
End Sub


'Export named ranges to csv file
Public Sub exportNamedRanges(wb As Workbook)
    Dim exportDir As String
    exportDir = Build.getSourceDir(wb.FullName, createIfNotExists:=True)
    Dim fileName As String
    fileName = exportDir & NAMED_RANGES_FILE_NAME

    Dim lines As Collection
    Set lines = New Collection
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If hasValidRange(aName) Then
            lines.Add aName.name & "," & aName.RefersTo & "," & aName.comment
        End If
    Next
    If lines.Count > 0 Then
        'We have some names to export
        Debug.Print "writing to  " & fileName

        Dim FSO As New Scripting.FileSystemObject
        Dim outStream As TextStream
        Set outStream = FSO.CreateTextFile(fileName, overwrite:=True, unicode:=False)
        On Error GoTo closeStream
        Dim line As Variant
        For Each line In lines
            outStream.WriteLine line
        Next line
closeStream:
        outStream.Close
    End If
End Sub


Private Function hasValidRange(aName As name) As Boolean
    On Error GoTo no
    hasValidRange = False
    Dim aRange As Range
    Set aRange = aName.RefersToRange
    hasValidRange = True
no:
End Function


' Clean up all named ranges that don't refer to a valid range.
' This sub is not used by the import and export functions.
' It is provided only for convenience and can be run manually.
Public Sub removeInvalidNamedRanges(wb As Workbook)
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If Not hasValidRange(aName) Then
            aName.Delete
        End If
    Next
End Sub
