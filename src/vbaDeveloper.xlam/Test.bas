Attribute VB_Name = "Test"

Option Explicit

Private Type myOwn
    name As String
    age As Integer
    car As Variant
End Type

Public Sub testMyCustomActions_Open()
    Dim myCustomAction As CustomActions
    Set myCustomAction = New MyCustomActions
    myCustomAction.afterOpen
End Sub


Public Sub testImport()
    Dim proj_name As String
    proj_name = "vbaDeveloper"
    
    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode vbaProject
End Sub


Public Sub testExport()
    Dim proj_name As String
    proj_name = "vbaDeveloper"
    
    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects(proj_name)
    Build.exportVbaCode vbaProject
End Sub


' Now we add some code to try out all the types of formatting
' this is to test the Formatter module

Private Property Get wbaWords() As Dictionary
    Set wbaWords = New Dictionary
End Property

Public Property Let meSleep(ByVal s As String)
    s = "hello"
End Property

Property Get vaWords() As Dictionary
    Set vaWords = wbaWords
End Property


Property Let vaWords(x As Dictionary)
    Dim y As Object
    Set y = x
End Property

Private Sub anotherPrivateSub()
    anotherPublicFunction
    Dim y As Integer
    y = 4
    Do Until y = 0
        y = y - 1
    Loop
    y = 5
End Sub

Public Function anotherPublicFunction() As String
    ' Lets do a for loop
    Dim myCollection As Collection
    Dim x
    For Each x In myCollection
        Debug.Print x
        Dim thisMethod, doesnt, matter, dont, thiscode
        x.doesNotHave thisMethod
        If 2 Then
            x.butThat doesnt, matter
        Else
            If False Then
                'we don't do anything here
            ElseIf True Then
                becauseWe dont.Run, thiscode
            Else
                If x > 0 Then
                    x = 0
                ElseIf x > -5 Then
                    x = -5
                Else
                    Debug.Print "x is less than -5"
                End If
            End If
        End If
        Debug.Print "we should not forget the indentation for nested stuff"
    Next x
End Function

Private Function becauseWe(x, y) As Variant
    On Error GoTo jail
    'now we do an indexed for loop
    Dim i As Integer
    For i = 1 To 5
        Debug.Print i
        If True Then
        Else
            'there was only false
        End If
    Next
jail:
    MsgBox "Error occurred!", , "you are now in jail"
End Function

Function withoutAccessModifier()
    ' and a do while loop
    Dim y As Integer
    Dim finished As Boolean
    finished = False
    Do While Not finished
        y = y + 1
        If y = 10 Then
            finished = True
        End If
    Loop
End Function

Sub aSubWithoutAccessModifier()
    Dim p As Object
somelabel:
    With p
        .codeIsNotSupposedToReachHere
    End With
anotherLabel:
    
End Sub

' some more comments
' end this is the last line
