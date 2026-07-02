Attribute VB_Name = "modMessageTemplate"

Public Function MessageTemplate(ByRef ArrayTemplate) As Long
    On err GoTo ErrHandler
    Dim fso As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    fileName = App.Path & "\SMSMessageTemplate.txt"
    If fso.FileExists(fileName) Then
        Set f = fso.GetFile(fileName)
        Set ts = f.OpenAsTextStream(ForReading)
    End If
    ts.ReadLine
    ts.ReadLine
    ts.ReadLine
    While ts.AtEndOfLine <> True
        ReDim Preserve ArrayTemplate(UBound(ArrayTemplate) + 1)
        ArrayTemplate(UBound(ArrayTemplate) - 1) = ts.ReadLine()
    Wend
    ts.Close
    MessageTemplate = 0
    Exit Function
ErrHandler:
    MessageTemplate = err.Number
End Function

Public Function SaveTemplate(MsgName As String, strMessageTemplate As String) As Long
    On err GoTo ErrHandler
    Dim fso As New FileSystemObject
    Dim fileName As String
    Dim ts As TextStream
    fileName = App.Path & "\" & Trim(MsgName) & ".txt"
    Set ts = fso.CreateTextFile(fileName, True, False)
    ts.Write strMessageTemplate
    ts.Close
    SaveTemplate = 0
    Exit Function
ErrHandler:
    SaveTemplate = err.Number
End Function


Public Function OpenTemplate(MsgName As String, ByRef strMessageTemplate As String) As Long
    On err GoTo ErrHandler
    Dim fso As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    fileName = App.Path & "\" & Trim(MsgName) & ".txt"
    If fso.FileExists(fileName) Then
        Set f = fso.GetFile(fileName)
        Set ts = f.OpenAsTextStream(ForReading)
        strMessageTemplate = ts.ReadAll
        ts.Close
    End If
    OpenTemplate = 0
    Exit Function
ErrHandler:
    OpenTemplate = err.Number

End Function


