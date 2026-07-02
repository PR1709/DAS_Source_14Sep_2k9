Attribute VB_Name = "modInnComm"


'Public CommPort As Integer

Public BaudRate As Long
Public Parity As String
Public DataBit As Integer
Public StopBit As Integer
'Public ErrorLog() As String

Private SMSMessage() As String
Private st As Byte


Public Function WriteLogFile(strText As String, MessageType As Integer, ByRef strFile As String) As Long
    On err GoTo ErrHandler
    Dim FSO As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    If MessageType = 0 Then
        fileName = LogFolder & "\SMSSendLog.txt"
    ElseIf MessageType = 1 Then
        fileName = LogFolder & "\SMSReceiveLog.txt"
    ElseIf MessageType = 2 Then
        fileName = RootPath & "\SMSDeleteLog.txt"
    End If
    If FSO.FileExists(fileName) Then
        Set f = FSO.GetFile(fileName)
        Set ts = f.OpenAsTextStream(ForAppending)
    Else
        Set ts = FSO.CreateTextFile(fileName, True, False)
    End If
    ts.Write strText
    ts.Close
    strFile = fileName
    WriteLogFile = 0
    Exit Function
ErrHandler:
    WriteLogFile = err.Number
End Function

Public Function ReadLogFile(MessageType As String, ByRef strText As String) As Long
On err GoTo ErrHandler
    Dim FSO As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    If MessageType = 0 Then
        fileName = LogFolder & "\SMSSendLog.txt"
    ElseIf MessageType = 1 Then
        fileName = LogFolder & "\SMSReceiveLog.txt"
    ElseIf MessageType = 2 Then
        fileName = LogFolder & "\SMSDeleteLog.txt"
    End If
    If FSO.FileExists(fileName) Then
        Set ts = FSO.OpenTextFile(fileName, ForReading)
        If Not ts.AtEndOfLine Then strText = ts.ReadAll
        ts.Close
    Else
        strText = ""
    End If
    ReadLogFile = 0
    Exit Function
ErrHandler:
    ReadLogFile = err.Number
End Function

Public Function SendMessage(CommPort As Integer, innComm As MSComm, PhoneNo As String, Message As String, SetTimeOut As Integer, rtb As RichTextBox) As String
    If CommPort <= 0 Then
        SendMessage = "Comm Port Not Specified"
        Exit Function
    End If
On Error GoTo ErrHandler
    Dim ATCommand As String
    Dim strResult As String
    Dim keyArray
    Dim msg As String
    Dim STR As String
    
    BreakMessage Message
    innComm.CommPort = CommPort
    innComm.settings = settings()
    innComm.InputLen = 0
    innComm.Handshaking = comRTS
    innComm.PortOpen = True
    For I = 0 To UBound(SMSMessage) - 1
        ATCommand = ""
        ATCommand = "AT+CMGS=" & """" & PhoneNo & """" & Chr$(13) & SMSMessage(I) & Chr$(26)
        innComm.Output = ATCommand
        status = True
        Wait SetTimeOut
''''''''        ATCommand = ""
''''''''        ATCommand = SMSMessage(I) & Chr$(26)
''''''''        innComm.Output = ATCommand
''''''''        Wait SetTimeOut
        STR = innComm.Input
        
        STR = Replace(STR, "AT+CMGS=", "")
        STR = Replace(STR, Chr(13), "")
        STR = Replace(STR, Chr(13) & Chr(10), "")
        STR = Replace(STR, Chr(10), "")
        STR = Replace(STR, Chr(13) & Chr(10), "")
        STR = Replace(STR, Chr(12), "")
        STR = Replace(STR, Chr(26), "")
        STR = Replace(STR, "+CMGS: ", "-")
        STR = Replace(STR, "OK", "-OK")
        STR = Replace(STR, "> ", "-")
        STR = Replace(STR, """", "")
        STR = STR & vbCrLf
'        rtb.Text = rtb.Text & mid( STR,1
        DoEvents
        innComm.PortOpen = False
    Next
    SendMessage = 0
    Exit Function
ErrHandler:
    SendMessage = GetErrorMessage(err)
    If innComm.PortOpen = True Then innComm.PortOpen = False
    AddPSSErrorLog "modInnComm - SendMessage"
End Function

Public Function ReceiveMessage(CommPort As Integer, innComm As MSComm, SMSListType As Integer, SetTimeOut As Integer, rtb As RichTextBox) As String
    If CommPort <= 0 Then
        'ReceiveMessage = "Comm Port Not Specified"
        Exit Function
    End If
    On Error GoTo ErrHandler
    Dim ATCommand As String
    Dim STR As String
    
    
    Dim PhoneNo As String
    Dim msg As String
    
    innComm.CommPort = CommPort
    innComm.settings = settings()
    innComm.InputLen = 0
    innComm.Handshaking = comRTS
'    innComm.CTSHolding = True
'    innComm.DSRHolding = True
'    innComm.DTREnable = True
    'innComm.InputMode = comInputModeText
    
    innComm.PortOpen = True
    
    Select Case SMSListType
    Case 0
        ATCommand = "AT+CMGL=" & """" & "REC UNREAD" & """" & Chr$(13)
    Case 1
        ATCommand = "AT+CMGL=" & """" & "REC READ" & """" & Chr$(13)
    Case 2
        ATCommand = "AT+CMGL=" & """" & "STO UNSENT" & """" & Chr$(13)
    Case 3
        ATCommand = "AT+CMGL=" & """" & "STO SENT" & """" & Chr$(13)
    Case 4
        ATCommand = "AT+CMGL=" & """" & "ALL" & """" & Chr$(13)
    End Select
    innComm.Output = ATCommand
    Wait SetTimeOut
    STR = innComm.Input
    STR = Replace(STR, "AT+CMGL=" & """" & "ALL" & """", "")
    STR = Replace(STR, Chr(13) & Chr(10), ",")
    STR = Replace(STR, Chr(10), ",")
    STR = Replace(STR, Chr(12), ",")
    STR = Replace(STR, Chr(26), ",")
    STR = Replace(STR, Chr(4), "")
    STR = Replace(STR, Chr(1), "")
    STR = Replace(STR, """", "")
    STR = Replace(STR, "+CMGL: ", vbCrLf)
    rtb.Text = rtb.Text & STR

    ReceiveMessage = STR
    innComm.PortOpen = False
    Exit Function
ErrHandler:
   ' ReceiveMessage = GetErrorMessage(err.Number)
    If innComm.PortOpen Then innComm.PortOpen = False
    AddPSSErrorLog "modInnComm - ReceiveMessage"
End Function


Public Function DeleteMessage(CommPort As Integer, innComm As MSComm, SMSListType As Integer)
    If CommPort <= 0 Then
        DeleteMessage = "Comm Port Not Specified"
        Exit Function
    End If
On Error GoTo ErrHandler
    Dim ATCommand As String
    Dim strResult As String
    
    
    Dim PhoneNo As String
    Dim msg As String
    
    
    innComm.CommPort = CommPort
    innComm.settings = settings()
    innComm.InputLen = 0
    innComm.PortOpen = True
    
    
    Select Case SMSListType
    Case 0
        ATCommand = "AT+CMGD=0"
    Case 1
        ATCommand = "AT+CMGD=1"
    Case 2
        ATCommand = "AT+CMGD=2"
    Case 3
        ATCommand = "AT+CMGD=3"
    Case 4
        ATCommand = "AT+CMGD=4"
    End Select
    
    innComm.Output = ATCommand
    strResult = innComm.Input
    innComm.PortOpen = False
    Exit Function
ErrHandler:
If innComm.PortOpen = True Then innComm.PortOpen = False
    AddPSSErrorLog "ModInnComm - DeleteMessage"
End Function

Public Function CommSettings() As String()
On err GoTo ErrHandler
    Dim FSO As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    Dim strText() As String
    fileName = SetupFolder & "\ComSettings.txt"
    If FSO.FileExists(fileName) Then
        Set ts = FSO.OpenTextFile(fileName, ForReading)
        strText = Split(ts.ReadAll, ",")
    End If
    BaudRate = strText(0)
    Parity = strText(1)
    DataBit = strText(2)
    StopBit = strText(3)
    ts.Close
    CommSettings = strText
    Exit Function
ErrHandler:
    ReDim strText(1)
    CommSettings = strText
End Function


Public Function SaveCommSettings(BaudRate As Long, Parity As String, DataBit As Integer, StopBit As Integer, CommPortNo As Integer, Wait As Integer) As Long
On err GoTo ErrHandler
    Dim FSO As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    fileName = SetupFolder & "\ComSettings.txt"
    If FSO.FileExists(fileName) Then
        Set ts = FSO.OpenTextFile(fileName, ForWriting)
        ts.Write BaudRate & "," & Parity & "," & DataBit & "," & StopBit & "," & Wait & "," & CommPortNo & ","
    End If
    ts.Close
    CommSettings
    SaveCommSettings = 0
    Exit Function
ErrHandler:
    SaveCommSettings = err.Number
End Function

Private Function settings() As String
    Parity = UCase(Trim(Parity))
    If BaudRate = 0 Then BaudRate = 9600
    If Parity = "" Then Parity = "N"
    If DataBit <= 0 Then DataBit = 8
    If StopBit <= 0 Then StopBit = 1
    
    settings = BaudRate & "," & Parity & "," & DataBit & "," & StopBit

End Function

Private Sub BreakMessage(stringMessage As String)
    Dim temp As String
    Dim bool As Boolean
    ReDim Preserve SMSMessage(1)
    bool = True
    While bool
            temp = Mid(stringMessage, 1, 160)
            SMSMessage(UBound(SMSMessage) - 1) = temp
            stringMessage = Mid(stringMessage, 161)
            If Len(stringMessage) <= 0 Then
            bool = False
        Else
            ReDim Preserve SMSMessage(UBound(SMSMessage) + 1)
            SMSMessage(UBound(SMSMessage) - 1) = stringMessage
        End If
    Wend
End Sub

Private Function GetErrorMessage(err As ErrObject) As String
    Select Case err.Number
        Case comEventBreak
            GetErrorMessage = "A Break was received."
        Case comEventFrame
            GetErrorMessage = "Framing Error."
        Case comEventOverrun
            GetErrorMessage = "Data Lost."
        Case comEventRxOver
            GetErrorMessage = "Receive buffer overflow."
        Case comEventRxParity
            GetErrorMessage = "Parity Error."
        Case comEventTxFull
            GetErrorMessage = "Transmit buffer full."
        Case comEventDCB
            GetErrorMessage = "Unexpected error retrieving DCB]"
        Case Default
            GetErrorMessage = err.Description
    End Select
End Function

Private Sub Wait(SetTimeOut As Integer)
    DoEvents
    Dim sec1
    Dim status As Boolean
    status = False
    sec1 = Time$
    While Not status
        If Abs(Second(Time$) - Second(sec1)) >= SetTimeOut Then
        status = True
        End If
    Wend
End Sub


