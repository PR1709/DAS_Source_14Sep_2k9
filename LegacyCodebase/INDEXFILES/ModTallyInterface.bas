Attribute VB_Name = "ModTallyInterface"
Option Explicit
Public rtb As RichTextBox

''''''Private Declare Function OpenTally Lib "RTSLINK.DLL" Alias "Open" () As Long
''''''Private Declare Function SendRequest Lib "RTSLINK.DLL" (ByVal strRequestXML As String) As Long
''''''Private Declare Function GetLastErrorMessage Lib "RTSLINK.DLL" () As String

Private oRTSlink As New RTSlink
Public strURI As String

Public Function ConnectToTally() As Integer
    Dim tallyRunning As Integer
    tallyRunning = oRTSlink.Send("", strURI)
    ConnectToTally = tallyRunning
End Function

Public Function SendData(xmlString As String) As Long
    Dim SendStatus As Long
    Dim STR As String
    Dim ERRMSG As String
    SendStatus = oRTSlink.Send(xmlString, strURI)
    STR = oRTSlink.GetResponse()
    If SendStatus = 0 Then
        
        ERRMSG = STR
'        If InStr(1, STR, "<ERROR>") Then
            STR = Mid(STR, InStr(1, STR, "<ERRORS>"), InStr(9, STR, "</ERRORS>") - 9)
'        End If
'        SendStatus = CLng(Mid(STR, InStr(1, STR, ">") + 1, InStr(9, STR, "</ERRORS>") - 9))
    End If
    SendData = SendStatus
    If SendStatus <> 0 Then
        On Error GoTo ErrHandler
        ERRMSG = Mid(ERRMSG, InStr(1, ERRMSG, "<LINEERROR>"), InStr(12, ERRMSG, "</LINEERROR>") - 12)
        ERRMSG = Mid(ERRMSG, InStr(1, ERRMSG, ">") + 1, InStr(12, ERRMSG, "</LINEERROR>") - 12)
        rtb.Text = rtb.Text & ERRMSG & vbCrLf
        rtb.SelStart = Len(rtb.Text)
    End If
ErrHandler:
    Exit Function
End Function

Public Function GetData(xmlString As String) As String
    Dim SendStatus As Integer
    SendStatus = oRTSlink.Send(xmlString, strURI)
    GetData = oRTSlink.GetResponse()
End Function


