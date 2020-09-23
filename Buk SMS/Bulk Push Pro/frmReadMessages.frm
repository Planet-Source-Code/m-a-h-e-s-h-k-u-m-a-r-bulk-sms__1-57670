VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmReadMessages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anto Messaging"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmReadMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5100
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDefaultMessage 
      Caption         =   "Default Reply "
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1725
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1155
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Message"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      Height          =   1095
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start "
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer tmrCheckMessage 
      Left            =   600
      Top             =   5520
   End
   Begin VB.TextBox txtMobilenumber 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   "+919820615407"
      Top             =   5880
      Width           =   3015
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   5160
      Width           =   3015
   End
   Begin VB.ListBox lstEvents 
      Height          =   5325
      Left            =   6480
      TabIndex        =   2
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fetch Message from Server"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtTelephone 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   4680
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParityReplace   =   32
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   4800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Reply Server "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "frmReadMessages.frx":0442
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Labelkeyword 
         Height          =   735
         Left            =   3360
         TabIndex        =   18
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label labelfrom 
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Keyword:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Recd. Message :"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Send Message :"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   5640
      Width           =   2415
   End
End
Attribute VB_Name = "frmReadMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEcho                   As Boolean
Public bOK                     As Boolean
Public bRing                   As Boolean
Public bError                  As Boolean
Public iRingTime               As Single
Public FirstRun                As Boolean
Public bErrorComm              As Boolean
Public bGreaterSign            As Boolean
Public bMessageStore           As Boolean
Public strMessageBuffer        As String
Public FileNumber              As Integer
Dim msgBreak()                 As String
Dim msgHeader()                As String
Dim MessageStatus As String
Dim time_stamp As String
Dim Message As String
Dim Mobile_Number As String
Dim MessageCount As String
Dim DefaultMessage As String
Public Baudrate, Set_parity, Brate, Com_count, fnam1
Dim AT_flag, AT_flag1, CMGF_flag, Data_check, Voice_check, SMS_check, Inet_check, config_check As Boolean
Dim portcheck As String
Public ATFLAG As String
Private Message_sent As Integer


Private Sub comm1_OnComm()
    Static stEvent             As String
    Dim stComChar               As String * 1


    Select Case Comm1.CommEvent

        Case comEvReceive

            Do
                stComChar = Comm1.Input
               If bMessageStore Then
                   strMessageBuffer = strMessageBuffer & stComChar
                  
                End If
                'bMessageStore = False
                
'                If bMessageStore Then
'
'                    strMessageBuffer = Text1.Text
'
'                End If
                
                
                
                Select Case stComChar
                    Case ">"
                         bGreaterSign = True
                         lstEvents.AddItem stComChar
                    Case vbLf

                    Case vbCr
                        If Len(stEvent) > 0 Then
                          ProcessEvent stEvent
                          stEvent = ""
                        End If
                    Case Else
                        stEvent = stEvent + stComChar
                    
                End Select
      
            Loop While Comm1.InBufferCount
    Case 3
                 MsgBox "Modem Unplugged. Close the appplication and restart"
                 
    End Select

End Sub

Private Sub Command1_Click()
If Len(Trim(txtMobilenumber.Text)) = 0 Then
   MsgBox "Please Enter Mobile number before sending ! " & vbCr & "The format is +919820615407", vbInformation + vbOKOnly, "Improper Number"
   Exit Sub
Else
   bGreaterSign = False
   Comm1.Output = "AT+CMGS=" & Chr(34) & Trim(txtMobilenumber.Text) & Chr(34) & vbCrLf
   While Not bGreaterSign
      DoEvents
      Wait
   Wend
   If bGreaterSign Then
      Comm1.Output = Trim(txtSend.Text) & Chr(26) & vbCrLf
      bOK = False
      bError = False
      While Not bOK Or bError
          DoEvents
          Wait
      Wend
      If bOK Then
         MsgBox "Message Sent", vbInformation + vbOKOnly, "Sent"
      Else
         MsgBox "Message Not Sent", vbCritical + vbOKOnly, "Cannot Send"
      End If
   Else
      MsgBox "Message cannot be sent", vbCritical + vbOKOnly, "Cannot Send"
   End If
   txtSend.Text = ""
   txtMobilenumber.Text = ""
End If
End Sub

Private Sub cmdstart_Click()
cmdDefaultMessage.Enabled = False
'frmChild.Timer1.Enabled = False
Initialise_Modem
Call Modem_checking
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
If command3.Enabled = True Then Call Command3_Click
'Command2.Enabled = False
'ReadInbox
cmdStart.Enabled = False
cmdDefaultMessage.Enabled = False

Dim DialString$, FromModem$
If Comm1.PortOpen = True Then Comm1.PortOpen = False
If Comm1.PortOpen = False Then
   'Comm1.PortOpen = True
   Comm1.DTREnable = True
   Comm1.RTSEnable = True
   Comm1.RThreshold = 1
   Comm1.InputLen = 1
   Comm1.Settings = "9600, n, 8, 1"
   bOK = False
   bError = False
   Comm1.PortOpen = True
   Comm1.Output = "AT" + vbCrLf
    delay (1)
   End If
 Do
   If Comm1.InBufferCount Then FromModem$ = FromModem$ + Comm1.Input
             If InStr(FromModem$, "OK") Then
             'MsgBox "ready"
             GoTo finish
            End If
    
  Loop
   'MsgBox "Port Already Open !", vbCritical + vbOKOnly, "Error opening port"
finish:
'MsgBox day(
'If Comm1.PortOpen = True Then Comm1.PortOpen = False
'frmChild.Timer1.Enabled = True
StatusBar1.Panels.item(1).Text = "Auto Messaging Started..."
If Message_sent = 0 Then
'StatusBar1.Panels(1).Text = "Sim card error"

Exit Sub

End If
End Sub
Private Sub ProcessEvent(stEvent As String)
  Dim stNumber As String
  
       lstEvents.AddItem stEvent
       'stEvent = Text1.Text
        If Mid$(stEvent, 1, 5) = "+CMTI" Then
           txtTelephone.Text = ""
           txtMessage.Text = ""
           strMessageBuffer = ""
            labelfrom.Caption = Mobile_Number
            Labelkeyword.Caption = Message
            stEvent = ""
            Command3_Click
           SendAutoMessage Mobile_Number, Message
           bOK = False
           bError = False
           
           If Comm1.PortOpen = False Then Comm1.PortOpen = True
           Comm1.Output = "AT+CMGD=1,4" & vbCrLf
           While Not bOK Or bError
                 DoEvents
                 Wait
           Wend
           If bError Then
              MsgBox "Unable to delete"
           End If
           Exit Sub
        End If
        If InStr(stEvent, "RING") Then
            MsgBox "Incoming Call Alert"
        
        End If
        If Mid$(stEvent, 1, 5) = "+CSQ" Then
        MsgBox tempb
        End If
        Select Case stEvent
           Case "OK"
             bOK = True
           Case "ERROR"
             bError = True
           Case "RING"
             If bRing = False Then
               bRing = True
             End If
             iRingTime = Timer
           Case Else
             Select Case Left(stEvent, 4)
               Case "TIME"
               Case "DATE"
               Case "NMBR"
               Case "NAME"
             End Select
             
        End Select


End Sub

Private Sub ReadInbox()

If MSComm1.PortOpen = True Then MSComm1.PortOpen = False

delay (1)
Call comm1_settings
bOK = False
bError = False
If Comm1.PortOpen = False Then Comm1.PortOpen = True
Comm1.Output = "AT+CMGL=" & Chr(34) & "ALL" & Chr(34) & vbCrLf
'    Do
'
'    If Comm1.InBufferCount Then FromModem$ = FromModem$ + Comm1.Input
'                 If InStr(FromModem$, "OK") Then
'                 MsgBox "ready"
'                 GoTo finish
'                End If
'    Loop
'finish:
While Not bOK Or bError
  bMessageStore = True
  DoEvents
  Wait
Wend
If bOK Then
    
   ReadMessage
   'MsgBox txtMessage.Text
   'MsgBox txtTelephone.Text
   
'  pos1 = InStr(txtMessage.Text, txtTelephone.Text, "/", vbTextCompare)
   'MsgBox txtMessage
   If InStr(1, UCase(txtMessage.Text), "NOTEPAD", vbTextCompare) <> 0 Then
      'Call ExecuteCommand("NotePad.exe")
   ElseIf InStr(1, UCase(txtMessage.Text), "CALC", vbTextCompare) <> 0 Then
      Call ExecuteCommand("Calc.exe")
   End If
End If
If bError Then
   txtMessage.Text = "Bad Read"
End If
If Comm1.PortOpen = True Then Comm1.PortOpen = False
'frmChild.Timer1.Enabled = True
'Comm1.PortOpen = False
'    ProcessEvent "+CC"
    
End Sub

Private Sub Wait()
Dim start

   start = Timer
   Do While Timer < start + 8
      DoEvents
      If bOK Then
        Exit Sub
      End If
      If bError Then
        Exit Sub
      End If
   Loop
 'ProcessEvent "gh ""
End Sub
Private Sub WaitLong()
Dim start

   start = Timer
   Do While Timer < start + 36
      DoEvents
      If bOK Then
        Exit Sub
      End If
      If bError Then
        Exit Sub
      End If
   Loop
End Sub


Private Sub ReadMessage()
Dim splitCMTI
Dim CMGLbreak() As String
Dim CMGLheader() As String
Dim j
'strMessageBuffer = Text1.Text ''mahi ,ahidummy unread message
If ParseFile Then

        'Debug.Print strMessageBuffer
        
        
        strMessageBuffer = Mid(strMessageBuffer, 3, Len(strMessageBuffer))
             CMGLbreak = Split(strMessageBuffer, "+CMGL:", , vbTextCompare)
            CMGLheader = Split(CMGLbreak(0), ",", , vbTextCompare)
         On Error GoTo finish
    For j = 0 To 20
           'Debug.Print CMTIbreak(i)
          
            strMessageBuffer = CMGLbreak(j)
           
           msgBreak = Split(strMessageBuffer, vbCrLf, , vbTextCompare)
           msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
           
           
            msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
           txtTelephone.Text = Mid$(Right$(msgHeader(2), 11), 1, 10)
           strMessageBuffer = ""
         For i = 1 To UBound(msgBreak(), 1)
               If i = 2 Then
               Message = strMessageBuffer
               End If
               strMessageBuffer = strMessageBuffer & msgBreak(i) & vbCrLf
         '      Message = strMessageBuffer
        Next i
            MessageCount = Mid(msgBreak(0), 7, 3)
        If Mid(MessageCount, 1, 1) = Chr(34) Then
            MessageCount = Mid(MessageCount, 2, 2)
        End If
           MessageStatus = msgHeader(1)
           'MsgBox Len(msgHeader(1))
        If InStr(msgHeader(1), "UNREAD") Then
            MessageStatus = Mid(msgHeader(1), 2, Len(msgHeader(1)) - 2)
        Else
            'DeleteMessage MessageCount
        End If
           
           time_stamp = msgHeader(4) + "," + msgHeader(5)
           time_stamp = Mid(time_stamp, 2, Len(time_stamp) - 5)
           txtTelephone.Text = Mid$(Right$(msgHeader(2), 11), 1, 10)
'            If txtTelephone.Text = "" Then
'            StatusBar1.Panels(1).Text = " Waiting for Messages"
            'MsgBox "inbox empty"
'            Exit Sub
'
'        End If
            'MsgBox msgBreak(6)
            Mobile_Number = txtTelephone.Text
        Debug.Print txtTelephone
        Debug.Print Message
        Debug.Print time_stamp
'        Else
'           txtMessage.Text = "Unable to decode Message"
  Next j
End If

finish:
End Sub

Private Sub ExecuteCommand(FileToExecute As String)
On Error GoTo OpenError
Dim Lng As Long
Lng = Shell(FileToExecute, vbNormalFocus)
OpenError:
If Err.Number <> 0 Then
   MsgBox "Cannot Understand Message! ", vbOKOnly, "Help"
   Resume Next
End If
End Sub

Private Sub cmdStop_Click()
StatusBar1.Panels(1).Text = "Auto Answering Stopped"
cmdStart.Enabled = True
cmdDefaultMessage.Enabled = True
If Comm1.PortOpen = True Then Comm1.PortOpen = False
'frmChild.Timer1.Enabled = True
cmdDefaultMessage.Enabled = False
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmdexit_Click()
If Comm1.PortOpen = True Then Comm1.PortOpen = False
'frmChild.Timer1.Enabled = True
Unload Me
End Sub

Private Sub cmdDefaultMessage_Click()
StatusBar1.Panels(1).Text = "Enter the default reply"
DefaultMessage = InputBox("", "", "Default Message")
StatusBar1.Panels(1).Text = "Auto Answer Stopped"
End Sub



Private Sub Form_Load()

''''''''''LoadInbox 9849706959#, "hi this is test", "04/05/25,15:11:25+00"
'frmChild.Timer1.Enabled = False

bMessageStore = False
command3.Enabled = True
'frmChild.SetFocus
frmMain.Enabled = False
frmMain.Toolbar1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Comm1.PortOpen = True Then Comm1.PortOpen = False


frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True


End Sub

Private Sub labelfrom_Click()
'labelfrom.Caption = "hi"
End Sub

Private Sub Labelkeyword_Click()
'Labelkeyword.Caption = "heheh"
End Sub

Private Sub lstEvents_DblClick()
lstEvents.Clear
End Sub
Public Function ParseFile() As Boolean
'strMessageBuffer = Text1.Text
Dim FirstOffSet As Long
Dim SecondOffSet As Long
Dim strBuffer1 As String
Dim strBuffer2 As String
Dim strBuffer3 As String
strBuffer1 = strMessageBuffer

Debug.Print strMessageBuffer
FirstOffSet = InStr(1, strBuffer1, "+CMGL:", vbTextCompare)
SecondOffSet = InStr(1, strBuffer1, vbCrLf & "OK", vbTextCompare)
If FirstOffSet <> 0 And SecondOffSet > FirstOffSet Then
   i = FirstOffSet
   While i < SecondOffSet
    strBuffer2 = strBuffer2 & Mid$(strBuffer1, i, 1)
    i = i + 1
   Wend
   Debug.Print strBuffer2
   ParseFile = True
   strMessageBuffer = strBuffer2
   Exit Function
End If
ParseFile = False
End Function
Public Sub delay(dlytime As Variant)
Dim delay_timer
Dim dli
delay_timer = Timer
dli = 0
Do
dli = dli + 1
   If Timer > (delay_timer + dlytime) Then
      Exit Do
   End If
Loop
End Sub

Private Sub LoadInbox(ByVal Mobileno As String, ByVal Message As String, ByVal TimeStamp As String)
Dim strQuery As String
Dim strTempTime As String
Dim tempName As String
Dim timeBreak()                 As String
Dim timeHeader()                As String
Dim temp_day As String
Dim temp_month As String
Dim temp_year As String


    timeBreak = Split(TimeStamp, ",", , vbTextCompare)
    timeHeader = Split(timeBreak(0), ",", , vbTextCompare)
       
      temp_day = Mid$(timeBreak(0), 7, 2)
      temp_year = Mid$(timeBreak(0), 1, 2)
      temp_month = Mid$(timeBreak(0), 4, 2)
       strTempTime = Format(timeBreak(1), "am")
        TimeStamp = temp_month + "/" + temp_day + "/" + "20" + temp_year + " " + Mid$(timeBreak(1), 1, 8)
       
        StatusBar1.Panels(1).Text = "Loading Message to Database...."
   On Error GoTo OpenError
   
   
            
            Set rs = New ADODB.Recordset
                strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & Mobileno & "'"
            
            rs.Open strQuery, con, 3, 2, 1
            
                tempName = rs.Fields("CONTACTNAME")
            
OpenError:
If Err.Number <> 0 Then
  '  MsgBox Err.Number
   tempName = "Unknow"
   Resume Next
End If
       Set rs = Nothing
       
       Trim (Message)
            strQuery = "insert into INBOX(MOBILENO,MESSAGE,TIME_STAMP,NAME) values ('" & Mobileno & "', '" & Message & "', '" & TimeStamp & "','" & tempName & "')"
       Debug.Print strQuery
       'InputBox "", "", strQuery
      ' On Error GoTo handler
       con.Execute strQuery
handler:
'       DeleteMessage (MessageCount)
End Sub

Private Sub CheckData()

If Mobile_Number = "9849706959" Then
           SendMsgToAdmin
        End If
    If MessageStatus = "REC UNREAD" Then
        StatusBar1.Panels(1).Text = "Replying...."
        SendAutoMessage Mobile_Number, Message
      End If
    If frmSend.Message_sent = 0 Then
    StatusBar1.Panels(1).Text = "Reset the Application"
    End If
 End Sub

Private Sub SendAutoMessage(ByVal Mobileno As String, ByVal recieved_message As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
        labelfrom.Caption = ""
        Labelkeyword.Caption = ""
    Set rs = New ADODB.Recordset
    strQuery = "select * from AUTOMESSAGE"
    rs.Open strQuery, con, 3, 2, 1
   
    While Not rs.EOF
        
       If InStr(1, UCase(recieved_message), UCase(rs.Fields("KEYWORD")), vbTextCompare) Then
            Label3.Caption = "TO"
            Label4.Caption = "Message"
            labelfrom.Caption = Mobileno
            'MsgBox MobileNo
            'MsgBox rs.Fields("AUTOMESSAGE")
            Labelkeyword.Caption = rs.Fields("AUTOMESSAGE")
            'SEND MESSAGE TO MOBILE_NUMBER
            StatusBar1.Panels(1).Text = "Auto Replying for" & Mobileno & "..."
            If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
            If Comm1.PortOpen = True Then Comm1.PortOpen = False
            
        'frmSend.dialsms Mobileno, rs.Fields("AUTOMESSAGE")
            If Comm1.PortOpen = False Then Comm1.PortOpen = True
                  If frmSend.Message_sent = 1 Then
                           
                           StatusBar1.Panels(1).Text = "Message successfully sent"
                           Message_sent = 1
                   End If
                  If frmSend.Service_Number = "wrong" Then
                           MsgBox "Sim Card expired or service center numnber wrong"
                           StatusBar1.Panels(1).Text = "Check service center number"
                           
                           Exit Sub
                  End If
                If frmSend.Modem_Connect = o Then
                    MsgBox "Check Modem connections"
                    StatusBar1.Panels.item(1).Text = " Check Modem connection"
                    cmdStop.Enabled = True
                 Exit Sub
                End If
            End If
            rs.MoveNext
            Wend
                  Label3.Caption = "TO"
                  Label4.Caption = "Message"
                  labelfrom.Caption = Mobile_Number
                  Labelkeyword.Caption = DefaultMessage
                  StatusBar1.Panels(1).Text = "Sending Default Message to " & Mobileno & "..."
                  If Comm1.PortOpen = True Then Comm1.PortOpen = False
            If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
            If Comm1.PortOpen = True Then Comm1.PortOpen = False
           ' frmSend.dialsms Mobile_Number, DefaultMessage
            If Comm1.PortOpen = False Then Comm1.PortOpen = True
                  If frmSend.Message_sent = 1 Then
                           StatusBar1.Panels(1).Text = "Message successfully sent"
                           Message_sent = 1
                   End If
                  If frmSend.Service_Number = "wrong" Then
                           MsgBox "Sim Card expired or service center numnber wrong"
                           StatusBar1.Panels(1).Text = "Sim Card expired or service center numnber wrong"
                  End If
                If frmSend.Modem_Connect = o Then
                    MsgBox "Check Modem connections"
                    StatusBar1.Panels.item(1).Text = " Check Modem connection"
                Exit Sub
                End If
       
       
       
   
    
    rs.Close
    Set rs = Nothing
    StatusBar1.Panels(1).Text = "Waiting for Messages....."
    Label3.Caption = "From"
    Label4.Caption = "Message"
    labelfrom.Caption = ""
    Labelkeyword.Caption = ""
LoadInbox Mobileno, recieved_message, time_stamp

End Sub



Private Sub DeleteMessage(ByVal MessageNumber As Integer)
Dim DialString$
DialString = "AT+CMGD=" & MessageNumber & vbCrLf
Call comm1_settings
    StatusBar1.Panels(1).Text = "Deleting message from Simcard....."
    If Comm1.PortOpen = False Then Comm1.PortOpen = True
           Comm1.Output = DialString$
           Do
   If Comm1.InBufferCount Then FromModem$ = FromModem$ + Comm1.Input
             If InStr(FromModem$, "OK") Then
             StatusBar1.Panels(1).Text = "Message Deleted from Simcard"
             GoTo Deleted
            End If
    
  Loop
Deleted:
StatusBar1.Panels(1).Text = "Waiting for Messages..."
End Sub

Private Sub Modem_checking()
    Dim DialString$, FromModem$, start, dummy, fname1, fbuff
    Call Comm_settings
'    If port_check = False Then
'       msgbox "Modem not connected"
'    End If
    FromModem$ = ""
    MSComm1.InBufferCount = 0
    MSComm1.InputLen = 0
    Do
    
        If AT_flag = False Then
            If MSComm1.PortOpen = True Then
            MSComm1.Output = "AT" + vbCr
            End If
        End If
        If AT_flag1 = True Then
            Exit Sub
        End If

        delay (2)
        Do 'AT
          dummy = DoEvents()
          FromModem$ = ""
          If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
             If InStr(FromModem$, "OK") Then
                AT_flag = True
                
                Debug.Print "AT"
                If AT_flag1 = True Then
                    Exit Sub
                End If
                Exit Do
             End If
             
             If Brate < 5 And AT_flag = False Then
                Set_parity = Baudrate(Brate) & ",n,8,1"
                
                MSComm1.Settings = Set_parity
                Debug.Print Brate
                Brate = Brate + 1
                Exit Do
             ElseIf AT_flag = False Then
                If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
                MsgBox "   Connect the Modem properly   "
                
                StatusBar1.Panels(1).Text = "Simcard Error"
                    cmdStart.Enabled = True
                    cmdDefaultMessage.Enabled = True
                 Exit Sub
                If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
                AT_flag = True
                Exit Sub
             ElseIf AT_flag = True Then
                
                Exit Sub
             End If
        Loop
        If AT_flag = True Then
           Exit Do
        End If
    Loop
        'AT+s0=3
        FromModem$ = ""
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        MSComm1.Output = "ATs0=3" + vbCr
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                    CMGF_flag = True
                
                    Debug.Print "ATs0=3"
                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    
                    CMGF_flag = False
                    MsgBox "ERROR"
                    Exit Sub
                End If
                If Timer > (start + 2) Then
                      Exit Sub
                End If
            Loop
        
        'AT+CMGF=1
        FromModem$ = ""
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        MSComm1.Output = "AT+CMGF=1" + vbCr
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                    CMGF_flag = True
                
                    Debug.Print "AT+cmgf=1"
                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    CMGF_flag = False
                    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
                    MsgBox "    Check the simcard    "
                    StatusBar1.Panels(1).Text = "Simcard Error"
                    cmdStart.Enabled = True
                    cmdDefaultMessage.Enabled = False
                    cmdDefaultMessage.Enabled = True
                    Exit Sub
                End If
                If Timer > (start + 2) Then
                      Exit Sub
                End If
            Loop

End Sub
Private Sub Initialise_Modem()
If Module1.Formload_check = False Then
        Brate = 0
        AT_flag = False
        CMGF_flag = False
        Baudrate = Array("9600", "19200", "38400", "57600", "115200")
        Set_parity = Baudrate(Brate) & ",n,8,1"
        
        Com_count = 1
    End If
    AT_flag1 = False
        SMS_check = False
    
End Sub
Private Sub Comm_settings()
On Error GoTo E1:
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    With MSComm1
        Debug.Print .CommPort
        '.CommPort = 2 'Val(Combo3.ListIndex) + 1
        .CommPort = Com_count
        .InBufferCount = 0
        .InputLen = 1
        .InBufferSize = 1024
        .OutBufferSize = 512
        .Settings = Set_parity '"9600,n,8,1"  'Combo4.Text & ",n,8,1"
        'MSmscomm1.DTREnable = True
        .RTSEnable = True
    End With
    port_check = True
If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
Exit Sub
E1:
 If Err.Number = 8005 Then
    
    MsgBox Err.Description
    port_check = False
 ElseIf Err.Number = 8002 Then
    Com_count = Com_count + 1
    port_check = False
    If Com_count = 4 Then
        MsgBox Err.Description
        port_check = False
        Exit Sub
    End If
    Call Comm_settings
 End If
End Sub

Private Sub comm1_settings()
If Comm1.PortOpen = True Then Comm1.PortOpen = False
If Comm1.PortOpen = False Then
   'Comm1.PortOpen = True
   Comm1.DTREnable = True
   Comm1.RTSEnable = True
   Comm1.RThreshold = 1
   Comm1.InputLen = 1
   Comm1.Settings = "9600, n, 8, 1"
   bOK = False
   bError = False
   Comm1.PortOpen = True
   Comm1.Output = "AT" + vbCrLf
    delay (1)
   End If
 Do
   If Comm1.InBufferCount Then FromModem$ = FromModem$ + Comm1.Input
             If InStr(FromModem$, "OK") Then
             'MsgBox "ready"
             GoTo finish
            End If
    
  Loop
finish:
End Sub

Private Sub Command3_Click()
Dim pos1, pos2
pos1 = 1
If Comm1.PortOpen = True Then Comm1.PortOpen = False
Call Command2_Click
bOK = False
bError = False
Comm1.Output = "AT+CMGL=" & Chr(34) & "ALL" & Chr(34) & vbCrLf
While Not bOK Or bError
  bMessageStore = True
  DoEvents
  Wait
Wend
If bOK Then
   ReadMessage
   'MsgBox txtMessage.Text
   'MsgBox txtTelephone.Text
'  pos1 = InStr(txtMessage.Text, txtTelephone.Text, "/", vbTextCompare)
   'MsgBox txtMessage
'   If InStr(1, UCase(txtMessage.Text), "NOTEPAD", vbTextCompare) <> 0 Then
'      'Call ExecuteCommand("NotePad.exe")
'   ElseIf InStr(1, UCase(txtMessage.Text), "CALC", vbTextCompare) <> 0 Then
'      Call ExecuteCommand("Calc.exe")

CheckData
   
End If
If bError Then
   txtMessage.Text = "Bad Read"
End If
'Comm1.PortOpen = False
End Sub
    

Private Sub Command2_Click()
Dim DialString$, FromModem$
If Comm1.PortOpen = False Then
   'Comm1.PortOpen = True
   Comm1.DTREnable = True
   Comm1.RTSEnable = True
   Comm1.RThreshold = 1
   Comm1.InputLen = 1
   Comm1.Settings = "9600, n, 8, 1"
   bOK = False
   bError = False
   Comm1.PortOpen = True
   Comm1.Output = "AT" + vbCrLf
    delay (1)
   End If
 Do
   If Comm1.InBufferCount Then FromModem$ = FromModem$ + Comm1.Input
             If InStr(FromModem$, "OK") Then
             'MsgBox "ready"
             GoTo finish
            End If
 Loop
   MsgBox "Port Already Open !", vbCritical + vbOKOnly, "Error opening port"
finish:
'MsgBox day(
'If Comm1.PortOpen = True Then Comm1.PortOpen = False
'frmChild.Timer1.Enabled = True
StatusBar1.Panels.item(1).Text = "Reading Message..."
End Sub

Private Sub SendMsgToAdmin()

 If InStr(1, UCase(Message), "PASSWORD", vbTextCompare) Then
 
 MsgBox "change password"
 
 End If
 
DeleteMessage MessageCount
End Sub
