VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmreadPhonebook 
   Caption         =   "Read Sim Phone book"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4560
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdread 
      Caption         =   "Get Sim Book"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frmreadPhonebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim msgBreak()                 As String
Dim msgHeader()                As String
Public Baudrate, Set_parity, Brate, Com_count, fnam1
Dim AT_flag, AT_flag1, CMGF_flag, Data_check, Voice_check, SMS_check, Inet_check, config_check As Boolean
Dim serviceNumber As String
Dim strbuffer As String





Private Sub Form_Load()
If frmSend.MSComm1.PortOpen = True Then frmSend.MSComm1.PortOpen = False
If frmChild.MSComm1.PortOpen = True Then frmChild.MSComm1.PortOpen = False
If frmReadMessages.Comm1.PortOpen = True Then frmReadMessages.Comm1.PortOpen = False
If frmReadMessages.MSComm1.PortOpen = True Then frmReadMessages.MSComm1.PortOpen = False

If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
cmdread.Enabled = False
Initialise_Modem
Call Modem_checking
'CheckAT
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
Private Sub CheckAT()
   
End Sub

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
        If cmdread.Enabled = False Then
       
        MSComm1.Output = "AT+CPBF=" & Chr(34) & "" & Chr(34) & vbCrLf
        
        Else
        'MSComm1.Output = "AT+CSCA=" + txtmessage + vbCr
                End If
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                strbuffer = FromModem$
                    CMGF_flag = True
                    Exit Sub
                'txtMessage.Text
'                If Command1.Enabled = False Then
'                msgBreak = Split(FromModem$, vbCrLf, , vbTextCompare)
'                msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
'                'MsgBox msgBreak(1)
'               serviceNumber = Mid$(Right$(msgBreak(1), 18), 1, 13)
'                 txtmessage.Text = serviceNumber
'                 Command1.Enabled = True
'                 Else
'                 MsgBox "number changed"
'                 End If
'                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    
                    CMGF_flag = False
                    MsgBox "ERROR"
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
Private Sub Command1_Click()
Initialise_Modem
Modem_checking
Unload Me
End Sub


