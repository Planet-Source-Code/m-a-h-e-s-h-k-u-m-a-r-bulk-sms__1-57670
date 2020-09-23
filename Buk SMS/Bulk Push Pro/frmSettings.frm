VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   ClipControls    =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   780
      Left            =   90
      TabIndex        =   9
      Top             =   2550
      Width           =   4650
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   315
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2445
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3510
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3525
      Width           =   1020
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      ForeColor       =   &H8000000A&
      Height          =   2340
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4620
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSettings.frx":030A
         Left            =   2340
         List            =   "frmSettings.frx":031A
         TabIndex        =   3
         Top             =   330
         Width           =   1995
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSettings.frx":0336
         Left            =   2340
         List            =   "frmSettings.frx":0338
         TabIndex        =   2
         Top             =   930
         Width           =   2000
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSettings.frx":033A
         Left            =   2340
         List            =   "frmSettings.frx":0389
         TabIndex        =   1
         Top             =   1515
         Width           =   2000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Modem Com Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   345
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Modem Baud Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   990
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Modem Ring Count"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1575
         Width           =   1965
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim AT_flag, CMGF_flag, port_check  As Boolean
Public Sub Comm_settings()
On Error GoTo E1:
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    With MSComm1
        .CommPort = 1  'Val(Combo1.ListIndex) + 1
        .InBufferCount = 0
        .InputLen = 1
        .InBufferSize = 1024
        .OutBufferSize = 512
        .InBufferCount = 0
        .Settings = frmSend.Set_parity    'Combo4.Text & ",n,8,1"
        .RTSEnable = True
    End With
If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
Combo1.Text = "Com" & frmSend.Com_count
port_check = True
Exit Sub
E1:
   MsgBox Err.Description
   port_check = False
   End Sub

Private Sub Command1_Click()
'frmChild.Timer1.Enabled = False
    Command1.Enabled = False
    PBar.Value = 0
    PBar.Max = 100
    PBar.Min = 0
    AT_flag = False
    CMGF_flag = False
    Call Modem_checking
    Command1.Enabled = True
    PBar.Value = 0
End Sub
Private Sub Modem_checking()
    Dim DialString$, FromModem$, start, dummy, fname1, fbuff
    Call Comm_settings
    If port_check = False Then
        Exit Sub
    End If

    FromModem$ = ""
    MSComm1.InBufferCount = 0
    MSComm1.InputLen = 0
    Do
        If AT_flag = False Then
            MSComm1.Output = "AT" + vbCr
        End If
        frmSend.delay (1)
        start = Timer
        Do 'AT
          dummy = DoEvents()
          If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
             If InStr(FromModem$, "OK") Then '+CSQ: 26,0
                AT_flag = True
                Debug.Print "AT"
                Exit Do
             End If
             'If Form4.Brate < 10 And AT_flag = False Then
             '   Form4.Set_parity = Form4.BaudRate(Form4.Brate) & ",n,8,1"
             '   MSComm1.Settings = Form4.Set_parity
              '  Form4.Brate = Form4.Brate + 1
              '  Exit Do
             If AT_flag = False Then
                MsgBox "MODEM NOT RESPONDING"
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
        PBar.Value = 25
        'AT+ipr
        FromModem$ = ""
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        MSComm1.Output = "AT+ipr=" + Combo2.Text + vbCr
        frmSend.delay (1)
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                    CMGF_flag = True
                    Debug.Print "AT+ipr"
                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    CMGF_flag = False
                    MsgBox "MODEM NOT RESPONDING"
                    Exit Sub
                End If
                If Timer > (start + 20) Then
                    Exit Sub
                End If
            Loop
        PBar.Value = 50
        'AT+S0=
        FromModem$ = ""
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        frmSend.Set_parity = Combo2.Text + ",n,8,1"
        frmSend.Brate = Combo2.ListIndex + 1
        Call Comm_settings
        If port_check = False Then
            Exit Sub
        End If

        frmSend.delay (1)
        MSComm1.Output = "ATS0=" + Combo3.Text + vbCr
        frmSend.delay (1)
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                    CMGF_flag = True
                    Debug.Print "AT+S0"
                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    CMGF_flag = False
                    MsgBox "MODEM NOT RESPONDING"
                    Exit Sub
                End If
                If Timer > (start + 2) Then
                    Exit Sub
                End If
            Loop
         PBar.Value = 70
        'ATe1
        FromModem$ = ""
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        frmSend.delay (1)
        MSComm1.Output = "ATE1" + vbCr
        frmSend.delay (1)
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                    CMGF_flag = True
                    Debug.Print "ATE1"
                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    CMGF_flag = False
                    MsgBox "MODEM NOT RESPONDING"
                    Exit Sub
                End If
                If Timer > (start + 2) Then
                    Exit Sub
                End If
            Loop
         PBar.Value = 90
        'AT&W
        FromModem$ = ""
        MSComm1.InBufferCount = 0
        MSComm1.InputLen = 0
        MSComm1.Output = "AT&W" + vbCr
        frmSend.delay (1)
        start = Timer
            Do
              dummy = DoEvents()
              If MSComm1.InBufferCount Then FromModem$ = FromModem$ + MSComm1.Input
                If InStr(FromModem$, "OK") Then
                    CMGF_flag = True
                    Debug.Print "AT&W"
                    Exit Do
                ElseIf InStr(FromModem$, "ERROR") Then
                    CMGF_flag = False
                    MsgBox "MODEM NOT RESPONDING"
                    Exit Sub
                End If
                If Timer > (start + 2) Then
                    Exit Sub
                End If
            Loop
      PBar.Value = 100
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Combo2.List(0) = 110
    Combo2.List(1) = 300
    Combo2.List(2) = 1200
    Combo2.List(3) = 2400
    Combo2.List(4) = 4800
    Combo2.List(5) = 9600
    Combo2.List(6) = 19200
    Combo2.List(7) = 38400
    Combo2.List(8) = 57600
    Combo2.List(9) = 115200

    Combo1.Text = "Com1"
    Combo2.Text = "9600"
    Combo3.Text = "2"
    Combo1.ListIndex = 0
    Combo2.ListIndex = 5
    Combo3.ListIndex = 1
    frmMain.Enabled = False
'frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False

    'Combo2.Text = Mid(Form4.Set_parity, 1, InStr(Form4.Set_parity, ",") - 1)
    'Combo2.ListIndex = Form4.Brate - 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
'frmChild.Timer1.Enabled = True
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub

