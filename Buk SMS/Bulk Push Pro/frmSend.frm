VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send"
   ClientHeight    =   7485
   ClientLeft      =   1860
   ClientTop       =   780
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSend.frx":0442
   ScaleHeight     =   7485
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtScheduleTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   36
      Text            =   "Text3"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSendLater 
      Caption         =   "Send Later"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   35
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Date/Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   34
         Top             =   720
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   33
         Top             =   240
         Width           =   1065
      End
      Begin VB.ComboBox cmbMinutes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   32
         Text            =   "cmbMinutes"
         Top             =   240
         Width           =   750
      End
      Begin VB.ComboBox cmbAMPM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   31
         Text            =   "cmbAMPM"
         Top             =   240
         Width           =   750
      End
      Begin VB.ComboBox cmbHours 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Text            =   "cmbHours"
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtTimer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4995
      TabIndex        =   26
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   4080
      Width           =   855
   End
   Begin MSComctlLib.ListView lstAddressBook 
      Height          =   2175
      Left            =   4800
      TabIndex        =   20
      Top             =   1440
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   3836
      View            =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483648
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Mobile Number"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   1560
      Width           =   3135
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7230
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Messages Sent "
            TextSave        =   "Messages Sent "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Failed Messages"
            TextSave        =   "Failed Messages"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveToOutbox 
      Caption         =   "Save to &Outbox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MouseIcon       =   "frmSend.frx":0594
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame frameSend 
      Caption         =   "Select Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   7455
      Begin VB.CheckBox CheckSave 
         Caption         =   "Save to sent messages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddNewContact 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   22
         Top             =   720
         Width           =   195
      End
      Begin VB.ComboBox cmbGroups 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   3075
      End
      Begin VB.TextBox txtSend 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddressBook 
         Caption         =   "&Phone Book"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton rdSelectFromAddressBook 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton rdSend 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton rdGroups 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   23
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label LabelNoOfRecords 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label LabelI 
         AutoSize        =   -1  'True
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   3720
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label LabelStat 
         Caption         =   "Sending       of "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label labelNumberOfCharacters 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   4
         Top             =   3360
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "Send To &Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   610
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Variant
Public Baudrate, Set_parity, Brate, Com_count, fnam1
Dim AT_flag, AT_flag1, CMGF_flag, Data_check, Voice_check, SMS_check, Inet_check, config_check, port_check As Boolean
Public fname1, start, dummy, fbuff, DialString$, FromModem$
Dim connect_flag, Voice_flag, disconnect_flag As Boolean
Public SelectedNode As String
Public Service_Number As String
Public Message_sent As Integer
Dim temp As Integer
Public Modem_Connect As String
Public totalsent
Public totalfailed
Dim stopsending
Public SendingMessage
'Dim fsys As New FileSystemObject
Private Sub cmdAddressBookl_Click()
    'frmSendSelect.Show
End Sub



Private Sub Check1_Click()
If Check1 = 1 Then
For i = 1 To frmSend.lstAddressBook.ListItems.Count
    frmSend.lstAddressBook.ListItems.item(i).Checked = True
    Next i
Else
For i = 1 To frmSend.lstAddressBook.ListItems.Count
    frmSend.lstAddressBook.ListItems.item(i).Checked = False
    Next i
    
    End If

End Sub


Private Sub cmbGroups_Click()
EnableAll
frmMain.LoadGroupSendList cmbGroups
End Sub


Private Sub cmbGroups_DropDown()
'MsgBox cmbGroups
'frmMain.LoadGroupSendList cmbGroups
End Sub



Private Sub cmdAddNewContact_Click()
frmAddEditContact.Show
If rdGroups = True Then
frmAddEditContact.cmbGroups.Text = cmbGroups.Text
End If

End Sub

Private Sub cmdAddressBook_Click()
    'frmSendSelect.Show
End Sub

Private Sub cmdCancel_Click()
MousePointer = 0
   Unload Me
  
End Sub
Private Sub cmdSaveToOutBox_Click()
If txtmessage = "" Then
MsgBox "Type message", vbInformation
Exit Sub
End If
SaveToOutbox
End Sub

Private Sub cmdSend_Click()

    If txtmessage = "" Then
        MsgBox "Type message", vbInformation
    Exit Sub
    End If
    
   stopsending = "no"
   
   ' frmChild.Timer1.Enabled = False
    cmdSend.Enabled = False
    txtmessage.Enabled = False
    cmdSaveToOutBox.Enabled = False
   cmdCancel.Enabled = False
   lstAddressBook.Enabled = False
   Command1.Enabled = False
   cmdAddNewContact.Enabled = False
   Check1.Enabled = False
   CheckSave.Enabled = False
   cmbGroups.Enabled = False
   cmdSendLater.Enabled = False
   
  SendingMessage = "yes"
   frmChild.Timer1.Enabled = False
  'Label2.Caption = SendingMessage
   Select Case True
    Case rdGroups
         GroupMessage
    Case rdSelectFromAddressBook
         SelectedSend
    Case rdSend
     
  End Select
  frmChild.Timer1.Enabled = True
  
 SendingMessage = "no"
 'Label2.Caption = SendingMessage
   EnableAll
    
End Sub

Private Sub cmdSendLater_Click()
If Frame2.Visible = False Then
Me.Height = 7170
Frame2.Visible = True
cmdSend.Enabled = False
Else
Frame2.Visible = False
Me.Height = 5895
cmdSend.Enabled = True
End If
End Sub

Private Sub cmdStop_Click()
stopsending = "yes"
End Sub

Private Sub Command1_Click()
txtmessage.Text = ""
End Sub





Private Sub Form_Load()
Me.Height = 5895
txtdate = Format(Now, "mm/dd/yyyy")
Dim i
            For i = "01" To 12
            cmbHours.AddItem i
            Next i
            cmbHours.Text = "00"
            For i = 0 To 59
            cmbMinutes.AddItem i
            Next i
            cmbMinutes.Text = "00"
            cmbAMPM.AddItem "AM"
            cmbAMPM.AddItem "PM"
            cmbAMPM.Text = "AM"

'frmChild.Timer1.Enabled = False
Dim start
Dim rs As ADODB.Recordset
Dim D() As String
    Dim strQuery As String

    Set rs = New ADODB.Recordset
    strQuery = "select GROUPNAME from GROUPS"

    rs.Open strQuery, con, 3, 2, 1

    While Not rs.EOF
        cmbGroups.AddItem rs.Fields("GROUPNAME")
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing

'   cmbGroups.ListIndex = 0

    If SelectedNode <> "" Then
        While True
            If cmbGroups = SelectedNode Then
                GoTo finish
            Else
                cmbGroups.ListIndex = cmbGroups.ListIndex + 1
            End If
        Wend
    End If

finish:
'MsgBox frmChild.bpp_tree.SelectedItem.Text
On Error GoTo handler
D = Split(frmChild.bpp_tree.SelectedItem.Tag, "|")


    If D(0) = "GRP" Then
cmbGroups = frmChild.bpp_tree.SelectedItem.Text
End If
handler:
'frmChild.Timer1.Enabled = False
'frmMain.StatusBar.Panels.item(2).Text = ""
LabelStat.Visible = False
'frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False
'Call Comm_settings
End Sub


Private Sub Form_Unload(Cancel As Integer)
  
    frmMain.Enabled = True
    frmMain.Toolbar1.Enabled = True
Unload Me
End Sub

Private Sub lstAddressBook_Click()
On Error Resume Next
lstAddressBook.ToolTipText = lstAddressBook.SelectedItem.ListSubItems.item(1)
End Sub

Private Sub lstAddressBook_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim i As Integer
'Check1 = 0
'For i = 1 To lstAddressBook.ListItems.count
'If lstAddressBook.ListItems.item(i).Checked = False Then
'Check1.Enabled = True
'End If
'Next i
End Sub

Private Sub rdGroups_Click()
    If rdGroups.Value Then
        txtSend.Enabled = False
        cmdAddressBook.Enabled = False
        cmbGroups.Enabled = True
        lstAddressBook.Enabled = False
    Else
        txtSend.Enabled = True
        cmdAddressBook.Enabled = True
        cmbGroups.Enabled = False
        lstAddressBook.Enabled = True
    End If
End Sub

Private Sub rdSelectFromAddressBook_Click()

'If frmScheduler.ScheduledMessage = 1 Then Exit Sub
    If rdSelectFromAddressBook Then
        txtSend.Enabled = False
        cmdAddressBook.Enabled = True
        lstAddressBook.Enabled = True
        cmbGroups.Enabled = False
    Else
        txtSend.Enabled = False
        cmdAddressBook.Enabled = True
        lstAddressBook.Enabled = True
        cmbGroups.Enabled = False
    End If
End Sub

Private Sub rdSend_Click()
    If rdSend Then
        txtSend.Enabled = True
        cmdAddressBook.Enabled = False
        cmbGroups.Enabled = False
        lstAddressBook.Enabled = False
    Else
        txtSend.Enabled = False
        cmdAddressBook.Enabled = True
        cmbGroups.Enabled = True
        lstAddressBook.Enabled = True
    End If
End Sub




Private Sub txtDate_Click()
frmSingleCalender.Show
End Sub

Private Sub txtMessage_Change()
If Len(txtmessage) >= 160 Then
MsgBox "Exceeded Messge length"
End If
labelNumberOfCharacters.Caption = Len(txtmessage)
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
End Sub




Private Sub GroupMessage()


totalsent = 0
totalfailed = 0
Dim txttempmessage
cmdSaveToOutBox.Enabled = False
MousePointer = 11
Dim temp As Integer
Dim RecordCount As Integer

If Len(txtmessage) = 160 Then
MsgBox "Message limit reached. Message sent as two messages", vbInformation

End If
    temp = 0
    RecordCount = 0
    If txtmessage = "" Then
         MsgBox "Type any Message"
         StatusBar1.Panels.item(1).Text = "Type any message"
       '  picgreen.Visible = True
         MousePointer = 0
         EnableAll
         Exit Sub
    End If
    For i = 1 To lstAddressBook.ListItems.Count
     If lstAddressBook.ListItems.item(i).Checked = True Then
        RecordCount = RecordCount + 1
     End If
    Next i
    For i = 1 To lstAddressBook.ListItems.Count
        ' picgreen.Visible = False
        If lstAddressBook.ListItems.item(i).Checked = True Then
         temp = temp + 1
         LabelI(0).Visible = True
         LabelNoOfRecords.Visible = True
          LabelStat.Visible = True
        
         LabelI(0).Caption = temp
         LabelNoOfRecords.Caption = RecordCount
         StatusBar1.Panels.item(1).Text = "Sending Message to " & lstAddressBook.ListItems.item(i)
         'LabelStatus.Caption = "Sending Message to " & lstAddressBook.ListItems(i)
         LabelStat.Visible = True
        ' picgreen.Visible = False
         
                txttempmessage = txtmessage
                    If stopsending = "yes" Then
                       Exit Sub
                       EnableAll
                       End If
                    
            
    frmChild.SendSms lstAddressBook.ListItems.item(i).SubItems(1), txtmessage
          If frmChild.Message_sent = "True" Then
          'picgreen.Visible = True
            StatusBar1.Panels.item(1).Text = " Message Sent "
            lstAddressBook.ListItems(i).Checked = False
                If CheckSave = 1 Then
                    SaveToSentMessages lstAddressBook.ListItems.item(i).SubItems(1), txtmessage, lstAddressBook.ListItems.item(i).Text
                End If
            totalsent = totalsent + 1
             End If
          If frmChild.Message_sent = "False" Then
          AddtoFailed lstAddressBook.ListItems.item(i).SubItems(1), txtmessage, lstAddressBook.ListItems.item(i).Text
          StatusBar1.Panels(1).Text = "Message Sent Failed"
          totalfailed = totalfailed + 1
          End If
          If frmChild.Message_sent = "NotPossible" Then
           totalfailed = totalfailed + 1
          AddtoFailed lstAddressBook.ListItems.item(i).SubItems(1), txtmessage, lstAddressBook.ListItems.item(i).Text
          StatusBar1.Panels(1).Text = "Message Sent Failed"
          End If
          
       '  LabelStatus.Caption = "Message Sent"
        'StatusBar1.Panels.Item(2).Text = "Sending"  '"& i &"' + "of" & lstAddressBook.ListItems.count
        LabelStat.Visible = False
        LabelI(0).Caption = temp
        LabelI(0).Visible = False
        LabelNoOfRecords.Visible = False
        StatusBar1.Panels.item(2).Text = "Messages Sent     " & totalsent
        StatusBar1.Panels.item(3).Text = "Failed Messages   " & totalfailed
        End If
        'picgreen.Visible = False
    Next i
   
End Sub

Private Sub SaveToOutbox()
Dim i
Dim TimeStamp As String
TimeStamp = Now
 Dim rs As ADODB.Recordset
    Dim strQuery As String
If MsgBox("Do you want to add this message to Outbox ? ", vbYesNo) = vbYes Then
For i = 1 To lstAddressBook.ListItems.Count
     If lstAddressBook.ListItems(i).Checked Then
        Set rs = New ADODB.Recordset
         strQuery = "insert into OUTBOX(NAME,MOBILENO,MESSAGE,TIME_STAMP) values ('" & lstAddressBook.ListItems(i).Text & "','" & lstAddressBook.ListItems.item(i).SubItems(1) & "','" & txtmessage & "', '" & TimeStamp & "')"
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing
      End If
    Next i
 End If
End Sub

Private Sub SelectedSend()


If Len(txtmessage) = 160 Then
MsgBox "Message limit reached. Message sent as two messages", vbInformation
End If
cmdSaveToOutBox.Enabled = False
MousePointer = 11

Dim i As Integer
Dim RecordCount As Integer
Dim temp As Integer
    temp = 0
    RecordCount = 0
    If txtmessage = "" Then
         'MsgBox "Type any Message"
         StatusBar1.Panels.item(1).Text = "Type any message"
         frmSend.Visible = True
         frmSend.SetFocus
         frmSend.txtmessage.SetFocus
         MousePointer = 0
         EnableAll
         Exit Sub
    End If
        cmdSend.Enabled = False
        cmdSaveToOutBox.Enabled = False
        MousePointer = 11
    If lstAddressBook.ListItems.Count = 0 Then
        StatusBar1.Panels.item(1).Text = " No records found in Phone Book"
        StatusBar1.Panels.item(2).Text = " No records"
    End If
    For i = 1 To lstAddressBook.ListItems.Count
     If lstAddressBook.ListItems.item(i).Checked = True Then
        RecordCount = RecordCount + 1
     End If
    Next i
    
    For i = 1 To lstAddressBook.ListItems.Count
    If lstAddressBook.ListItems.item(i).Checked = True Then
        temp = temp + 1
        LabelI(0).Visible = True
        LabelNoOfRecords.Visible = True
        LabelNoOfRecords.Visible = True
        LabelStat.Visible = True
       
        LabelI(0).Caption = temp
        LabelNoOfRecords.Caption = RecordCount
        StatusBar1.Panels.item(1).Text = "Sending Message to " & lstAddressBook.ListItems.item(i)
        'LabelStatus.Caption = "Sending Message to " & lstAddressBook.ListItems(i)
        LabelStat.Visible = True
                If stopsending = "yes" Then
                Exit Sub
                EnableAll
                End If

        
            frmChild.SendSms lstAddressBook.ListItems.item(i).SubItems(1), txtmessage
       If frmChild.Message_sent = "True" Then
         ' picgreen.Visible = True
            StatusBar1.Panels.item(1).Text = " Message Sent "
            lstAddressBook.ListItems(i).Checked = False
                If CheckSave = 1 Then
                    SaveToSentMessages lstAddressBook.ListItems.item(i).SubItems(1), txtmessage, lstAddressBook.ListItems.item(i).Text
                End If
            totalsent = totalsent + 1
             End If
          If frmChild.Message_sent = "False" Then
          AddtoFailed lstAddressBook.ListItems.item(i).SubItems(1), txtmessage, lstAddressBook.ListItems.item(i).Text
          StatusBar1.Panels(1).Text = "Message Sent Failed"
           totalfailed = totalfailed + 1
          End If
          If frmChild.Message_sent = "NotPossible" Then
            totalfailed = totalfailed + 1
          AddtoFailed lstAddressBook.ListItems.item(i).SubItems(1), txtmessage, lstAddressBook.ListItems.item(i).Text
          StatusBar1.Panels(1).Text = "Message Sent Failed"
          End If
                  
        
        LabelStat.Visible = False
        LabelI(0).Caption = temp
        LabelI(0).Visible = False
        LabelNoOfRecords.Visible = False
        StatusBar1.Panels.item(2).Text = "Total Messages Sent  " & temp
        StatusBar1.Panels.item(3).Text = "Failed Messages   " & totalfailed
        
        End If
    Next i
   
End Sub


Public Sub SaveToSentMessages(ByVal Mobile As String, ByVal Message As String, ByVal name As String)
Dim TimeStamp As String
TimeStamp = Now
 Dim rs As ADODB.Recordset
    Dim strQuery As String
        
        Set rs = New ADODB.Recordset
        On Error Resume Next
         strQuery = "insert into SENTMESSAGES(MOBILENO,MESSAGE,TIME_STAMP,STATUS,NAME) values ('" & Mobile & "', '" & Message & "', '" & TimeStamp & "',1,'" & name & "')"
         Debug.Print strQuery
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing

           

End Sub

Private Sub EnableAll()

cmdSendLater.Enabled = True
        cmdSend.Enabled = True
         cmdSaveToOutBox.Enabled = True
  txtmessage.Enabled = True
    cmdSaveToOutBox.Enabled = True
    cmdCancel.Enabled = True
    lstAddressBook.Enabled = True
    Command1.Enabled = True
    cmdAddNewContact.Enabled = True
    CheckSave.Enabled = True
    MousePointer = 0
    Check1.Enabled = True
    cmbGroups.Enabled = True
    StatusBar1.Panels(1).Text = "Status"


End Sub

Public Sub AddtoFailed(ByVal Mobileno As String, ByVal Message As String, ByVal name As String)
If Mobileno = "" Then Mobileno = "Blank"
If Message = "" Then Message = "Blank"
Dim TimeStamp As String
TimeStamp = Now
 Dim rs As ADODB.Recordset
    Dim strQuery As String
            Set rs = New ADODB.Recordset
         strQuery = "insert into FAILEDMESSAGES(MOBILENO,MESSAGE,TIME_STAMP,NAME) values ('" & Mobileno & "', '" & Message & "', '" & TimeStamp & "','" & name & "')"
      
            con.Execute strQuery
            Set rs = Nothing

End Sub




Private Sub Command5_Click()
txtScheduleTime.Text = Format(txtdate, "mm/dd/yyyy") + " " + cmbHours + ":" + cmbMinutes + " " + cmbAMPM
If Text1.Text = "" Or Text2.Text = "" Or txtScheduleTime = "" Then
    
    MsgBox "Enter Valid Details"
    Exit Sub
End If

SavetoSchedule


Me.Height = 5895
Frame2.Visible = False

cmdSend.Enabled = True
End Sub

Private Sub Command6_Click()
Me.Height = 5895
Frame2.Visible = False
cmdSend.Enabled = True
End Sub

Private Sub SavetoSchedule()
Dim i
If txtmessage = "" Then
MsgBox "Message Cannot be Blank"
Exit Sub
End If
Dim TimeStamp As String
TimeStamp = Format(txtScheduleTime, "mm/dd/yyyy hh:mm")
 Dim rs As ADODB.Recordset
    Dim strQuery As String
If MsgBox("Do you want to add this messages to Schedule ? ", vbYesNo) = vbYes Then
For i = 1 To lstAddressBook.ListItems.Count
    If lstAddressBook.ListItems(i).Checked Then
    
        Set rs = New ADODB.Recordset
         strQuery = "insert into SCHEDULEMESSAGES(NAME,MOBILENO,MESSAGE,SCHEDULETIME) values ('" & lstAddressBook.ListItems(i).Text & "','" & lstAddressBook.ListItems.item(i).SubItems(1) & "','" & txtmessage & "', '" & TimeStamp & "')"
         Debug.Print strQuery
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing
     End If
    Next i
  MsgBox "Message Schedule Added to Scheduled Messages", vbInformation
    StatusBar1.Panels(1).Text = "Message Added to Scheduled Messages"
    
End If
    
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

