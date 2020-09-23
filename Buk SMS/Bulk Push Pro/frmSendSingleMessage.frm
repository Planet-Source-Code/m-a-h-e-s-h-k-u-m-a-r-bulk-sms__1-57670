VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSendSingleMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Message"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmSendSingleMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtScheduleTime 
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Text            =   "Text3"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Date/Time"
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   6255
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
         TabIndex        =   32
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbHours 
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Text            =   "cmbHours"
         Top             =   240
         Width           =   750
      End
      Begin VB.ComboBox cmbAMPM 
         Height          =   315
         Left            =   3600
         TabIndex        =   29
         Text            =   "cmbAMPM"
         Top             =   240
         Width           =   750
      End
      Begin VB.ComboBox cmbMinutes 
         Height          =   315
         Left            =   2880
         TabIndex        =   28
         Text            =   "cmbMinutes"
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         Height          =   360
         Left            =   4560
         TabIndex        =   27
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   4560
         TabIndex        =   26
         Top             =   720
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4560
      TabIndex        =   24
      Top             =   3720
      Width           =   1065
   End
   Begin VB.CommandButton cmdSendLater 
      Caption         =   "Send Later"
      Height          =   360
      Left            =   3360
      TabIndex        =   23
      Top             =   3720
      Width           =   1065
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send Now"
      Default         =   -1  'True
      Height          =   360
      Left            =   720
      TabIndex        =   22
      Top             =   3720
      Width           =   1065
   End
   Begin VB.CommandButton cmdSaveToOutBox 
      Caption         =   "Save to Outbox"
      Height          =   360
      Left            =   1920
      TabIndex        =   21
      Top             =   3720
      Width           =   1305
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5520
      Picture         =   "frmSendSingleMessage.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   19
      Top             =   960
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   5595
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11465
            MinWidth        =   11465
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Type Message"
      Height          =   3315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         Picture         =   "frmSendSingleMessage.frx":0884
         ScaleHeight     =   615
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   5280
         Picture         =   "frmSendSingleMessage.frx":0CC6
         ScaleHeight     =   975
         ScaleWidth      =   495
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox checkSaveSent 
         Caption         =   "Save to sent messages"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   2175
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   2715
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   345
         LinkTimeout     =   0
         TabIndex        =   1
         Top             =   540
         Width           =   1830
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2670
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1605
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1605
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1605
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2130
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2130
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "Clear"
         Height          =   465
         Index           =   10
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2670
         Width           =   1140
      End
      Begin VB.Label lblLength 
         Caption         =   " "
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   2880
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSendSingleMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NameOfContact

Private Sub cmdClear_Click()
Text2.Text = ""
End Sub

Private Sub cmdSaveToOutBox_Click()
Dim tempName As String
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Enter Valid Details", vbInformation
Exit Sub
End If
If Text2.Text = "" Then Text2.Text = "Blank"
Dim TimeStamp As String
TimeStamp = Now
  Dim rs As ADODB.Recordset
    Dim strQuery As String
      
           On Error GoTo OpenError
        
            Set rs = New ADODB.Recordset
                strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & Text1.Text & "'"
            
            rs.Open strQuery, con, 3, 2, 1
            
                tempName = rs.Fields("CONTACTNAME")
            
OpenError:
If Err.Number <> 0 Then
  '  MsgBox Err.Number
   tempName = "Unknown"
   Resume Next
End If
        
        If MsgBox("Do you want to add this message to Outbox ? ", vbYesNo) = vbYes Then
        Set rs = Nothing
        Set rs = New ADODB.Recordset
         strQuery = "insert into OUTBOX(MOBILENO,MESSAGE,TIME_STAMP,NAME) values ('" & Text1.Text & "', '" & Text2.Text & "', '" & TimeStamp & "','" & tempName & "')"
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing
        End If
        StatusBar1.Panels(1).Text = "Message saved in Outbox"

End Sub



Private Sub cmdSend_Click()
If Not IsNumeric(Text1.Text) Then
MsgBox "Enter valid number", vbCritical
Exit Sub
End If
'frmChild.Timer1.Enabled = False
pic1.Visible = False
pic2.Visible = True



Dim i
If Text1.Text = "" Then
    MsgBox "Type number"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "Type Message"
    Text2.SetFocus
    Exit Sub
End If
getName (Text1.Text)
Dim Mobile As String
Dim Message As String
    cmdSend.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    cmdSendLater.Enabled = False
    'Command2.Enabled = False
    cmdClear.Enabled = False
    checkSaveSent.Enabled = False
    For i = 0 To 10
    Command1.item(i).Enabled = False
    Next i
    
    cmdSaveToOutBox.Enabled = False
    MousePointer = 11
    StatusBar1.Panels.item(1).Text = "Sending Message to " & Text1.Text
    Mobile = Text1.Text
    Message = Text2.Text
    pic2.Visible = False
    
    frmChild.Timer1.Enabled = False
                frmChild.SendSms Mobile, Message
    frmChild.Timer1.Enabled = True
    If frmChild.Message_sent = "True" Then
        StatusBar1.Panels.item(1).Text = "Message Succesfully Sent to  " & Text1.Text
       MsgBox "Message Successfully Sent"
            pic2.Visible = True
       If checkSaveSent Then
            frmSend.SaveToSentMessages Text1.Text, Text2.Text, NameOfContact
        End If
    End If
    
    If frmChild.Message_sent = "False" Then
        MsgBox "Message Not Sent", vbCritical
        cmdSend.Enabled = True
        cmdSaveToOutBox.Enabled = True
            Text1.Enabled = True
            Text2.Enabled = True
            cmdClear.Enabled = True
            MousePointer = 0
        StatusBar1.Panels(1).Text = "Message Sending Failed"
        AddtoFailed Text1.Text, Text2.Text, NameOfContact
        EnableAll
        Exit Sub
    End If
    
    If frmChild.Message_sent = "Not Possible" Then
    MsgBox "Message Not Sent", vbCritical
    cmdSend.Enabled = True
    cmdSaveToOutBox.Enabled = True
    AddtoFailed Text1.Text, Text2.Text, NameOfContact
    StatusBar1.Panels(1).Text = "Message Sending Failed"
      EnableAll
      Exit Sub
    End If
   

   
      
   EnableAll
    pic1.Visible = True

End Sub

Private Sub cmdSendLater_Click()
If Frame2.Visible = False Then
    Frame2.Visible = True
    Me.Height = 6345
    cmdSend.Enabled = False
Else
   Frame2.Visible = False
   Me.Height = 4905
   cmdSend.Enabled = True
   
End If
End Sub

Private Sub Command1_Click(Index As Integer)

    If Index = 11 Then
        If Combo1.ListIndex = 0 Or Combo1.ListIndex = 2 Then
            If Text1.Text = "" Then
                MsgBox "Enter Phone Number"
                Exit Sub
            End If
            If Combo1.ListIndex = 0 Or Combo1.ListIndex = 2 Then
                
                If connect_flag = 1 Then
                    Command5.Visible = True
                End If
            End If
        End If
        If Combo1.ListIndex = 1 Then
           If Text2.Text = "" Then
                MsgBox "Type any Message"
                Exit Sub
           End If
           
           
        End If
    ElseIf Index = 10 Then
          Text1.Text = ""
'        If Text1.Text = "" Then
'            Exit Sub
'        Else
'            Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
'        End If
    Else
        Text1.Text = Text1.Text & Index
    End If
End Sub


Private Sub Command2_Click()


frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True
Unload Me

End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
 txtScheduleTime.Text = Format(txtdate, "mm/dd/yyyy") + " " + cmbHours + ":" + cmbMinutes + " " + cmbAMPM
    

If Text1.Text = "" Or Text2.Text = "" Or txtScheduleTime = "" Then
    
    MsgBox "Enter Valid Details"
    Exit Sub
End If
If MsgBox("Do you want to add this messages to Schedule ? ", vbYesNo) = vbYes Then
addNewSchedule
MsgBox "Message Schedule Added to Scheduled Messages", vbInformation
StatusBar1.Panels(1).Text = "Message Added to Scheduled Messages"

End If

Me.Height = 4905
Frame2.Visible = False
cmdSend.Enabled = True
End Sub

Private Sub Command6_Click()
Me.Height = 4905
Frame2.Visible = False
cmdSend.Enabled = True
End Sub

Private Sub Form_Load()
Me.Height = 4905
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmChild.Show
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True
'frmChild.Timer1.Enabled = True
'frmChild.Show

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text2_Change()
If Len(Text2) = 160 Then
MsgBox "Message limit reached. Message will be sent as two messages", vbInformation
End If
lblLength.Caption = Len(Text2.Text)
End Sub


Private Sub EnableAll()
cmdSend.Enabled = True
    'Command2.Enabled = True
     For i = 0 To 10
    Command1.item(i).Enabled = True
    
    Next i
    checkSaveSent.Enabled = True
    cmdSaveToOutBox.Enabled = True
'    Command2.Enabled = True
    MousePointer = 0
    Text1.Enabled = True
    Text2.Enabled = True
    cmdClear.Enabled = True
    cmdSendLater.Enabled = True

End Sub


Private Sub getName(ByVal Number As String)

 Set rs = New ADODB.Recordset
               strQuery = "select * from CONTACTS where MOBILE = '" & Number & "'"
                Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
              NameOfContact = rs.Fields("CONTACTNAME")
                Else
              NameOfContact = "Unknown"
            End If


End Sub

Private Sub txtDate_Click()
frmSingleCalender.Show
End Sub


Private Sub addNewSchedule()
Dim tempName As String

  Dim rs As ADODB.Recordset
    Dim strQuery As String
      
           On Error GoTo OpenError
        
            Set rs = New ADODB.Recordset
                strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & Text1.Text & "'"
            
            rs.Open strQuery, con, 3, 2, 1
            
                tempName = rs.Fields("CONTACTNAME")
            
OpenError:
If Err.Number <> 0 Then
  '  MsgBox Err.Number
   tempName = "Unknown"
   Resume Next
End If


        Set rs = Nothing
        Set rs = New ADODB.Recordset
         strQuery = "insert into SCHEDULEMESSAGES(MOBILENO,MESSAGE,NAME,SCHEDULETIME) values ('" & Text1.Text & "', '" & Text2.Text & "', '" & tempName & "','" & txtScheduleTime.Text & "')"
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing

        

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
            Debug.Print strQuery
            con.Execute strQuery
            
            Set rs = Nothing

End Sub


