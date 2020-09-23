VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScheduler 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddSchedule 
      Caption         =   "Add Message"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add a Scheduled Message"
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
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
         Left            =   240
         TabIndex        =   11
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox cmbHours 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Text            =   "cmbHours"
         Top             =   3720
         Width           =   615
      End
      Begin VB.ComboBox cmbAMPM 
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Text            =   "cmbAMPM"
         Top             =   3720
         Width           =   615
      End
      Begin VB.ComboBox cmbMinutes 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Text            =   "cmbMinutes"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtScheduleTime 
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtMessage 
         Height          =   1935
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Date/Time"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Message"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "To "
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstSchedule 
      Height          =   1455
      Left            =   -360
      TabIndex        =   4
      Top             =   6000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2566
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglist"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
         ImageIndex      =   3
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public temptime
Dim A As ListItem
Option Explicit
Dim RecordCount As String
Private Type DataFile
    Index As Long
    projectpath As String
    projectdate As String
    projecttime As String
End Type
Public ScheduledMessage
Dim DTFile() As DataFile
Dim cnt As Long
Dim strmessage
'Dim files2 As clsFile
Dim comm As String

Dim jam As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
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
            
    LoadSchedule
End Sub

Private Sub Timer1_Timer()

    txtTimer.Text = Format(Now(), "yyyy-mm-dd hh:mm:ss ")
    frmSendSingleMessage.txtTimer.Text = Format(Now(), "yyyy-mm-dd hh:mm:ss ")
    frmSend.txtTimer.Text = Format(Now(), "yyyy-mm-dd hh:mm:ss ")
   ' Debug.Print Now
    CheckforRecords
    
    
    
End Sub



Private Sub cmdAddSchedule_Click()

    txtScheduleTime.Text = txtDate + " " + cmbHours + ":" + cmbMinutes + " " + cmbAMPM
        
    If txtMessage = "" Or txtNumber = "" Or txtScheduleTime = "" Then
    
    MsgBox "Enter Valid Details"
    Exit Sub
    
    
    End If
 
    
addNewSchedule
Unload Me
End Sub

Public Sub LoadSchedule()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   On Error Resume Next
   
        lstSchedule.ListItems.Clear
        lstSchedule.ColumnHeaders.Clear
        lstSchedule.ColumnHeaders.Add , , "Name", 2000
        lstSchedule.ColumnHeaders.Add , , "Number", 1500
        lstSchedule.ColumnHeaders.Add , , "Message", 5000
        lstSchedule.ColumnHeaders.Add , , "Schedule time"
        lstSchedule.ColumnHeaders.Add , , "ID"
        
    Set rs = New ADODB.Recordset
    strQuery = "select * from SCHEDULEMESSAGES"
    rs.Open strQuery, con, 3, 2, 1
   
    While Not rs.EOF
        
       Set A = lstSchedule.ListItems.Add(, , rs.Fields("NAME"))
               
               A.SubItems(1) = rs.Fields("MOBILENO")
               A.SubItems(2) = rs.Fields("MESSAGE")
               A.SubItems(3) = rs.Fields("SCHEDULETIME")
               A.SubItems(4) = rs.Fields("ID")
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub addNewSchedule()
Dim tempName As String

  Dim rs As ADODB.Recordset
    Dim strQuery As String
      
           On Error GoTo OpenError
        
            Set rs = New ADODB.Recordset
                strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & txtNumber.Text & "'"
            
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
         strQuery = "insert into SCHEDULEMESSAGES(MOBILENO,MESSAGE,NAME,SCHEDULETIME) values ('" & txtNumber.Text & "', '" & txtMessage.Text & "', '" & tempName & "','" & txtScheduleTime.Text & "')"
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing

        

End Sub

Private Sub CheckforRecords()

Dim i As Integer
Dim strQuery As String
Dim rs As ADODB.Recordset


RecordCount = 0
temptime = LCase(Format(Now(), "yyyy-mm-dd hh:mm"))
   ' Debug.Print Now
        Set rs = New ADODB.Recordset
    
           strQuery = "Select * from SCHEDULEMESSAGES where  format(SCHEDULETIME,'" & "yyyy-mm-dd hh:mm" & " ') = '" & temptime & "'"
    'Debug.Print strQuery
           
        rs.Open strQuery, con, 3, 2, 1
        i = 0
    While Not rs.EOF
            
               RecordCount = RecordCount + 1
               ' Timer1.Enabled = False
        rs.MoveNext
    Wend
                
               
   
    
   If RecordCount <> 0 Then
   frmSendSchedule.Show
   'Timer1.Enabled = True
   Exit Sub
   End If
End Sub



Private Sub txtDate_Click()
frmSingleCalender.Show
'txtDate.Text = frmSingleCalender
End Sub

