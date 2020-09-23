VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSendSchedule 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sending Schedule Messages"
   ClientHeight    =   4365
   ClientLeft      =   2535
   ClientTop       =   2535
   ClientWidth     =   7590
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4110
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstSchedule 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3413
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
   Begin VB.Frame Frame1 
      Caption         =   "Sending Scheduled Messages"
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7455
      Begin VB.Label Label1 
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   2760
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmSendSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mobileno
Dim Message
Dim ContactName


Public Sub LoadSchedule()
Dim temptime
temptime = frmChild.temptime
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   On Error Resume Next
   'lstSchedule.Sorted = True
  ' lstSchedule.SortKey = 4
        lstSchedule.ListItems.Clear
        lstSchedule.ColumnHeaders.Clear
        lstSchedule.ColumnHeaders.Add , , "Name", 2000
        lstSchedule.ColumnHeaders.Add , , "Number", 1500
        lstSchedule.ColumnHeaders.Add , , "Message", 5000
        lstSchedule.ColumnHeaders.Add , , "Schedule time"
        lstSchedule.ColumnHeaders.Add , , "ID"
        
    Set rs = New ADODB.Recordset
     strQuery = "Select * from SCHEDULEMESSAGES where  format(SCHEDULETIME,'" & "yyyy-mm-dd hh:mm" & " ') = '" & temptime & "'"
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
    
     
    For i = 1 To lstSchedule.ListItems.Count
    
    lstSchedule.ListItems.item(i).Checked = True
    
    Next i
   
    
    
    
    
  
End Sub
Public Sub CheckStatus()
   Dim totalsent
   Dim totalfailed
    If frmChild.Message_sent = "True" Then
          
            StatusBar1.Panels.item(1).Text = " Message Sent "
            'lstAddressBook.ListItems(i).Checked = False
             
               SaveToSentMessages Mobileno, Message, ContactName
            
            totalsent = totalsent + 1
             StatusBar1.Panels(1).Text = "Number of Messages sent   " & totalsent
             End If
          If frmChild.Message_sent = "False" Then
      AddtoFailed Mobileno, Message, ContactName
          StatusBar1.Panels(1).Text = "Message Sent Failed"
          
          totalfailed = totalfailed + 1
           StatusBar1.Panels(2).Text = "Number of Messages Failed" & totalfailed
          End If
          If frmChild.Message_sent = "NotPossible" Then
           totalfailed = totalfailed + 1
            StatusBar1.Panels(2).Text = "Number of Messages Failed" & totalfailed
          AddtoFailed Mobileno, Message, ContactName
          
          End If
          
       
    
End Sub

Private Sub Form_Load()
frmMain.Enabled = False
LoadSchedule
Me.Show
SendMessages

End Sub

Private Sub SendMessages()

Label1.Caption = " Sending Scheduled Messages..........."

For i = 1 To lstSchedule.ListItems.Count
If lstSchedule.ListItems.item(i).Checked = True Then
Mobileno = lstSchedule.ListItems(i).ListSubItems.item(1).Text
Message = lstSchedule.ListItems.item(i).ListSubItems.item(2).Text
ContactName = lstSchedule.ListItems.item(i).Text
frmChild.SendSms Mobileno, Message
CheckStatus
End If
Next i
Label1.Caption = "Scheduled Messages Sent"
DeleteSentSchedule
frmMain.Enabled = True
frmChild.Timer2.Enabled = True
Exit Sub
Unload Me
End Sub


Private Sub DeleteSentSchedule()
Dim temptime
temptime = frmChild.temptime
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   On Error Resume Next


    strQuery = "delete from SCHEDULEMESSAGES where  format(SCHEDULETIME,'" & "yyyy-mm-dd hh:mm" & "') = '" & temptime & "'"
    Debug.Print strQuery
    con.Execute strQuery

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
                Timer1.Enabled = False
        rs.MoveNext
    Wend
                
               
   
    
   If RecordCount <> 0 Then
   frmSendSchedule.Show
   Timer1.Enabled = True
   Exit Sub
   End If
End Sub

Public Sub SaveToSentMessages(ByVal Mobile As String, ByVal Message As String, ByVal name As String)
Dim TimeStamp As String
TimeStamp = Now
 Dim rs As ADODB.Recordset
    Dim strQuery As String
        
        Set rs = New ADODB.Recordset
         strQuery = "insert into SENTMESSAGES(MOBILENO,MESSAGE,TIME_STAMP,STATUS,NAME) values ('" & Mobile & "', '" & Message & "', '" & TimeStamp & "',1,'" & name & "')"
       'strQuery = "insert into OUTBOX(MOBILENO, MESSAGE,STATUS,TIMESTAMP) values ('" & txtSend & "', '" & txtMessage & "','Saved Message ', '" & Now & "')"
            'InputBox "", "", strQuery
            con.Execute strQuery
            Set rs = Nothing

           

End Sub

Public Sub AddtoFailed(ByVal Mobileno As String, ByVal Message As String, ByVal name As String)

Dim TimeStamp As String
TimeStamp = Now
 Dim rs As ADODB.Recordset
    Dim strQuery As String
            Set rs = New ADODB.Recordset
         strQuery = "insert into FAILEDMESSAGES(MOBILENO,MESSAGE,TIME_STAMP,NAME) values ('" & Mobileno & "', '" & Message & "', '" & TimeStamp & "','" & name & "')"
      
            con.Execute strQuery
            Set rs = Nothing

End Sub

