VERSION 5.00
Begin VB.Form frmDailyReports 
   Caption         =   "Daily Reports"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Generation"
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3375
      Begin VB.OptionButton optionInbox 
         Caption         =   "Inbox "
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton optionOutbox 
         Caption         =   "Failed Messages"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1500
         Width           =   1815
      End
      Begin VB.OptionButton optionSentmessages 
         Caption         =   "SentMessage"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   990
         Width           =   2655
      End
      Begin VB.TextBox txtdate 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date "
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmDailyReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalender_Click()
frmSingleCalender.Show
End Sub



Private Sub cmdGenerateReport_Click()
    
    If txtdate = "" Then MsgBox "Choose the Date"
    
    If optionInbox Then generateReport "INBOX"

    If optionOutbox Then generateReport "FAILEDMESSAGES"
    If optionSentmessages Then generateReport "SENTMESSAGES"
 
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub generateReport(ByVal table As String)
Dim srchflag
Dim tempdate
        srchflag = False
    
    'txtDate = Format(txtDate, "dd/mm/yyyy")
    tempdate = Format(txtdate, "dd/mm/yyyy")
    Debug.Print txtTodate
    'If Check1 Then
        Set rs = New ADODB.Recordset
               strQuery = "select * from " & table & " where format(TIME_STAMP,'mm/dd/yyyy') =  '" & tempdate & "'  "
        Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                rpMessages.Show
                'rpMessages.SetFocus
               frmMain.Enabled = True
                srchflag = True

            End If
    
If srchflag = False Then MsgBox "No Records Found"
 
   
End Sub

Private Sub Form_Load()
txtdate = Format(Now, "dd/mm/yyyy")
optionInbox = True
End Sub

Private Sub txtDate_Click()
frmSingleCalender.Show
End Sub
