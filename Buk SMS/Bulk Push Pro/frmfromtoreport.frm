VERSION 5.00
Begin VB.Form frmfromtoreport 
   Caption         =   "Report"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose date"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   " Failed Messages"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   " Sent Messages"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   780
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "  Inbox"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtTodate 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   2370
         Width           =   1215
      End
      Begin VB.TextBox txtfromdate 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   2370
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmfromtoreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option1 Then
    generateReport "INBOX"
    End If
    
    If Option2 Then
    generateReport "SENTMESSAGES"
    End If
    
    If Option3 Then
    generateReport "FAILEDMESSAGES"
    End If


End Sub





Private Sub generateReport(ByVal table As String)

On Error Resume Next
        srchflag = False
    
    txtTodate = Format(txtTodate, "mm/dd/yyyy")
    txtfromdate = Format(txtfromdate, "mm/dd/yyyy")
    Debug.Print txtTodate
    
        Set rs = New ADODB.Recordset
               strQuery = "select * from " & table & " where ((format(TIME_STAMP,'mm/dd/yyyy') >=  '" & txtfromdate & "' and format(TIME_STAMP,'mm/dd/yyyy')<='" & txtTodate & "')) "
        'Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                rpMessages.Show
                'rpMessages.SetFocus
               frmMain.Enabled = True
                srchflag = True
  
    
            End If
     If srchflag = False Then 'Display msg when search not found
        MsgBox "Search Not Found", vbInformation, "Search Result"
        Exit Sub
    End If
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub txtfromdate_Change()
frmCalendar.Show
End Sub

Private Sub txtfromdate_Click()
frmCalendar.Show
End Sub

Private Sub txtTodate_Change()
frmCalendar.Show
End Sub
