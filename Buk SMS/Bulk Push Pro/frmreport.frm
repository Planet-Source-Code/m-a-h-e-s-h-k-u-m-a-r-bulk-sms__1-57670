VERSION 5.00
Begin VB.Form frmReport 
   AutoRedraw      =   -1  'True
   Caption         =   " Dynamic Report Generation"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report Generation"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.OptionButton optionAll 
         Caption         =   "All Messages"
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   3120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Number"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Name"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNumber 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   915
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton cmdCalender 
         Caption         =   "&Calender"
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtToDate 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdInbox 
         Caption         =   "&Inbox"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdOutbox 
         Caption         =   "&Outbox"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdSentMessage 
         Caption         =   "&Sent Messages"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtFromDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "(mm-dd-yyyy)"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "(mm-dd-yyyy)"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "&From Date"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "&To Date"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Fields As Integer
Dim table As String
Dim srchflag As Boolean

Private Sub cmbbox_Change()
Label1.Caption = cmbbox.Text

End Sub

Private Sub cmbbox_Click()
Label1.Caption = cmbbox.Text
   If cmbbox = "All" Then
   Label1.Caption = ""
   txtitem.Enabled = False
   Else
   txtitem.Enabled = True
   End If
End Sub

Private Sub cmdCalender_Click()
    
    frmCalendar.Show
    
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdInbox_Click()
Fields = 0
    CheckFields
    If Fields = 0 Then Exit Sub
   
        If Option1 Then
            generateNameReport ("INBOX")
            
         End If
            
         If Option2 Then
             generateNumberReport ("INBOX")
            End If
     
     If srchflag = False Then Exit Sub
    
    
     
     
     
End Sub

Private Sub cmdOutbox_Click()
Fields = 0
    CheckFields
        If Fields = 0 Then Exit Sub

          If Option1 Then
            generateNameReport ("OUTBOX")
            
         End If
            
         If Option2 Then
             generateNumberReport ("OUTBOX")
            End If
            
        If srchflag = False Then Exit Sub
            
    
End Sub

Private Sub cmdSentMessage_Click()
Fields = 0
    CheckFields
        If Fields = 0 Then Exit Sub
            If Option1 Then
                generateNameReport ("SENTMESSAGES")
            
            End If
            
            If Option2 Then
                 generateNumberReport ("SENTMESSAGES")
            End If
            
     If srchflag = False Then Exit Sub
            
    
End Sub

Private Sub CheckFields()
    If Option1 Then
        If txtName = "" Then
            MsgBox "Enter Name"
         fieilds = 0
        Exit Sub
        End If
    End If
    
    If Option2 Then
        If txtNumber = "" Then
            MsgBox "Enter Number"
         Fields = 0
         Exit Sub
       End If
    End If
        If Check1 Then
            If txtFromDate = "" Or txtToDate = "" Then
                MsgBox "Choose from and to Dates"
                Fields = 0
            Exit Sub
            End If
        End If
 Fields = 1
End Sub



Private Sub Form_Load()
'frmMain.Enabled = False
Option1 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub Option1_Click()
        txtNumber.Enabled = False
        txtName.Enabled = True
        txtNumber.Text = ""
        Option1 = True
End Sub

Private Sub Option2_Click()
        txtNumber.Enabled = True
        txtName.Enabled = False
        txtName.Text = ""
       Option2 = True
End Sub

Private Sub optionAll_Click()
    optionAll = True
End Sub

Private Sub txtName_Change()
    Option1 = True
End Sub

Private Sub txtNumber_Change()
    Option2 = True
End Sub
Private Sub generateNameReport(ByVal table As String)


        srchflag = False
    
    txtToDate = Format(txtToDate, "mm/dd/yyyy")
    txtFromDate = Format(txtFromDate, "mm/dd/yyyy")
    Debug.Print txtToDate
    If Check1 Then
        Set rs = New ADODB.Recordset
               strQuery = "select * from " & table & " where ((format(TIME_STAMP,'mm/dd/yyyy') >=  '" & txtFromDate & "' and format(TIME_STAMP,'mm/dd/yyyy')<='" & txtToDate & "')) and NAME='" & txtName & "'"
        'Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                rpMessages.Show
                'rpMessages.SetFocus
               frmMain.Enabled = True
                srchflag = True

            End If
    Else
    
    Set rs = New ADODB.Recordset
               strQuery = "select * from " & table & " where NAME = '" & txtName.Text & " '"
                Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                'rpMessages.
                rpMessages.Show
              '  rpMessages.SetFocus
                  frmMain.Enabled = True
                srchflag = True

            End If
    
    
    
   End If
     If srchflag = False Then 'Display msg when search not found
        MsgBox "Search Not Found", vbInformation, "Search Result"
        Exit Sub
    End If
End Sub
Private Sub generateNumberReport(ByVal table As String)
    
        srchflag = False
    
    txtToDate = Format(txtToDate, "mm/dd/yyyy")
    txtFromDate = Format(txtFromDate, "mm/dd/yyyy")
    'Debug.Print txtToDate
    If Check1 Then
        Set rs = New ADODB.Recordset
               strQuery = "select * from " & table & " where ((format(TIME_STAMP,'mm/dd/yyyy') >=  '" & txtFromDate & "' and format(TIME_STAMP,'mm/dd/yyyy')<='" & txtToDate & "')) and MOBILENO='" & txtName & "'"
        'Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                rpMessages.Show
               ' rpMessages.SetFocus
                frmMain.Enabled = True
                srchflag = True

            End If
    Else
    
    
    
                Set rs = New ADODB.Recordset
               strQuery = "select * from " & table & " where MOBILENO ='" & txtNumber.Text & "'"
                rs.Open strQuery, con, 3, 2, 1
            If rs.EOF <> True Then
                rpMessages.Show
                rpMessages.Show
               ' rpMessages.SetFocus
               frmMain.Enabled = True
                srchflag = True
                  
            End If
         
      End If
      If srchflag = False Then 'Display msg when search not found
        MsgBox "Search Not Found", vbInformation, "Search Result"
        Exit Sub
    End If
      
End Sub
