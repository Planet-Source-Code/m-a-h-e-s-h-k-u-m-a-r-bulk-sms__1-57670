VERSION 5.00
Begin VB.Form frmPhonebookReport 
   Caption         =   "Phone Book Report"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin VB.ComboBox cmbGroups 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Text            =   "friends"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   " All Phone Book"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "  Group Name"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   450
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmPhonebookReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1 Then
generateReport LCase(cmbGroups.Text)
End If
If Option2 Then
generateAllReport
End If
End Sub

Private Sub generateReport(ByVal grpName As String)
Dim srchflag
Dim tempdate
        srchflag = False
    
    'txtDate = Format(txtDate, "dd/mm/yyyy")
    tempdate = Format(txtdate, "dd/mm/yyyy")
    Debug.Print txtTodate
    'If Check1 Then
        Set rs = New ADODB.Recordset
               strQuery = "select * from CONTACTS WHERE GROUPNAME = '" & grpName & "'"
        Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                DataReport1.Show
                'rpMessages.SetFocus
               frmMain.Enabled = True
                srchflag = True

            End If
    
If srchflag = False Then MsgBox "No Records Found"
 
   
End Sub

Private Sub generateAllReport()
Dim srchflag
Dim tempdate
        srchflag = False
    
    'txtDate = Format(txtDate, "dd/mm/yyyy")
    tempdate = Format(txtdate, "dd/mm/yyyy")
    Debug.Print txtTodate
    'If Check1 Then
        Set rs = New ADODB.Recordset
               strQuery = "select * from CONTACTS"
        Debug.Print strQuery
                rs.Open strQuery, con, 3, 2, 1

            If rs.EOF <> True Then
                DataReport1.Show
                'rpMessages.SetFocus
               frmMain.Enabled = True
                srchflag = True

            End If
    
If srchflag = False Then MsgBox "No Records Found"
 
   
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Option1 = True
Dim rs As ADODB.Recordset
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
    
    cmbGroups.ListIndex = 0
    
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

'frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False













End Sub
