VERSION 5.00
Begin VB.Form frmAddEditGroup 
   Caption         =   "Add Group"
   ClientHeight    =   3315
   ClientLeft      =   5385
   ClientTop       =   3375
   ClientWidth     =   4680
   Icon            =   "frmAddEditGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtDescription 
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtGroup 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Group name"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmAddEditGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

   If txtDescription = "" Then
   txtDescription = "Unknown"
   End If
    AddGroup txtGroup
    frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub

Private Sub cmdCancel_Click()
    frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True


    Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub AddGroup(ByVal GroupName As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    If GroupName = "" Then
        MsgBox "Group name cannot be empty", vbInformation
        Exit Sub
    Else
        Set rs = New ADODB.Recordset
        
        strQuery = "select GROUPNAME from GROUPS where GROUPNAME = '" & GroupName & "'"
        rs.Open strQuery, con, 3, 2, 1
        If rs.RecordCount Then
            MsgBox "A group already exists with '" & GroupName & "', please choose another", vbCritical
        Else
            rs.Close
            strQuery = "insert into GROUPS(GROUPNAME, GROUPDESC, CREATEDON) values ('" & GroupName & "', '" & txtDescription & "', '" & Now & "')"
            con.Execute strQuery
            Set rs = Nothing
            frmChild.RefreshTree
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()
'frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub
