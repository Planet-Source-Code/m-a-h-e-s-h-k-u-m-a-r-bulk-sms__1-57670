VERSION 5.00
Begin VB.Form frmGroupEdit 
   Caption         =   "Edit Group"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   Icon            =   "frmGroupEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      Begin VB.TextBox txtGroup 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtDescription 
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Group name"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveGroup 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strGroupName As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSaveGroup_Click()
    
    If txtDescription = "" Then
    txtDescription = "Unknown"
    End If
    If txtGroup = "" Then
    MsgBox "Group name cannot be empty", vbInformation
    Exit Sub
    End If
    SaveGroup txtGroup, txtDescription
End Sub

Private Sub Form_Load()
    txtGroup.Text = frmChild.bpp_tree.SelectedItem.Text
'    frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False

End Sub
Private Sub SaveGroup(GroupName As String, ByVal Description As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    If Trim(strGroupName) = Trim(GroupName) Then
        strQuery = "update GROUPS set GROUPDESC = '" & Description & "' where GROUPNAME = '" & strGroupName & "'"
        con.Execute strQuery
        MsgBox "Saved group information", vbInformation
        frmChild.RefreshTree
        Unload Me
    Else
        strQuery = "select GROUPNAME from GROUPS where GROUPNAME = '" & GroupName & "'"
        Set rs = New ADODB.Recordset
        rs.Open strQuery, con, 3, 2, 1
        If Not rs.EOF Then
            If GroupName = rs.Fields("GROUPNAME") Then
                MsgBox "Group name already exists, choose another", vbInformation
            End If
        Else
            strQuery = "update GROUPS set GROUPNAME = '" & GroupName & "', GROUPDESC = '" & Description & "' where GROUPNAME = '" & strGroupName & "'"
            rs.Close
            con.Execute strQuery
            MsgBox "Saved group information", vbInformation
            frmChild.RefreshTree
            Unload Me
        End If
    End If
    
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub
