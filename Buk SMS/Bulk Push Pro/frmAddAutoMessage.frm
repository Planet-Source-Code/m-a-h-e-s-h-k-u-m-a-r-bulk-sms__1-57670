VERSION 5.00
Begin VB.Form frmAddAutoMessage 
   Caption         =   "Add Auto Message"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   Icon            =   "frmAddAutoMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      Begin VB.TextBox txtKeyword 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtAutoMessage 
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Keyword"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Auto Message"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
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
End
Attribute VB_Name = "frmAddAutoMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    If txtAutoMessage = "" Then
        MsgBox "Reply Cannot be empty"
    End If
    AddAutoMessage (txtKeyword)
    
    frmMain.Enabled = True
    frmMain.Toolbar1.Enabled = True
End Sub
Private Sub AddAutoMessage(ByVal Keyword As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    If Keyword = "" Then
        MsgBox "Keyword cannot be empty", vbInformation
    Else
        Set rs = New ADODB.Recordset
        
        strQuery = "select KEYWORD from AUTOMESSAGE where KEYWORD = '" & Keyword & "'"
        rs.Open strQuery, con, 3, 2, 1
        If rs.RecordCount Then
            MsgBox "A Keyword already exists with '" & Keyword & "', please choose another", vbCritical
        Else
            rs.Close
            strQuery = "insert into AUTOMESSAGE(KEYWORD, AUTOMESSAGE, CREATEDON) values ('" & Keyword & "', '" & txtAutoMessage & "', '" & Now & "')"
            con.Execute strQuery
            Set rs = Nothing
            frmChild.LoadAutoMessages
            Unload Me
        End If
    End If
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True


End Sub

Private Sub cmdCancel_Click()
Unload Me
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub

Private Sub Form_Load()
'frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub
