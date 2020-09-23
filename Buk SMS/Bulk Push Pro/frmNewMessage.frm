VERSION 5.00
Begin VB.Form frmNewMessage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "New Mesage Recieved"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   915
         TabIndex        =   1
         Top             =   315
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmNewMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

