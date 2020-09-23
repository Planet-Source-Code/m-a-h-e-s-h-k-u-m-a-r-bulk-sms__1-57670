VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   2880
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000001&
         Caption         =   "Enter Username and Password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000001&
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000001&
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000001&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   2280
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload frmMain
Unload Me
End Sub

Private Sub Command2_Click()

If Text1.Text = "" Then Text1.SetFocus: Exit Sub
If Text2.Text = "" Then Text2.SetFocus: Exit Sub

If LCase(Text1.Text) = "admin" And Text2.Text = "12345" Then
    frmMain.Enabled = True
    Unload Me
  Else
 Label4.Caption = "Invalid Login ID/Password"
 End If
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
frmMain.Enabled = False
End Sub


