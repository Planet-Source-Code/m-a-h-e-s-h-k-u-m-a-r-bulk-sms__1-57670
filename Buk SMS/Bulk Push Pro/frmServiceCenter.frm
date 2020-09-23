VERSION 5.00
Begin VB.Form frmServiceCenter 
   Caption         =   "Service Center Number "
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtMessage 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Service Center Number"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmServiceCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Changenumber

Private Sub Command1_Click()
frmChild.MSComm1.Output = "AT+CSCA=" + txtMessage + vbCr
Changenumber = 1
frmChild.GetServiceNumber
Unload Me
End Sub

Private Sub Command2_Click()

Changenumber = 0
Unload Me
End Sub

