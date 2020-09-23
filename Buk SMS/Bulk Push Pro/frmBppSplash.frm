VERSION 5.00
Begin VB.Form frmBppSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   0
      Picture         =   "frmBppSplash.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   2760
   End
End
Attribute VB_Name = "frmBppSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Load()
'Timer1.Enabled = True
'
'End Sub
'
'Private Sub Timer1_Timer()
'Unload Me
'frmChild.Show
'End Sub
