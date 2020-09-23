VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmSingleCalender 
   Caption         =   "Click arrows to select month and year then click the day."
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4530
   Icon            =   "frmsingleCalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2280
      TabIndex        =   0
      Top             =   3240
      Width           =   1170
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   420
      Left            =   840
      TabIndex        =   1
      Top             =   3240
      Width           =   1170
   End
   Begin VB.TextBox TextBoxCalendar1 
      Height          =   330
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4440
      Width           =   2070
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2910
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4275
      _Version        =   524288
      _ExtentX        =   7541
      _ExtentY        =   5133
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   10
      Day             =   8
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Select Date"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSingleCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StartDate
Dim EndDate


Private Sub Calendar1_Click()
TextBoxCalendar1.Text = Format(Calendar1.Value, "Long Date")
End Sub

Private Sub cmdCancel_Click()
EndDate = -1
Unload Me
End Sub

Private Sub cmdOk_Click()
EndDate = Calendar1.Value
frmDailyReports.txtdate = Format(Calendar1.Value, "dd/mm/yyyy")
'frmScheduler.txtdate = Calendar1.Value
frmSendSingleMessage.txtdate = Calendar1.Value
frmSend.txtdate = Calendar1.Value
Unload Me
End Sub

Private Sub Form_Load()
Calendar1.Value = Now
StartDate = -1
EndDate = -1
End Sub


