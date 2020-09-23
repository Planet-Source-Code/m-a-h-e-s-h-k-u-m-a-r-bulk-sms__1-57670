VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   Caption         =   "Click arrows to select month and year then click a day on calendar."
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8790
   Icon            =   "frmcalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextBoxCalendar2 
      Height          =   315
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   150
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.TextBox TextBoxCalendar1 
      Height          =   315
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   2460
   End
   Begin MSACAL.Calendar Calendar2 
      Height          =   2685
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Width           =   4005
      _Version        =   524288
      _ExtentX        =   7064
      _ExtentY        =   4736
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   2685
      Left            =   -30
      TabIndex        =   4
      Top             =   240
      Width           =   4005
      _Version        =   524288
      _ExtentX        =   7064
      _ExtentY        =   4736
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4005
      TabIndex        =   3
      Top             =   1680
      Width           =   810
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   420
      Left            =   4005
      TabIndex        =   2
      Top             =   945
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "To This Date:"
      Height          =   270
      Left            =   4860
      TabIndex        =   1
      Top             =   195
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "From This Date:"
      Height          =   270
      Left            =   -15
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StartDate As Date
Public EndDate As Date

Private Sub Calendar1_Click()
TextBoxCalendar1.Text = Format(Calendar1.Value, "mm/dd/yyyy")
End Sub

Private Sub Calendar2_Click()
TextBoxCalendar2.Text = Format(EndDate, "mm/dd/yyyy")
End Sub

Private Sub cmdCancel_Click()
StartDate = -1
EndDate = -1
Unload Me
End Sub

Private Sub cmdOk_Click()
StartDate = Calendar1.Value
EndDate = Calendar2.Value
frmReport.txtfromdate = frmCalendar.StartDate
frmReport.txtTodate = frmCalendar.EndDate
frmfromtoreport.txtfromdate = frmCalendar.StartDate
frmfromtoreport.txtTodate = frmCalendar.EndDate
Unload Me
End Sub


Private Sub Form_Load()
Calendar1.Value = Now
Calendar2.Value = Now
StartDate = -1
EndDate = -1
End Sub


