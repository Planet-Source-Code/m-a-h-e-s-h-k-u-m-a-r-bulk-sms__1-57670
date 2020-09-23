VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmChild 
   AutoRedraw      =   -1  'True
   Caption         =   "BPP"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   FontTransparent =   0   'False
   Icon            =   "frmChild.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   600
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   120
      Top             =   4800
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contact Details"
      Height          =   2175
      Left            =   8160
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Lemail 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Ldesignition 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Lmobile 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Lname 
         Caption         =   "dd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Designition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Mobile no"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox txtrecivemessage 
      Height          =   375
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmChild.frx":08CA
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTelephone 
      Height          =   285
      Left            =   960
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList imglist 
      Left            =   4680
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":0D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1874
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   1575
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   195
      Left            =   7440
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.ListView bpp_list 
      Height          =   1455
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "imglist"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
         ImageIndex      =   3
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":1F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":22C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2614
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2966
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":320A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":375C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":3AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":3F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":4356
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":47AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":4C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":54E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":5806
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":5C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":5F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":6FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":78A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":7CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":85D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":8A28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView bpp_tree 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6588
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   139
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   3
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin MSComctlLib.ImageList imglTlbr 
      Left            =   3720
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":9302
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":99FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":A0FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":ADD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":B4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":B926
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":BD7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":C1CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   6720
      TabIndex        =   4
      Top             =   8040
      Width           =   6255
      Begin VB.Label LabelDesign 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label LabelTime 
         BackColor       =   &H8000000A&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label LabelFrom 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      FillColor       =   &H00008000&
      ForeColor       =   &H8000000F&
      Height          =   4815
      Left            =   5880
      ScaleHeight     =   4755
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   6240
      Width           =   9375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmval As Integer
Dim z As Integer
Dim captn, ch, part
Public Baudrate, Set_parity, Brate, Com_count, fnam1
Dim AT_flag, AT_flag1, CMGF_flag, Data_check, Voice_check, SMS_check, Inet_check, config_check, port_check As Boolean
Public ATFLAG As String
Dim tempCombuffer As String
Dim tempComInput$
Dim key
Dim FromModem$
Dim msgBreak()                 As String
Dim msgHeader()                As String
Dim MessageStatus As String
Dim time_stamp As String
Dim Message As String
Dim Mobile_Number As String
Dim MessageCount As String
Public Message_sent As String
Public temptime


'Dim fsys As New FileSystemObject
Dim f1
Public bEcho                   As Boolean
Public bOK                     As Boolean
Public bRing                   As Boolean
Public bError                  As Boolean
Public iRingTime               As Single
Public FirstRun                As Boolean
Public bErrorComm              As Boolean
Public bGreaterSign            As Boolean
Public bMessageStore           As Boolean
Public strMessageBuffer        As String
Public FileNumber              As Integer

Dim ndBPP As Node
Dim ndAll_GROUP As Node
Dim ndGRP As Node
Dim ndCONT As Node
Dim ndINB As Node
Dim ndOUT As Node
Dim ndSNT As Node
Dim ndAUT As Node
Dim AUT_MSG As Node
Dim ndDEV As Node
Dim ndREP As Node
Dim ndSCE
Dim ndFAILED
Dim ndREPMON As Node
Dim ndREPDAY As Node
Private BTN As Integer
Public Node_key As String
Public Node_tag As String
Dim A As ListItem
Public SignalValue As Integer
Public SignalStrenth As String
Dim keypress As Integer
Dim form_load_check As String
Public Deleted
Public tempBuffer
Dim tempBuff$





Private Sub bpp_list_Click()
 'bpp_list.Checkboxes = 1

End Sub

Private Sub bpp_list_DblClick()
 'bpp_list.SelectedItem.Checked = True

End Sub

Private Sub bpp_list_GotFocus()
' If bpp_list.SelectedItem.Checked = False Then
' bpp_list.SelectedItem.Checked = True
'  Exit Sub
'  End If
' If bpp_list.SelectedItem.Checked = True Then
' bpp_list.SelectedItem.Checked = False
' Exit Sub
' End If
End Sub


Private Sub bpp_list_ItemClick(ByVal item As MSComctlLib.ListItem)
    Dim D() As String
    Dim A As ListItem
    Dim temp
    'MsgBox bpp_tree.SelectedItem.Tag
    On Error Resume Next
    frmChild.bpp_list.ToolTipText = frmChild.bpp_list.SelectedItem.ListSubItems.item(3).Text
 On Error Resume Next
 


Node_key = bpp_tree.SelectedItem.key
    frmMain.cmdDelete.Enabled = False
    'frmMain.Toolbar1.Buttons.item(3).Enabled = False

On Error Resume Next

'Debug.Print bpp_tree.SelectedItem.key

temp = InStr(bpp_tree.SelectedItem.key, "GRP")
handler2:
If BTN = 2 Then
BTN = 0
If temp = 1 Then PopupMenu frmMain.mnulstContactRightClick
Select Case bpp_tree.SelectedItem.key
   
      
            
       Case "All_GROUP"
         'PopupMenu frmMain.mnuGroupMenu
       Case "INB"
            
         PopupMenu frmMain.mnuInboxMessage
       Case "OUT"
       PopupMenu frmMain.mnuOutMessage
       Case "SNT"
       PopupMenu frmMain.mnuSentmessages
       Case "AUT"
       frmMain.mnuAutoRemoveKeyword.Enabled = True
       PopupMenu frmMain.mnuAutoReply
       Case "DEV"
       MsgBox "Right click Listview Device"
       Case "SCE"
        PopupMenu frmMain.mnuSchedulList
       Case "REP"
       MsgBox "right click Listview Report"
       Case "FAI"
       PopupMenu frmMain.mnuFailedMessages
       Case "REPDLY"
       MsgBox "right click Listview Daily"
       Case "REPMON"
       MsgBox "Right click on"
  
    End Select
  End If
  On Error GoTo handler1
        If BTN = 1 Then
                BTN = 0
       ' On Error Resume Next
        bpp_tree.SetFocus
         Select Case bpp_tree.SelectedItem.key
              
'       Case Split(check_node.Tag, "|")(0)
'         PopupMenu frmMain.mnuGroupMenu
       Case "All_GROUP"
         'PopupMenu frmMain.mnuGroupMenu
       Case "INB"
           ShowMessage
        ' PopupMenu frmMain.mnuInboxMessage
       Case "OUT"
            ShowMessage
       'PopupMenu frmMain.mnuOutMessage
       Case "SNT"
            ShowMessage
       'PopupMenu frmMain.mnuSentmessages
       Case "AUT"
           ShowAutoKeywords
           frmMain.mnuAutoRemoveKeyword.Enabled = True
        Case "FAI"
            ShowMessage
       'PopupMenu frmMain.mnuAutoReply
       Case "SCE"
       ShowMessage
       Case "DEV"
       'MsgBox "Right click Listview Device"
       Case "REP"
       'MsgBox "right click Listview Report"
       Case "REPDLY"
      ' MsgBox "right click Listview Daily"
       Case "REPMON"
      ' MsgBox "Right click on"
    Case Split(bpp_tree.SelectedItem.key, "|")(0)
    Node_tag = bpp_tree.SelectedItem.Tag
    CheckNode1 bpp_tree.SelectedItem
     End Select
   End If
   

On Error GoTo finsh
'bpp_tree.SetFocus
  
finsh:
'bpp_list.SelectedItem.Checked = False
handler1:
End Sub

Private Sub bpp_list_KeyDown(KeyCode As Integer, Shift As Integer)

key = KeyCode

'
'    If KeyCode = 40 Then
'        'frmChild.bpp_list.SetFocus
'
'    End If
'    If KeyCode = 116 Then
'    frmChild.bpp_list.Refresh
'
'    End If
''MsgBox KeyCode
End Sub

Private Sub bpp_list_LostFocus()
'bpp_list.SelectedItem.Checked = False
End Sub

Private Sub bpp_list_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BTN = Button
''''''
''''''If bpp_list.SelectedItem.Checked = True Then
''''''bpp_list.SelectedItem.Checked = False
''''''Exit Sub
''''''End If
''''''If bpp_list.SelectedItem.Checked = False Then
''''''bpp_list.SelectedItem.Checked = True
''''''Exit Sub
''''''End If
End Sub

Private Sub bpp_list_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    bpp_list.SelectedItem.Checked = False
End Sub

Private Sub bpp_tree_DblClick()
    If bpp_tree.SelectedItem.key = "DEV" Then
        frmSettings.Show
    End If
    If bpp_tree.SelectedItem.key = "AUT" Then
        'If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
        'frmReadMessages.Show
    End If
    If InStr(bpp_tree.SelectedItem.key, "CONT") Then
      Dim rs As ADODB.Recordset
    Dim strQuery As String
        Set rs = New ADODB.Recordset
        strQuery = "select MOBILE from CONTACTS where CONTACTNAME = '" & frmChild.bpp_tree.SelectedItem & "'"
        rs.Open strQuery, con, 3, 2, 1
        frmSendSingleMessage.Text1.Text = rs.Fields("MOBILE")
        rs.Close
        
    frmSendSingleMessage.Show
    
    
    
    End If
End Sub

Private Sub bpp_tree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BTN = Button
End Sub



Private Sub bpp_tree_NodeClick(ByVal Node As MSComctlLib.Node)

Dim D() As String
Dim A As ListItem
'MsgBox bpp_tree.SelectedItem.Tag
Node_key = Node.key
Select Case Node.key
    
    Case "BPP"
    frmMain.cmdDelete.Enabled = False
    Frame1.Visible = False
'    MsgBox "Bulk Push pro"
    Case "All_GROUP"
       Label4.Caption = "Groups"
       frmMain.StatusBar.Panels(1).Text = "Groups"
      frmMain.Toolbar1.Buttons(4).Enabled = False
      frmMain.Toolbar1.Buttons(3).Enabled = True
      If BTN = 2 Then
        PopupMenu frmMain.mnuAllGroups
    Else
        LoadListViewGroups
        'bpp_list.ColumnHeaders.Clear
        
        frmChild.bpp_list.Refresh
        frmChild.bpp_list.SetFocus
    End If
  '  MsgBox "All Groups"
    Case "INB"
    Label4.Caption = "Inbox"
    frmMain.StatusBar.Panels(1).Text = "Inbox"
    frmMain.Toolbar1.Buttons.item(4).Enabled = False
    frmMain.cmdDelete.Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    LoadInbox
    
    If BTN = 2 Then
        PopupMenu frmMain.mnuCommonMenu
    End If
'    MsgBox "Inbox"
    Case "OUT"
    frmMain.StatusBar.Panels(1).Text = "Outbox"
    Label4.Caption = "Outbox"
    frmMain.cmdDelete.Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
   
    LoadOutbox
    If BTN = 2 Then
    PopupMenu frmMain.mnuCommonMenu
    End If
  '  MsgBox "Outbox"
    Case "SNT"
    frmMain.StatusBar.Panels(1).Text = "Sent Messages"
    Label4.Caption = "Sent Messages"
    frmMain.cmdDelete.Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    
    LoadSentMessages
    If BTN = 2 Then
    PopupMenu frmMain.mnuCommonMenu
    End If
   ' MsgBox "Sent Messages"
    Case "AUT"
    frmMain.StatusBar.Panels(1).Text = "Auto Messages"
    Label4.Caption = "Auto Messages"
    frmMain.cmdDelete.Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    'bpp_list.Checkboxes = True
    LoadAutoMessages
    If BTN = 2 Then
    frmMain.mnuAutoRemoveKeyword.Enabled = False
    PopupMenu frmMain.mnuAutoReply
    End If
   ' MsgBox "Auto Reply"
    Case "DEV"
    frmMain.Toolbar1.Buttons(4).Enabled = False
    'MsgBox "Device Settings"
    Case "REP"
    MsgBox "Reports"
    Case "REPMON"
    If BTN = 2 Then
    PopupMenu frmMain.mnuCommonMenu
    End If
    MsgBox "Monthly"
'    DataReport1.Show
    Case "REPDAY"
    If BTN = 2 Then
    PopupMenu frmMain.mnuCommonMenu
    End If
    MsgBox "Daily"
    Case "SCE"
    frmMain.StatusBar.Panels(1).Text = "Scheduled Messages"
    If BTN = 2 Then
    PopupMenu frmMain.mnuCommonMenu
    End If
    Label4.Caption = "Scheduled Messages"
    LoadSchedule
    Case "FAI"
    LoadFailedMessages
    If BTN = 2 Then
    PopupMenu frmMain.mnuCommonMenu
    End If
    Label4.Caption = "Failed Messages"
    frmMain.StatusBar.Panels(1).Text = "Failed Messages"
    Case Split(Node.key, "|")(0)
    Node_tag = Node.Tag
    CheckNode Node
     
End Select



End Sub
Private Sub CheckNode(ByVal check_node As Node)
Dim D() As String

D = Split(check_node.Tag, "|")
    If D(0) = "GRP" Then
        frmMain.StatusBar.Panels(1).Text = bpp_tree.SelectedItem.Text
        Label4.Caption = bpp_tree.SelectedItem.Text
        frmMain.cmdDelete.Enabled = True
        'bpp_list.Checkboxes = True
        frmMain.Toolbar1.Buttons.item(3).Enabled = True
        frmMain.Toolbar1.Buttons.item(4).Enabled = True
        LoadListViewContacts check_node
        frmChild.bpp_list.Refresh
        'frmChild.bpp_list.SetFocus
        KeyBoard
'        frmMain.Toolbar1.Buttonmenu.
        If BTN = 2 Then
            PopupMenu frmMain.mnuGroupMenu
        End If
    End If
    If D(0) = "CONT" Then
        frmMain.cmdDelete.Enabled = True
  
        frmMain.StatusBar.Panels(1).Text = bpp_tree.SelectedItem.Parent.Text
        Label4.Caption = bpp_tree.SelectedItem.Text
        frmMain.Toolbar1.Buttons.item(4).Enabled = True
        
        LoadListViewContactDetails check_node
        frmChild.bpp_list.Refresh
        frmChild.bpp_tree.SetFocus
        If BTN = 2 Then
        PopupMenu frmMain.mnuContactMenu
        End If
    End If
End Sub

Public Sub RefreshTree()
    
    SaveGroupExpandState
    bpp_tree.Nodes.Clear
    
    Set ndBPP = bpp_tree.Nodes.Add(, , "BPP", "Bulk Push Pro", 21, 21)
        ndBPP.Bold = True
        ndBPP.Expanded = True
        
    Set ndAll_GROUP = bpp_tree.Nodes.Add("BPP", tvwChild, "All_GROUP", "All Groups", 7, 7)
        ndAll_GROUP.Bold = True
        ndAll_GROUP.Expanded = True
    
    LoadGroupsToAllGroups
    
    Set ndINB = bpp_tree.Nodes.Add("BPP", tvwChild, "INB", "Inbox", 8, 8)
        ndINB.Bold = True
        
    Set ndOUT = bpp_tree.Nodes.Add("BPP", tvwChild, "OUT", "Outbox", 9, 9)
        ndOUT.Bold = True
        
    Set ndSNT = bpp_tree.Nodes.Add("BPP", tvwChild, "SNT", "Sent Messages", 11, 11)
        ndSNT.Bold = True

    Set ndAUT = bpp_tree.Nodes.Add("BPP", tvwChild, "AUT", "Auto Reply Messages", 11, 11)
        ndAUT.Bold = True
        
'    Set ndDEV = bpp_tree.Nodes.Add("BPP", tvwChild, "DEV", "Device Settings", 13, 13)
'        ndDEV.Bold = True
    
'    Set ndREP = bpp_tree.Nodes.Add("BPP", tvwChild, "REP", "Daily Reports", 7, 7)
'        ndREP.Bold = True
    Set ndSCE = bpp_tree.Nodes.Add("BPP", tvwChild, "SCE", "Schedule Messages", 20, 20)
        ndSCE.Bold = True
     Set ndFAILED = bpp_tree.Nodes.Add("BPP", tvwChild, "FAI", "Failed Messages", 7, 7)
            ndFAILED.Bold = True
    LoadGroupExpandState
'    frmMain.StatusBar.Align = Left
    
End Sub


    


Private Sub cmdModemSetup_Click()

End Sub

Private Sub Form_Resize()
On Error GoTo handler
    bpp_tree.Move 50, 50, bpp_tree.Width, Me.ScaleHeight - 100
    bpp_list.Move bpp_tree.Left + bpp_tree.Width + 50, 50, Me.ScaleWidth - (150 + bpp_tree.Width), (Me.ScaleHeight - 100) / 2
    bpp_list.Top = Label4.Top + Label4.Height
    Picture1.Top = bpp_list.Top + bpp_list.Height
    Picture1.Left = bpp_list.Left
handler:
    If Err.Description = "Object variable or With block variable not set" Then
        Exit Sub
    End If
End Sub

Public Sub LoadGroupsToAllGroups()

    Dim rs As ADODB.Recordset
    Dim strQuery As String
    Dim strGroupName As String
    Dim iGroupId As Integer
    On Error GoTo handler
    Set rs = New ADODB.Recordset
    strQuery = "select * from GROUPS"
    rs.Open strQuery, con, 3, 2, 1
    While Not rs.EOF
        strGroupName = rs.Fields("GROUPNAME")
        iGroupId = rs.Fields("GROUPID")
        
        Set ndGRP = bpp_tree.Nodes.Add(ndAll_GROUP.key, tvwChild, "GRP" & iGroupId, strGroupName, 1, 2)
            ndGRP.Tag = "GRP|" & iGroupId
        LoadContactsToGroup ndGRP
'        ndGRP.Expanded = True
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
handler:
End Sub

Private Sub LoadContactsToGroup(ByVal s_ndGRP As Node)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    Dim strContactName As String
    Dim strDesignation As String
    Dim strMobile As String
    Dim strEmail As String
    Dim iContactId As Integer
    
    Set rs = New ADODB.Recordset
    strQuery = "select * from CONTACTS where GROUPNAME = '" & s_ndGRP & "'"
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
        strContactName = rs.Fields("CONTACTNAME")
        strMobile = rs.Fields("MOBILE")
        strDesignation = rs.Fields("DESIGNATION")
        strEmail = rs.Fields("EMAIL")
        iContactId = rs.Fields("CONTACTID")
        'MsgBox strContactFName & "  " & strContactMName & "  " & strContactLName & "  " & iContactId
        Set ndCONT = bpp_tree.Nodes.Add(s_ndGRP.key, tvwChild, "CONT" & iContactId, strContactName, 3, 4)
            ndCONT.Tag = "CONT|" & iContactId

       
        rs.MoveNext
    Wend
    
    rs.Close

End Sub
    


Public Sub LoadInbox()
Dim i As Integer
Dim temp As String
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    On Error Resume Next
    

    bpp_list.ListItems.Clear
    bpp_list.ColumnHeaders.Clear
    bpp_list.ColumnHeaders.Add , , "Name", 2000, , 5
    bpp_list.ColumnHeaders.Add , , "Number", 1500, , 14
    bpp_list.ColumnHeaders.Add , , "Message", 5000, , 18
    bpp_list.ColumnHeaders.Add , , "Date/Time"
    bpp_list.ColumnHeaders.Add , , "Inbox ID"
   ' bpp_list.ColumnHeaders.Add , , "Name"

       Set rs = New ADODB.Recordset
    strQuery = "select * from INBOX order by inboxid desc"
    rs.Open strQuery, con, 3, 2, 1
   i = 0
   While Not rs.EOF
                            
              
       Set A = bpp_list.ListItems.Add(, , rs.Fields("NAME"))
               A.SubItems(2) = rs.Fields("MESSAGE")
               A.SubItems(3) = rs.Fields("TIME_STAMP")
               A.SubItems(1) = rs.Fields("MOBILENO")
               A.SubItems(4) = rs.Fields("INBOXID")
              
               rs.MoveNext
        
    
    Wend
'bpp_tree.SelectedItem.key = "INB"
'bpp_list.Sorted = True
''
'bpp_list.SortKey = 3
'bpp_list.SortOrder = lvwDescending
End Sub
    
Public Sub LoadOutbox()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   On Error Resume Next
    bpp_list.ListItems.Clear
    bpp_list.ColumnHeaders.Clear
    bpp_list.ColumnHeaders.Add , , "To", 2000, , 5
    bpp_list.ColumnHeaders.Add , , "Number", 1500, , 14
    bpp_list.ColumnHeaders.Add , , "Message", 5000, , 18
    bpp_list.ColumnHeaders.Add , , "Date/Time"
    bpp_list.ColumnHeaders.Add , , "Out Box ID"
        
    Set rs = New ADODB.Recordset
    strQuery = "select * from OUTBOX order by ID DESC"
    rs.Open strQuery, con, 3, 2, 1
   
    While Not rs.EOF
        
       Set A = bpp_list.ListItems.Add(, , rs.Fields("NAME"))
               A.SubItems(2) = rs.Fields("MESSAGE")
               A.SubItems(3) = rs.Fields("TIME_STAMP")
               A.SubItems(1) = rs.Fields("MOBILENO")
               A.SubItems(4) = rs.Fields("ID")
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
 
End Sub


Public Sub LoadSentMessages()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   On Error Resume Next
    bpp_list.ListItems.Clear
    bpp_list.ColumnHeaders.Clear
    bpp_list.ColumnHeaders.Add , , "To", 2000, , 5
    bpp_list.ColumnHeaders.Add , , "Number", 1500, , 14
    bpp_list.ColumnHeaders.Add , , "Message", 5000, , 18
    bpp_list.ColumnHeaders.Add , , "Date/Time"
    bpp_list.ColumnHeaders.Add , , "Sent Box ID"
        
    Set rs = New ADODB.Recordset
    strQuery = "select * from SENTMESSAGES order by ID DESC"
    rs.Open strQuery, con, 3, 2, 1
   
    While Not rs.EOF
        
       Set A = bpp_list.ListItems.Add(, , rs.Fields("NAME"))
               A.SubItems(2) = rs.Fields("MESSAGE")
               A.SubItems(3) = rs.Fields("TIME_STAMP")
               A.SubItems(1) = rs.Fields("MOBILENO")
               A.SubItems(4) = rs.Fields("ID")
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
'bpp_list.Sorted = True
'
'bpp_list.SortKey = 4
'bpp_list.SortOrder = lvwDescending
    
    
    
End Sub

Public Sub LoadAutoMessages()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   
    bpp_list.ListItems.Clear
    bpp_list.ColumnHeaders.Clear
    bpp_list.ColumnHeaders.Add , , "Keyword", 1000, , 19
    bpp_list.ColumnHeaders.Add , , "Auto Message", 4000, , 11
    bpp_list.ColumnHeaders.Add , , "Create on"
    bpp_list.ColumnHeaders.Add , , "ID"
    Set rs = New ADODB.Recordset
    strQuery = "select * from AUTOMESSAGE ORDER BY ID DESC"
    rs.Open strQuery, con, 3, 2, 1
   
    While Not rs.EOF
        
       Set A = bpp_list.ListItems.Add(, , rs.Fields("KEYWORD"))
               A.SubItems(1) = rs.Fields("AUTOMESSAGE")
               A.SubItems(2) = rs.Fields("CREATEDON")
               A.SubItems(3) = rs.Fields("ID")
           rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
  
'bpp_list.Sorted = True
'
'bpp_list.SortKey = 3
'bpp_list.SortOrder = lvwDescending
'
  
  
End Sub

Private Sub LoadListViewGroups()
    Dim i
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    bpp_list.ListItems.Clear
    
    bpp_list.ColumnHeaders.Clear
    
    bpp_list.ColumnHeaders.Add , , "Group Name"
    bpp_list.ColumnHeaders.Add , , "Description"
    bpp_list.ColumnHeaders.Add , , "Created on"
    bpp_list.ColumnHeaders.Add , , "Contacts Count", 2000
     
   ' bpp_list.ListItems.Item
    
    Set rs = New ADODB.Recordset
    strQuery = "select * from GROUPS"
    
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
       Set A = bpp_list.ListItems.Add(, , rs.Fields("GROUPNAME"))
            A.SubItems(1) = rs.Fields("GROUPDESC")
            A.SubItems(2) = rs.Fields("CREATEDON")
        rs.MoveNext
    Wend
    Set rs = Nothing
     
    For i = 1 To bpp_list.ListItems.Count
    Set rs = New ADODB.Recordset
    strQuery = "select * from CONTACTS where GROUPNAME = '" & bpp_list.ListItems.item(i).Text & "'"
    rs.Open strQuery, con, 3, 2, 1
        bpp_list.ListItems(i).SubItems(3) = rs.RecordCount
    Next i
    
End Sub

Public Sub LoadListViewContacts(ByVal s_ndGRP As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    On Error Resume Next
    bpp_list.ListItems.Clear
    
    bpp_list.ColumnHeaders.Clear
    
    bpp_list.ColumnHeaders.Add , , "Name", 2000, , 5
    bpp_list.ColumnHeaders.Add , , "Mobile No", 2000, , 14
    bpp_list.ColumnHeaders.Add , , "Designation", 2000, , 15
    bpp_list.ColumnHeaders.Add , , "Email", 2000
    'bpp_list.li
    
    Set rs = New ADODB.Recordset
    strQuery = "select * from CONTACTS where GROUPNAME = '" & s_ndGRP & "' ORDER BY CONTACTID DESC"
    
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
       Set A = bpp_list.ListItems.Add(, , rs.Fields("CONTACTNAME"))
            A.SubItems(1) = rs.Fields("MOBILE")
            A.SubItems(2) = rs.Fields("DESIGNATION")
            A.SubItems(3) = rs.Fields("EMAIL")
        rs.MoveNext
    Wend
End Sub

Private Sub LoadListViewContactDetails(ByVal check_node As Node)
    Dim Dt() As String
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    Dt() = Split(check_node.Tag, "|")
    
    bpp_list.ListItems.Clear
    
    bpp_list.ColumnHeaders.Clear
   ' bpp_list.Sorted = False
    bpp_list.ColumnHeaders.Add , , "Field", 2000, , 17
    bpp_list.ColumnHeaders.Add , , "Value", 2000, , 17
    
    Set rs = New ADODB.Recordset
    strQuery = "select * from CONTACTS where CONTACTID = " & Dt(1)
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
         Set A = bpp_list.ListItems.Add(, , "Name")
            A.SubItems(1) = rs.Fields("CONTACTNAME")
            A.Bold = True
       Set A = bpp_list.ListItems.Add(, , "Mobile No")
            A.Bold = True
            A.SubItems(1) = rs.Fields("MOBILE")
       
       
       
        Set A = bpp_list.ListItems.Add(, , "Email")
            A.SubItems(1) = rs.Fields("EMAIL")
            A.Bold = True
             
              
           
        Set A = bpp_list.ListItems.Add(, , "Designation")
            A.SubItems(1) = rs.Fields("DESIGNATION")
            A.Bold = True
         
            
           
            
        rs.MoveNext
    Wend
    
    
End Sub

Private Sub SaveGroupExpandState()
    Dim ndTemp  As Node
    Dim strQuery As String
    
    On Error GoTo handler
    strQuery = "delete from EXPANDSTATE"
    con.Execute strQuery
    
    
    Set ndTemp = ndAll_GROUP.Child
    While True
        If ndTemp.Expanded = True Then
            strQuery = "insert into EXPANDSTATE(GROUPNAME) values ('" & ndTemp & "')"
            con.Execute strQuery
        End If
'        MsgBox ndTemp & "  " & ndTemp.Index
        Set ndTemp = ndTemp.Next
    Wend
handler:
    If Err.Description = "Object variable or With block variable not set" Then
        Exit Sub
    Else
        'MsgBox Err.Description
    End If
    
End Sub

Private Sub LoadGroupExpandState()
    Dim ndTemp As Node
    Dim strQuery As String
    On Error GoTo handler
    
    Set ndTemp = ndAll_GROUP.Child
    While True
        If IsExpanded(ndTemp) Then
            ndTemp.Expanded = True
        Else
            ndTemp.Expanded = False
        End If
        Set ndTemp = ndTemp.Next
    Wend
    
handler:
    If Err.Description = "Object variable or With block variable not set" Then
        strQuery = "delete from EXPANDSTATE"
'        con.Execute strQuery
        Exit Sub
    Else
        MsgBox Err.Description
    End If
End Sub

Private Function IsExpanded(ByVal ndTemp As Node) As Boolean
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    Set rs = New ADODB.Recordset
    strQuery = "select GROUPNAME from EXPANDSTATE where GROUPNAME = '" & ndTemp & "'"
    rs.Open strQuery, con, 3, 2, 1
    If rs.RecordCount Then
        IsExpanded = True
    Else
        IsExpanded = False
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo handler
    SaveGroupExpandState
handler:
'Unload frmMain
'
'Unload ADODB
'Exit Sub
End Sub




Private Sub KeyBoard()

Select Case keypress
       
       Case 46
           If bpp_tree.SelectedItem.Tag = "" Then
           MsgBox "Cannot delete this item"
           Exit Sub
           End If
           If Split(bpp_tree.SelectedItem.Tag, "|")(0) = "GRP" Then
             frmMain.RemoveGroup bpp_tree.SelectedItem
           Else
           If Split(bpp_tree.SelectedItem.Tag, "|")(0) = "CONT" Then
           frmMain.RemoveContact bpp_tree.SelectedItem.Parent, bpp_tree.SelectedItem
           End If
           End If
      
       Case 45
                      
            If bpp_tree.SelectedItem.Tag = "" Then
                If bpp_tree.SelectedItem = "All Groups" Then
                     'frmAddEditGroup.StartUpPosition
                     End If

                Exit Sub
           End If
           
           If Split(bpp_tree.SelectedItem.Tag, "|")(0) = "GRP" Then
                 frmAddEditContact.SelectedNode = bpp_tree.SelectedItem
                 frmAddEditContact.Show
           Else
                If Split(bpp_tree.SelectedItem.Tag, "|")(0) = "CONT" Then
                     frmAddEditContact.SelectedNode = bpp_tree.SelectedItem.Parent
                     frmAddEditContact.Show
                End If
'                If bpp_tree.SelectedItem.Key = "AUT" Then
'                    frmAddAutoMessage.Show
'                End If
           End If
        
       Case Default
            keypress = 0
            
            
End Select


End Sub
Private Sub bpp_tree_KeyDown(KeyCode As Integer, Shift As Integer)
    
    keypress = KeyCode
    KeyBoard
    keypress = 0
    'bpp_tree.Refresh
    bpp_tree.SetFocus
End Sub


Private Sub Form_Load()
'Debug.Print Now
    'frmScheduler.LoadSchedule
    RefreshTree
    LoadGroupExpandState
   ' Initialise_Modem
   ' Text1.RightToLeft = True
    'Frame1.Visible = False
    txtMessage.RightToLeft = frmChild.RightToLeft
    form_load_check = "false"
  '  command3_Click
   ' delay (1)
   'form_load_check = "true"
    Command2_Click
  '  SendSms "9849706959", "Application start" &now
 End Sub





Private Sub ShowMessage()
    Label1.Caption = "Mobile:"
    Label2.Caption = "Date/Time:"
    Label3.Caption = "Message"
    Frame2.Visible = False
    Frame1.Visible = True
    On Error Resume Next
    LabelFrom.Caption = bpp_list.SelectedItem.Text
    LabelTime.Caption = Format(bpp_list.SelectedItem.ListSubItems.item(3).Text, "dd/mm/yyyy hh:mm AM/PM")
    If bpp_list.SelectedItem.Text = "Unknown" Then
    LabelFrom.Caption = bpp_list.SelectedItem.ListSubItems.item(1).Text
    End If
   ' DateTime.Time
    LabelDesign.Visible = False
    txtMessage.Visible = True
    txtMessage.Text = bpp_list.SelectedItem.ListSubItems.item(2).Text
    
End Sub
Private Sub ShowAutoKeywords()
    Label1.Caption = "Keyword:"
    Label2.Caption = "Time:"
    Label3.Caption = "Message"
    Frame1.Visible = True
    LabelFrom.Caption = bpp_list.SelectedItem.Text
    LabelTime.Caption = bpp_list.SelectedItem.ListSubItems.item(2).Text
    txtMessage.Text = bpp_list.SelectedItem.ListSubItems.item(1).Text
    
End Sub
Private Sub CheckNode1(ByVal check_node As Node)
Dim D() As String
 On Error Resume Next
D = Split(check_node.Tag, "|")
    If D(0) = "GRP" Then
        txtMessage.Visible = False
        
        ShowContactList
    End If
    'bpp_list.Checkboxes =
    
End Sub

Private Sub ShowContactList()
    Frame2.Visible = True
    Frame1.Visible = False
    If Not bpp_tree.SelectedItem = "Bulk Push Pro" Then
    Lname.Caption = bpp_list.SelectedItem.Text
    Lmobile.Caption = bpp_list.SelectedItem.ListSubItems.item(1).Text
    Ldesignition.Caption = bpp_list.SelectedItem.ListSubItems.item(2).Text
    Lemail.Caption = bpp_list.SelectedItem.ListSubItems.item(3).Text
    End If
End Sub



Public Sub Command2_Click()
On Error GoTo handler
Dim DialString$, FromModem$
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
If MSComm1.PortOpen = False Then
   'mscomm1.PortOpen = True
   MSComm1.DTREnable = True
   MSComm1.RTSEnable = True
   MSComm1.RThreshold = 1
   MSComm1.InputLen = 1
   MSComm1.Settings = "9600, n, 8, 1"
   bOK = False
   bError = False
   
    MSComm1.PortOpen = True
 '  MSComm1.Output = "AT" + vbCrLf
   End If
' Do
'   If mscomm1.InBufferCount Then FromModem$ = FromModem$ + mscomm1.Input
'             If InStr(FromModem$, "OK") Then
'             MsgBox "ready"
'             GoTo finish
'            End If
'
'  Loop
   'MsgBox "Port Already Open !", vbCritical + vbOKOnly, "Error opening port"
finish:
handler:
Debug.Print Err.Number
If Err.Number = 8005 Then
MsgBox "Another Program is using Modem"
Unload Me
Unload frmMain
Exit Sub
End If


'MsgBox day(
End Sub

Public Sub delay(dlytime As Variant)
Dim delay_timer
Dim dli
delay_timer = Timer
dli = 0
Do
dli = dli + 1
   If Timer > (delay_timer + dlytime) Then
      Exit Do
   End If
Loop
End Sub



Private Sub mscomm1_OnComm()
    Static stEvent             As String
    Dim stComChar               As String * 1
 tempBuff$ = ""

    Select Case MSComm1.CommEvent

        Case comEvReceive

            Do
                
               
                stComChar = MSComm1.Input
                tempBuffer = tempBuffer + stComChar
                If bMessageStore Then
                   strMessageBuffer = strMessageBuffer & stComChar
                End If
                Select Case stComChar
                    Case ">"
                         bGreaterSign = True
                         'lstEvents.AddItem stComChar
                    Case vbLf

                    Case vbCr
                        If Len(stEvent) > 0 Then
                          ProcessEvent stEvent
                          stEvent = ""
                        End If
                    Case Else
                        stEvent = stEvent + stComChar
                End Select

            Loop While MSComm1.InBufferCount
    
    Case 3
    MsgBox "Modem Unplugged", vbInformation
    End Select

End Sub

Private Sub ProcessEvent(stEvent As String)
  Dim stNumber As String
  
       'lstEvents.AddItem stEvent
        If Mid$(stEvent, 1, 5) = "+CMTI" Then
           frmSendSingleMessage.Enabled = False
           Timer1.Enabled = False
           txtTelephone.Text = ""
           txtMessage.Text = ""
           strMessageBuffer = ""
           frmNewMessage.Show
           While frmSend.SendingMessage = "yes"
           DoEvents
           Wend
              stEvent = ""
              Command3_Click
            LoadInbox
            bpp_list.Refresh
           Unload frmNewMessage
          
           bOK = False
           bError = False
           
           MSComm1.Output = "AT+CMGD=1,3" & vbCrLf
           While Not bOK Or bError
                 DoEvents
                 Wait
           Wend
           Timer1.Enabled = True
           frmSendSingleMessage.Enabled = True
           If bError Then
              MsgBox "Unable to delete"
           End If
           Exit Sub
        End If
        
          
        
        Select Case stEvent
           Case "OK"
             bOK = True
           Case "ERROR"
             bError = True
           Case "RING"
           MsgBox "Incoming Call Alert", vbInformation
             If bRing = False Then
               bRing = True
             End If
             iRingTime = Timer
           Case Else
             Select Case Left(stEvent, 4)
               Case "TIME"
               Case "DATE"
               Case "NMBR"
               Case "NAME"
             End Select
             
        End Select


End Sub

Public Sub Command3_Click()
Dim pos1, pos2
pos1 = 1
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
Call Command2_Click
bOK = False
bError = False
Timer1.Enabled = False
'Timer2.Enabled = False
frmSendSingleMessage.Enabled = False
MSComm1.Output = "AT+CMGL=" & Chr(34) & "ALL" & Chr(34) & vbCrLf

While Not bOK Or bError
  bMessageStore = True
  DoEvents
  Wait
Wend
If bOK Then

   Debug.Print strMessageBuffer
    If Len(strMessageBuffer) < 25 Then
    
    Exit Sub
    End If
    
  Timer1.Enabled = True
   frmSendSingleMessage.Enabled = True
   ReadMessage
   'MsgBox txtMessage.Text
   'MsgBox txtTelephone.Text
'  pos1 = InStr(txtMessage.Text, txtTelephone.Text, "/", vbTextCompare)
   'MsgBox txtMessage
 'AddToInbox txtTelephone, txtrecivemessage, time_stamp
'   If InStr(1, UCase(txtmessage.Text), "NOTEPAD", vbTextCompare) <> 0 Then
'      'Call ExecuteCommand("NotePad.exe")
'   ElseIf InStr(1, UCase(txtmessage.Text), "CALC", vbTextCompare) <> 0 Then
'     ' Call ExecuteCommand("Calc.exe")
'   End If

End If
If bError Then
   txtrecivemessage.Text = "Bad Read"
End If
DeleteMessages
End Sub
Public Sub Wait()
Dim start

   start = Timer
   Do While Timer < start + 8
      DoEvents
      If bOK Then
        Exit Sub
      End If
      If bError Then
        Exit Sub
      End If
   Loop
End Sub
Private Sub WaitLong()
Dim start

   start = Timer
   Do While Timer < start + 36
      DoEvents
      If bOK Then
        Exit Sub
      End If
      If bError Then
        Exit Sub
      End If
   Loop
End Sub
Private Sub ReadMessage()

Dim CMGLbreak() As String
Dim CMGLheader() As String
Dim j
Dim i
'strMessageBuffer = Text1.Text ''mahi ,ahidummy unread message
If ParseFile Then

        'Debug.Print strMessageBuffer
        
        
        strMessageBuffer = Mid(strMessageBuffer, 3, Len(strMessageBuffer))
             CMGLbreak = Split(strMessageBuffer, "+CMGL:", , vbTextCompare)
            CMGLheader = Split(CMGLbreak(0), ",", , vbTextCompare)
         On Error GoTo finish
    For j = 0 To 20
           'Debug.Print CMTIbreak(i)
          
            strMessageBuffer = CMGLbreak(j)
           
           msgBreak = Split(strMessageBuffer, vbCrLf, , vbTextCompare)
           msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
           
           
            msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
           txtTelephone.Text = Mid$(Right$(msgHeader(2), 11), 1, 10)
           strMessageBuffer = ""
         For i = 1 To UBound(msgBreak(), 1)
               If i = 2 Then
               Message = strMessageBuffer
               End If
               strMessageBuffer = strMessageBuffer & msgBreak(i) & vbCrLf
         '      Message = strMessageBuffer
        Next i
            MessageCount = Mid(msgBreak(0), 7, 3)
        If Mid(MessageCount, 1, 1) = Chr(34) Then
            MessageCount = Mid(MessageCount, 2, 2)
        End If
           MessageStatus = msgHeader(1)
           'MsgBox Len(msgHeader(1))
        If InStr(msgHeader(1), "UNREAD") Then
            MessageStatus = Mid(msgHeader(1), 2, Len(msgHeader(1)) - 2)
        Else
            'DeleteMessage MessageCount
        End If
           
           time_stamp = msgHeader(4) + "," + msgHeader(5)
           time_stamp = Mid(time_stamp, 2, Len(time_stamp) - 5)
           txtTelephone.Text = Mid$(Right$(msgHeader(2), 11), 1, 10)
'            If txtTelephone.Text = "" Then
'            StatusBar1.Panels(1).Text = " Waiting for Messages"
            'MsgBox "inbox empty"
'            Exit Sub
'
'        End If
            'MsgBox msgBreak(6)
            Mobile_Number = txtTelephone.Text
            If InStr(txtTelephone.Text, Chr(34)) Then
            txtTelephone = Mid(txtTelephone.Text, 2, Len(txtTelephone) - 2)
            Debug.Print txtTelephone
            End If
'        Debug.Print txtTelephone
'        Debug.Print Message
'        Debug.Print time_stamp
        AddToInbox txtTelephone, Message, time_stamp
'        Else
'           txtMessage.Text = "Unable to decode Message"
  Next j
End If

finish:

' DeleteMessages
 
End Sub
Public Function ParseFile() As Boolean
'strMessageBuffer = Text1.Text
Dim i
Dim FirstOffSet As Long
Dim SecondOffSet As Long
Dim strBuffer1 As String
Dim strBuffer2 As String
Dim strBuffer3 As String
strBuffer1 = strMessageBuffer
'Debug.Print strBuffer1
FirstOffSet = InStr(1, strBuffer1, "+CMGL:", vbTextCompare)
SecondOffSet = InStr(1, strBuffer1, vbCrLf & "OK", vbTextCompare)
If FirstOffSet <> 0 And SecondOffSet > FirstOffSet Then
   i = FirstOffSet
   While i < SecondOffSet
    strBuffer2 = strBuffer2 & Mid$(strBuffer1, i, 1)
    i = i + 1
   Wend
   ParseFile = True
   strMessageBuffer = strBuffer2
   Exit Function
End If
ParseFile = False
End Function
Private Sub DeleteMessages()

bOK = False
           bError = False
           Timer1.Enabled = False
           MSComm1.Output = "AT+CMGD=1,3" & vbCrLf
           While Not bOK Or bError
                 DoEvents
                 Wait
           Wend
           If bError Then
              MsgBox "Unable to delete"
           End If
           Timer1.Enabled = True
    Deleted = 1
End Sub
Private Sub AddToInbox(ByVal Mobileno As String, ByVal Message As String, ByVal TimeStamp As String)
Dim strQuery As String
Dim strTempTime As String
Dim tempName As String
Dim timeBreak()                 As String
Dim timeHeader()                As String
Dim temp_day As String
Dim temp_month As String
Dim temp_year As String

CheckAutoMessages Message, Mobileno



If Message = "" Then Message = "Blank"

    timeBreak = Split(TimeStamp, ",", , vbTextCompare)
    timeHeader = Split(timeBreak(0), ",", , vbTextCompare)
       
      temp_day = Mid$(timeBreak(0), 7, 2)
      temp_year = Mid$(timeBreak(0), 1, 2)
      temp_month = Mid$(timeBreak(0), 4, 2)
       strTempTime = Format(timeBreak(1), "am")
        TimeStamp = temp_month + "/" + temp_day + "/" + "20" + temp_year + " " + Mid$(timeBreak(1), 1, 8)
       
        'StatusBar1.Panels(1).Text = "Loading Message to Database...."
   On Error GoTo OpenError
   
   
            
            Set rs = New ADODB.Recordset
                strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & Mobileno & "'"
            
            rs.Open strQuery, con, 3, 2, 1
            
                tempName = rs.Fields("CONTACTNAME")
            
'        Set rs = Nothing
'            strQuery = "delete from inbox where "
'
'
OpenError:
If Err.Number <> 0 Then
  '  MsgBox Err.Number
   tempName = "Unknown"
   Resume Next
End If
       Set rs = Nothing
       
       Trim (Message)
           ' MsgBox "New message Recived", vbInformation
            strQuery = "insert into INBOX(MOBILENO,MESSAGE,TIME_STAMP,NAME,READ) values ('" & Mobileno & "', '" & Message & "', '" & TimeStamp & "','" & tempName & "','FALSE')"
       Debug.Print strQuery
       'InputBox "", "", strQuery
      ' On Error GoTo handler
       con.Execute strQuery
handler:
'       DeleteMessage (MessageCount)
End Sub




Public Sub SendSms(ByVal Number As String, ByVal Message As String)
Dim messagetosend As String
Debug.Print frmSend.SendingMessage
messagetosend = Message
  
        If Len(messagetosend) > 160 Then
            SendSms1 Number, Mid(messagetosend, 1, 160)
            'Debug.Print Mid(messagetosend, 160, 320)
            SendSms1 Number, Mid(messagetosend, 160, 320)
        Else
                SendSms1 Number, Message
    
    End If
    
    
End Sub


Public Sub SendSms1(ByVal Number As String, ByVal Message As String)
Dim i
i = 0
tempBuffer = ""
Dim start
 bGreaterSign = False
 Message_sent = "False"
 
 start = Timer
  MSComm1.Output = "AT+CMGS=" + Chr$(34) + Trim(Number) + Chr$(34) + vbCr
   While Not bGreaterSign
      DoEvents
      Wait
            If Timer > (start + 50) Then
                      MsgBox "Time Out"
                      Message_sent = "False"
                      Exit Sub
                  End If
      Wend
 
   If bGreaterSign Then
      MSComm1.Output = Trim(Message) + Chr$(26)
      start = ""
        start = Timer
      bOK = False
      bError = False
      While Not bOK Or bError
      
                      Debug.Print tempBuffer
                      If InStr(tempBuffer, "ERROR") Then
                      Message_sent = "False"
                      tempBuffer = ""
                      Exit Sub
                      End If
                      Debug.Print Timer
                Debug.Print start + 50
                If Timer > (start + 50) Then
                      MsgBox "Time out"
                      Message_sent = "False"
                      Exit Sub
                  End If
            
            
          DoEvents
          Wait
      Wend
       If InStr(tempBuffer, "ERROR") Then
        Message_sent = "False"
        tempBuffer = ""
        Exit Sub
        End If
     
      If bOK Then
        ' MsgBox "Message Sent", vbInformation + vbOKOnly, "Sent"
         Message_sent = "True"
         
      Else
         'MsgBox "Message Not Sent", vbCritical + vbOKOnly, "Cannot Send"
         Message_sent = "False"
      End If
   Else
      'MsgBox "Message cannot be sent", vbCritical + vbOKOnly, "Cannot Send"
      Message_sent = "NotPossible"
   End If
   If bError Then MsgBox " fAILED"


End Sub




Public Sub LoadSchedule()
Dim i As Integer
Dim temp As String
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    On Error Resume Next
    
    bpp_list.ListItems.Clear
    bpp_list.ColumnHeaders.Clear
    bpp_list.ColumnHeaders.Add , , "Name", 2000, , 5
    bpp_list.ColumnHeaders.Add , , "Number", 1500, , 14
    bpp_list.ColumnHeaders.Add , , "Message", 5000, , 18
    bpp_list.ColumnHeaders.Add , , "Date/Time"
    bpp_list.ColumnHeaders.Add , , "ID"
   ' bpp_list.ColumnHeaders.Add , , "Name"

       Set rs = New ADODB.Recordset
    strQuery = "select * from SCHEDULEMESSAGES ORDER BY ID DESC"
Debug.Print strQuery
    rs.Open strQuery, con, 3, 2, 1
   i = 0
   While Not rs.EOF
                            
                
               
                
       Set A = bpp_list.ListItems.Add(, , rs.Fields("NAME"))
               A.SubItems(2) = rs.Fields("MESSAGE")
               A.SubItems(3) = rs.Fields("SCHEDULETIME")
               A.SubItems(1) = rs.Fields("MOBILENO")
               A.SubItems(4) = rs.Fields("ID")
              
               rs.MoveNext
        
    
    Wend

'bpp_list.Sorted = True
'
'bpp_list.SortKey = 4
'bpp_list.SortOrder = lvwDescending
'
End Sub


Public Sub LoadFailedMessages()
Dim i As Integer
Dim temp As String
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    On Error Resume Next
    
    bpp_list.ListItems.Clear
    bpp_list.ColumnHeaders.Clear
    bpp_list.ColumnHeaders.Add , , "Name", 2000, , 5
    bpp_list.ColumnHeaders.Add , , "Number", 1500, , 14
    bpp_list.ColumnHeaders.Add , , "Message", 5000, , 18
    bpp_list.ColumnHeaders.Add , , "Date/Time"
    bpp_list.ColumnHeaders.Add , , "ID"
   ' bpp_list.ColumnHeaders.Add , , "Name"

       Set rs = New ADODB.Recordset
    strQuery = "select * from FAILEDMESSAGES ORDER BY ID DESC"
    rs.Open strQuery, con, 3, 2, 1
   i = 0
   While Not rs.EOF
                            
                
               
                
       Set A = bpp_list.ListItems.Add(, , rs.Fields("NAME"))
               A.SubItems(2) = rs.Fields("MESSAGE")
               A.SubItems(3) = rs.Fields("TIME_STAMP")
               A.SubItems(1) = rs.Fields("MOBILENO")
               A.SubItems(4) = rs.Fields("ID")
              
               rs.MoveNext
        
    
    Wend

End Sub


Private Sub Timer1_Timer()
'Debug.Print Now
If MSComm1.PortOpen = True Then
    tempBuffer = ""
    MSComm1.Output = "AT+CSQ" + vbCr
bOK = False
        While Not bOK Or bError
        DoEvents
        Wait
        Wend

      'Debug.Print Trim(tempBuffer)
      SignalValue = (Val(Mid(tempBuffer, InStr(tempBuffer, ":") + 1, 3)) * 100) / (31)
          Debug.Print tempBuffer
          DrawSignalLines
    End If


End Sub

Private Sub DrawSignalLines()
Dim temp As String
Dim NoOfLines As Integer
    temp = frmMain.StatusBar.Panels.item(1).Text
    If SignalValue = 319 Then
    MsgBox "No network Coverage", vbCritical
    frmMain.StatusBar.Panels.item(5).Text = "No Network Coverage"
    Exit Sub
    End If
    NoOfLines = SignalValue / 10
    frmMain.StatusBar.Panels.item(5).Text = String(NoOfLines, ">")
    
End Sub


Private Sub Timer2_Timer()
    frmMain.StatusBar.Panels.item(3).Text = Now
    frmSendSingleMessage.txtTimer.Text = Now
    frmSend.txtTimer.Text = Now
   ' Debug.Print Now
    CheckforRecords
End Sub


Private Sub CheckforRecords()
'Debug.Print Now
Dim i As Integer
Dim strQuery As String
Dim rs As ADODB.Recordset
Dim RecordCount

RecordCount = 0
temptime = LCase(Format(Now(), "yyyy-mm-dd hh:mm"))
    'Debug.Print Now
        Set rs = New ADODB.Recordset
    
        strQuery = "Select * from SCHEDULEMESSAGES where  format(SCHEDULETIME,'" & "yyyy-mm-dd hh:mm" & "') = '" & temptime & "'"
   
           
        rs.Open strQuery, con, 3, 2, 1
        i = 0
    While Not rs.EOF
            
               RecordCount = RecordCount + 1
                
        rs.MoveNext
    Wend
                
               
   On Error GoTo finish
    
   If RecordCount <> 0 Then
   Timer2.Enabled = False
   frmSendSchedule.Show
    Exit Sub
   End If
finish:
Timer2.Enabled = True
End Sub

Public Sub GetServiceNumber()
'Dim servicenumber
'Timer1.Enabled = False

'If frmServiceCenter.Changenumber = 1 Then
'  MSComm1.Output = "AT+CSCA=" + frmServiceCenter.txtmessage + vbCr
'  While Not bOK Or bError
'        DoEvents
'        Wait
'        Wend
'
'   If bOK Then
'   MsgBox "Service number changed", vbInformation
'   frmServiceCenter.Changenumber = 0
'   frmServiceCenter.txtmessage = ""
'   Else
'   Exit Sub
'   End If
'   Else
'
'
'
'
'tempBuffer = ""
'MSComm1.Output = "AT+CSCA?" + vbCr
'
'While Not bOK Or bError
'        DoEvents
'        Wait
'        Wend
'
'Debug.Print tempBuffer
'msgBreak = Split(tempBuffer, vbCrLf, , vbTextCompare)
'                msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
'                'MsgBox msgBreak(1)
'               servicenumber = Mid$(Right$(msgBreak(1), 18), 1, 13)
'               frmServiceCenter.txtmessage.Text = servicenumber
'               frmServiceCenter.Show
'
'
'
'
'
'   End If
'
'
End Sub


Private Sub CheckAutoMessages(ByVal Message As String, ByVal Number As String)
Dim mess$
Dim strQuery
Message = LCase(Trim(Message))

            Set rs = New ADODB.Recordset
                strQuery = "select * from AUTOMESSAGE"
            
            rs.Open strQuery, con, 3, 2, 1
     While Not rs.EOF
        If InStr(1, UCase(Message), UCase(rs.Fields("KEYWORD")), vbTextCompare) <> 0 Then
        
       SendSms Number, rs.Fields("AUTOMESSAGE")
      End If
      rs.MoveNext
 Wend
End Sub
