VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Bulk Push Pro"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6135
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2010
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   3545
      ButtonWidth     =   2752
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglTlbr"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "New contact or group"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewGroup"
                  Text            =   "New &Group"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewContact"
                  Text            =   "New &Contact"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send Message"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Phone Book Reports"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "s"
            Key             =   "3s"
            ImageIndex      =   13
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Message Reports"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Daily Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Reports Between Dates"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Dynamic Reports"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            ImageIndex      =   18
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Read Messages"
         Height          =   375
         Left            =   9360
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "X"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11160
         TabIndex        =   0
         Top             =   -360
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imglTlbr 
      Left            =   2880
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":239E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3342
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3796
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4742
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":579A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":653E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAllGroups 
      Caption         =   "All Groups"
      Visible         =   0   'False
      Begin VB.Menu mnuAddGroup1 
         Caption         =   "&Add Group"
         Index           =   1
      End
      Begin VB.Menu mnuMessageToAll 
         Caption         =   "&Message to All"
      End
   End
   Begin VB.Menu mnuGroupMenu 
      Caption         =   "Group Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSendMessageToGroup 
         Caption         =   "Send message to group"
      End
      Begin VB.Menu mnuSepGM1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddGroup2 
         Caption         =   "&Add Group"
      End
      Begin VB.Menu mnuGroupEdit 
         Caption         =   "&Edit Group"
      End
      Begin VB.Menu mnuRemoveGroup 
         Caption         =   "&Remove Group"
      End
      Begin VB.Menu mnuSepGM2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddContact1 
         Caption         =   "Add &contact in this group"
      End
      Begin VB.Menu mnuSepGM3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindGroup 
         Caption         =   "&Find in group"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuCommonMenu 
      Caption         =   "CommonMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEmptyFolder 
         Caption         =   "&Empty this folder"
      End
   End
   Begin VB.Menu mnuContactMenu 
      Caption         =   "Contact Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSendMessageToContact 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuContactSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddContact2 
         Caption         =   "&Add contact"
      End
      Begin VB.Menu mnuEditContact 
         Caption         =   "&Edit contact"
      End
      Begin VB.Menu nmuContactSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveContact 
         Caption         =   "&Remove Contact"
      End
   End
   Begin VB.Menu mnuInboxMessage 
      Caption         =   "Inbox Message"
      Visible         =   0   'False
      Begin VB.Menu mnuInboxReplyToSender 
         Caption         =   "Reply to &Sender"
      End
      Begin VB.Menu mnuInboxForwardMessage 
         Caption         =   "&Forward Message"
         Begin VB.Menu Sep1 
            Caption         =   "To all"
         End
         Begin VB.Menu Sep2 
            Caption         =   "To a group"
         End
         Begin VB.Menu send_one_contact 
            Caption         =   "One Contact"
         End
      End
      Begin VB.Menu mnuInboxAddPhonebook 
         Caption         =   "&Add to Phonebook"
      End
      Begin VB.Menu mnuInboxDeleteMessage 
         Caption         =   "&Delete Message"
      End
   End
   Begin VB.Menu mnuOutMessage 
      Caption         =   "Outbox Message"
      Visible         =   0   'False
      Begin VB.Menu mnuOutMessageToRecipent 
         Caption         =   "Message to Recipient"
      End
      Begin VB.Menu mnuDeleteOutboxMessage 
         Caption         =   "&Delete Message"
      End
      Begin VB.Menu removeAllOutboxofText 
         Caption         =   "Delete All Messages of this Text"
      End
      Begin VB.Menu mnuOutForwardMessage 
         Caption         =   "&Forward Message"
         Begin VB.Menu Sep5 
            Caption         =   "To group"
         End
         Begin VB.Menu to_all 
            Caption         =   "To all"
         End
         Begin VB.Menu Sep6 
            Caption         =   "One Contact"
         End
      End
      Begin VB.Menu mnuOutAddPhoneBook 
         Caption         =   "Add to &Phonebook"
      End
   End
   Begin VB.Menu mnuSentmessages 
      Caption         =   "Sent Messages"
      Visible         =   0   'False
      Begin VB.Menu mnuSentMessageToRecipient 
         Caption         =   "Message to Recipient"
      End
      Begin VB.Menu mnuForwardMessage 
         Caption         =   "&Forward Message"
         Begin VB.Menu Sep8 
            Caption         =   "To all"
         End
         Begin VB.Menu to_group 
            Caption         =   "To a group"
         End
         Begin VB.Menu one_contact 
            Caption         =   "One Contact"
         End
      End
      Begin VB.Menu mnuSentDeleteMessage 
         Caption         =   "&Delete Message"
      End
      Begin VB.Menu deleteAllContaining 
         Caption         =   "Delete All messages of this text"
      End
      Begin VB.Menu Sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSentAddPhoneBook 
         Caption         =   "Add to Phonebook"
      End
   End
   Begin VB.Menu mnuAutoReply 
      Caption         =   "Auto Reply"
      Visible         =   0   'False
      Begin VB.Menu mnuAutoAddKeyword 
         Caption         =   "&Add Keyword"
      End
      Begin VB.Menu mnuAutoRemoveKeyword 
         Caption         =   "&Remove Keyword"
      End
      Begin VB.Menu Sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoRemoveAll 
         Caption         =   "Remove a&ll keywords"
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnulstContactRightClick 
      Caption         =   "ListContactrighClick"
      Visible         =   0   'False
      Begin VB.Menu mnuEditLstContact 
         Caption         =   "Edit Contact"
      End
      Begin VB.Menu mnuDeletelistContact 
         Caption         =   "&Delete Contact"
      End
      Begin VB.Menu mnuSendlistcontact 
         Caption         =   "&SendMessage"
      End
   End
   Begin VB.Menu mnuSchedulList 
      Caption         =   "SendSchedule"
      Visible         =   0   'False
      Begin VB.Menu mnuScheduleSendnow 
         Caption         =   "Send Now"
      End
      Begin VB.Menu mnuDeleteScheduleList 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnudeleallofthisschedulelist 
         Caption         =   "Delete All of this text"
      End
   End
   Begin VB.Menu mnuFailedMessages 
      Caption         =   "FailedMessages"
      Visible         =   0   'False
      Begin VB.Menu mnuSendAllFailed 
         Caption         =   "Send now"
      End
      Begin VB.Menu mnuDeleteFailed 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnudeleteAllofTextFailed 
         Caption         =   "Delete &All of this text"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()

'MsgBox frmChild.bpp_tree.SelectedItem.Tag
If frmChild.bpp_tree.SelectedItem.Tag = "" Then
Exit Sub
End If
If Split(frmChild.bpp_tree.SelectedItem.Tag, "|")(0) = "GRP" Then
    RemoveGroup frmChild.bpp_tree.SelectedItem
    Else
        If Split(frmChild.bpp_tree.SelectedItem.Tag, "|")(0) = "CONT" Then
            RemoveContact frmChild.bpp_tree.SelectedItem.Parent, frmChild.bpp_tree.SelectedItem
        End If
        
End If

End Sub

'Private Sub cmdSend_Click()
'    AT_flag1 = True
'    SMS_check = True
'    If AT_flag = True Then
'        AT_flag = False
'
'        frmSend.Show
'    ElseIf Formload_check = False Then
'        MsgBox "Please Wait Checking Modem Connections........"
'        AT_flag1 = False
'        SMS_check = False
'    End If
''frmSend.Show
'End Sub

Private Sub Command1_Click()
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    frmReadMessages.Show
End Sub

Private Sub deleteAllContaining_Click()
On Error Resume Next
removeAllOfText frmChild.bpp_list.SelectedItem.SubItems(2), frmChild.bpp_tree.SelectedItem.Text
End Sub

Private Sub MDIForm_Load()
   Dim i As Integer
   On Error GoTo handler:
     ConnectDB
    frmChild.Show
    StatusBarSet
   ' frmSendSchedule.Timer1.Enabled = True


handler:
'Debug.Print Err.Number
 If Err.Number = -2147467259 Then
 MsgBox "Data base not found"
 Unload Me
 End If

End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo handler
    DisconnectDB
   
    'If frmChild.MSComm1.PortOpen = True Then frmChild.MSComm1.PortOpen = False
handler:
'frmScheduler.Timer1.Enabled = False
Unload frmChild
Unload frmAddAutoMessage
Unload frmAddEditContact
Unload frmAddEditGroup
Unload frmBppSplash
Unload frmCalendar
Unload frmContactEdit
Unload frmDailyReports
Unload frmFind
Unload frmGroupEdit
Unload frmMain
'Unload frmReadMessages
Unload frmReport
'Unload frmScheduler
Unload frmSend
Unload frmSendSchedule
Unload frmSendSingleMessage
'Unload frmServiceCenter
Unload frmSettings
Unload frmSingleCalender
Unload frmPhonebookReport
Unload frmfromtoreport




frmChild.Timer1 = False
frmChild.Timer2 = False
Exit Sub

End Sub

Private Sub mnuAddContact1_Click()
  frmAddEditContact.SelectedNode = frmChild.bpp_tree.SelectedItem
  frmChild.Show
  'frmChild.Enabled = False
  frmAddEditContact.Show
  'frmAddEditContact.Appearance
End Sub

Private Sub mnuAddContact2_Click()
  frmAddEditContact.SelectedNode = frmChild.bpp_tree.SelectedItem.Parent
  frmAddEditContact.Show
  frmChild.bpp_list.Refresh
'    frmChild.bpp_list.SetFocus
End Sub

Private Sub mnuAddGroup1_Click(Index As Integer)
     frmAddEditGroup.Show
     
End Sub

Private Sub mnuAddGroup2_Click()
    
    frmAddEditGroup.Show
End Sub

Private Sub mnuAutoAddKeyword_Click()
frmAddAutoMessage.Show
End Sub

Private Sub mnuAutoRemoveAll_Click()
    RemoveAllAutoMessage
End Sub

Private Sub mnuAutoRemoveKeyword_Click()
    RemoveAutoMessage (frmChild.bpp_list.SelectedItem)
    frmChild.bpp_list.Refresh
    frmChild.bpp_list.SetFocus
End Sub

Private Sub mnudeleallofthisschedulelist_Click()
removeAllOfText frmChild.bpp_list.SelectedItem.SubItems(2), frmChild.bpp_tree.SelectedItem.Text
End Sub

Private Sub mnudeleteAllofTextFailed_Click()
removeAllOfText frmChild.bpp_list.SelectedItem.SubItems(2), frmChild.bpp_tree.SelectedItem.Text
End Sub

Private Sub mnuDeleteFailed_Click()
frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "' messages from Failed Messages ? ", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
       RemoveFailedItem frmChild.bpp_list.ListItems.item(i).ListSubItems.item(4).Text
        End If
     
     Next i
     End If


  frmChild.LoadFailedMessages
    frmChild.RefreshTree
    frmChild.bpp_list.Refresh
    frmChild.bpp_tree.SetFocus
End Sub

Private Sub mnuDeleteOutboxMessage_Click()
    frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "' messages from Outbox ? ", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
          RemoveOutboxItem frmChild.bpp_list.ListItems.item(i).SubItems(4)
        End If
     
     Next i
     End If
     frmChild.LoadOutbox
   frmChild.bpp_list.Refresh
   frmChild.bpp_list.SetFocus
    
End Sub

Private Sub mnuDeletelistContact_Click()

 frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "'contacts ? ", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
         RemoveContact frmChild.bpp_tree.SelectedItem, frmChild.bpp_list.ListItems.item(i).Text
        End If
     
     Next i
     End If


  frmChild.LoadListViewContacts frmChild.Label4.Caption
    frmChild.RefreshTree
    frmChild.bpp_list.Refresh
    'frmChild.bpp_tree.SetFocus
End Sub

Private Sub amnuDeleteOutboxMessage_Click()
    frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "' messages from Inbox ? ", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
          RemoveOutboxItem frmChild.bpp_list.ListItems.item(i).SubItems(4)
        End If
     
     Next i
     End If
     frmChild.LoadOutbox
   frmChild.bpp_list.Refresh
   frmChild.bpp_list.SetFocus
    
  
End Sub




Private Sub mnuEditContact_Click()
 frmContactEdit.FillForm frmChild.bpp_tree.SelectedItem.Parent, frmChild.bpp_tree.SelectedItem
    frmContactEdit.Show
    
End Sub

Private Sub mnuEditLstContact_Click()
On Error Resume Next
 frmContactEdit.FillForm frmChild.bpp_tree.SelectedItem, frmChild.bpp_list.SelectedItem
frmContactEdit.Show
End Sub

Private Sub mnuEmptyFolder_Click()
    'MsgBox frmChild.bpp_tree.SelectedItem.Text
    EmptyFolder (frmChild.bpp_tree.SelectedItem.Text)
End Sub

'Private Sub mnuForwardMessage_Click()
'    frmSend.Show
'    frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(1)
'End Sub

Private Sub mnuGroupEdit_Click()
    frmGroupEdit.strGroupName = frmChild.bpp_tree.SelectedItem.Text
    frmGroupEdit.Show
End Sub

Private Sub mnuInboxAddPhonebook_Click()
    frmAddEditContact.Show
    frmAddEditContact.txtMobile = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
End Sub

Private Sub mnuInboxDeleteMessage_Click()
frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "' messages from Inbox ? ", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
            RemoveInboxItem (frmChild.bpp_list.ListItems.item(i).SubItems(4))
        End If
     
     Next i
     End If
     frmChild.LoadInbox
   frmChild.bpp_list.Refresh
   'frmChild.bpp_list.SetFocus
End Sub

Private Sub AddContactToGroup(ByVal GroupName As String, ByVal ContactName As String)
    MsgBox GroupName & "  " & ContactName
End Sub

Private Sub mnuInboxReplyToSender_Click()
 On Error Resume Next
   frmSendSingleMessage.Text1.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
    frmSendSingleMessage.Show
    
    
End Sub

Private Sub mnuMessageToAll_Click()
LoadSendList
End Sub

Private Sub mnuOutAddPhoneBook_Click()
    frmAddEditContact.Show
    frmAddEditContact.txtMobile = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
End Sub

'Private Sub mnuOutForwardMessage_Click()
'    frmSend.Show
'    frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(1)
'End Sub

Private Sub mnuOutMessageToRecipent_Click()
On Error Resume Next
frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(2).Text
   frmSendSingleMessage.Text1.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
   frmSendSingleMessage.Show
End Sub



Private Sub mnuRemoveContact_Click()

   RemoveContact frmChild.bpp_tree.SelectedItem.Parent, frmChild.bpp_tree.SelectedItem
    frmChild.RefreshTree
    frmChild.bpp_list.Refresh
    frmChild.bpp_list.SetFocus
End Sub

Private Sub mnuRemoveGroup_Click()
    RemoveGroup frmChild.bpp_tree.SelectedItem
End Sub

Public Sub RemoveGroup(ByVal GroupName As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    Dim ChildCount As Integer
    Dim Reply As Integer
    
    Set rs = New ADODB.Recordset
    strQuery = "select count(*) as CHILD_COUNT from CONTACTS where GROUPNAME = '" & GroupName & "'"
    
    rs.Open strQuery, con, 3, 2, 1
    ChildCount = rs.Fields("CHILD_COUNT")
    rs.Close
    
    strQuery = "delete from GROUPS where GROUPNAME = '" & GroupName & "'"
    If ChildCount Then
        If MsgBox("Group '" & GroupName & "' contains " & ChildCount & " contacts. Do you want to delete ?", vbYesNo) = vbYes Then
            con.Execute strQuery
            frmChild.RefreshTree
        Else
        frmChild.RefreshTree
        End If
    Else
        If MsgBox("Are you sure to delete group '" & GroupName & "' ?", vbYesNo) = vbYes Then
            con.Execute strQuery
            frmChild.RefreshTree
        End If
    End If
    
    Set rs = Nothing

End Sub
Public Sub RemoveContact(ByVal GroupName As String, ByVal ContactName As String)
    Dim strQuery As String
    If MsgBox("Do you want to delete " & ContactName & " ? ", vbYesNo) = vbYes Then
    strQuery = "delete from CONTACTS where GROUPNAME = '" & GroupName & "' and CONTACTNAME = '" & ContactName & "'"
    
    
        con.Execute strQuery
           frmChild.RefreshTree
   End If
        
'bpp_tree.Enabled = True
End Sub
Private Sub RemoveInboxItem(ByVal inboxid As String)
On Error GoTo handler
Dim strQuery As String
    
    strQuery = "delete from INBOX where INBOXID = " & inboxid & " "
    
  '   Debug.Print strQuery
        
        con.Execute strQuery
    
    
handler:


End Sub




Private Sub mnuScheduleSendnow_Click()
frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(2).Text
frmSendSingleMessage.Text1.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
frmSendSingleMessage.Show

End Sub

Private Sub mnuSendAllFailed_Click()
LoadFailedlist frmChild.bpp_list.SelectedItem.ListSubItems.item(4).Text
frmSendSingleMessage.Show
End Sub

Private Sub mnuSendlistcontact_Click()
On Error Resume Next
frmSendSingleMessage.Text1.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
frmSendSingleMessage.Show
End Sub

Private Sub mnuSendMessageToContact_Click()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
        Set rs = New ADODB.Recordset
        strQuery = "select MOBILE from CONTACTS where CONTACTNAME = '" & frmChild.bpp_tree.SelectedItem & "'"
        rs.Open strQuery, con, 3, 2, 1
        frmSendSingleMessage.Text1.Text = rs.Fields("MOBILE")
        rs.Close
        
    frmSendSingleMessage.Show
    
End Sub

Public Sub mnuSendMessageToGroup_Click()
LoadGroupSendList frmChild.bpp_tree.SelectedItem

frmSend.Show
End Sub

Private Sub mnuSentAddPhoneBook_Click()
    
    frmAddEditContact.txtMobile = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
    frmAddEditContact.Show
    
End Sub

Private Sub mnuSentDeleteMessage_Click()
   frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "' messages from Inbox ?", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
        RemoveSentMessage frmChild.bpp_list.ListItems.item(i).SubItems(4)
        End If
     
     Next i
     End If
    frmChild.LoadSentMessages
    frmChild.bpp_list.Refresh
    frmChild.bpp_list.SetFocus
End Sub

Private Sub mnuSentMessageToRecipient_Click()
   On Error Resume Next
   frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(2).Text
   frmSendSingleMessage.Text1.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
   frmSendSingleMessage.Show
End Sub

Private Sub one_contact_Click()
    frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.SubItems(2)
  
    frmSendSingleMessage.Show
End Sub

Private Sub removeAllOutboxofText_Click()
 removeAllOfText frmChild.bpp_list.SelectedItem.SubItems(2), frmChild.bpp_tree.SelectedItem.Text
End Sub

Private Sub send_one_contact_Click()
    frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.SubItems(2)
  
    frmSendSingleMessage.Show
End Sub

Private Sub Sep1_Click()
    frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(2)
    LoadSendList
    frmSend.Show
End Sub

Private Sub Sep2_Click()
frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(2)
LoadGroupSendList frmSend.cmbGroups

frmSend.Show
End Sub

Private Sub Sep5_Click()
frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(2)
LoadGroupSendList frmSend.cmbGroups

frmSend.Show
End Sub

Private Sub Sep6_Click()
frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.SubItems(2)

    frmSendSingleMessage.Show
End Sub

Private Sub Sep8_Click()
    frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(2)
    LoadSendList

    frmSend.Show
    

End Sub

Private Sub to_all_Click()
    frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(2)
    LoadSendList

    frmSend.Show
End Sub

Private Sub to_group_Click()
frmSend.txtMessage = frmChild.bpp_list.SelectedItem.SubItems(2)
LoadGroupSendList frmSend.cmbGroups

frmSend.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim D() As String
    Dim A As ListItem
    'MsgBox bpp_tree.SelectedItem.Tag
    
    
    'MsgBox bpp_tree.SelectedItem.Tag
    On Error GoTo handler2
    Node_key = frmChild.bpp_tree.SelectedItem.key
    
    Select Case Button.Index
    Case 2
        frmAddEditGroup.Show
    Case 3
            frmSendSingleMessage.Show

    Case 4
        cmdDelete_Click
'    Case 5
'        'frmChild.Timer1.Enabled = False
          
    Case 5
        '1.Enabled = False
        frmPhonebookReport.Show
    Case 6
    
    Case 7
            frmDailyReports.Show
    Case 8
           frmFind.Show
    End Select
handler2:
'  frmChild.bpp_list.SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    
    Debug.Print ButtonMenu.Parent
    
    If ButtonMenu.Parent = "Message Reports" Then
     Select Case ButtonMenu.Index
    Case 1
        frmDailyReports.Show
    Case 2
        frmfromtoreport.Show
    Case 3
         frmReport.Show
    End Select
    Exit Sub
   End If
    
    Select Case ButtonMenu.Index
    Case 1
        frmAddEditGroup.Show
    Case 2
        frmAddEditContact.SelectedNode = ""
        frmAddEditContact.Show
    End Select
    
End Sub

Private Sub RemoveOutboxItem(ByVal id As String)

Dim strQuery As String
    
    strQuery = "delete from OUTBOX where ID = " & id & " "
 '   InputBox "", "", strQuery
      
    Debug.Print strQuery
        con.Execute strQuery
        
    



End Sub

Private Sub RemoveSentMessage(ByVal id As String)

Dim strQuery As String
    
    strQuery = "delete from SENTMESSAGES where ID =" & id & " "
 '   InputBox "", "", strQuery
    
    
        con.Execute strQuery
        
    

End Sub

Private Sub RemoveAutoMessage(ByVal Keyword As String)

Dim strQuery As String

    strQuery = "delete from AUTOMESSAGE where KEYWORD = '" & Keyword & "' "
 '   InputBox "", "", strQuery

    If MsgBox("Are you sure you want to delete selected keyword ? ", vbYesNo) = vbYes Then
        con.Execute strQuery
        frmChild.LoadAutoMessages
    End If
End Sub

Private Sub RemoveAllAutoMessage()

Dim strQuery As String

    strQuery = "delete * from AUTOMESSAGE "
 '   InputBox "", "", strQuery

    If MsgBox("Are you sure you want to delete all Keywords ? ", vbYesNo) = vbYes Then
        con.Execute strQuery
        frmChild.LoadAutoMessages
    End If
End Sub

Private Sub EmptyFolder(ByVal FolderName As String)
' Trim FolderName, ""
    If LCase(FolderName) = LCase("sent messages") Then FolderName = "SentMessages"
    If LCase(FolderName) = LCase("Schedule Messages") Then FolderName = "ScheduleMessages"
    If LCase(FolderName) = LCase("Failed Messages") Then FolderName = "FailedMessages"
    Dim strQuery As String

    strQuery = "delete * from " & FolderName & ""
 '   InputBox "", "", strQuery

    If MsgBox("Are you sure you want to delete all Messages of '" & FolderName & "' ", vbYesNo) = vbYes Then
        con.Execute strQuery
    frmChild.bpp_list.Refresh
    End If
End Sub
Private Sub StatusBarSet()
'StatusBar.Panels.item(1).Text = "Signal Value"
StatusBar.Panels.item(4).Text = "Signal Strength"
'MsgBox StatusBar.Panels.Item(1).Text
End Sub

Public Sub LoadSendList()
Dim i As Integer
Dim rs As ADODB.Recordset
Dim NumberOfRecords As Integer
    Dim strQuery As String
    
    frmSend.lstAddressBook.ListItems.Clear
'    frmSendSelect.lstSelectedContacts.ListItems.Clear
    frmSend.lstAddressBook.ColumnHeaders.Add , , "Name"
   frmSend.lstAddressBook.ColumnHeaders.Add , , "Mobile Number"
    
    Set rs = New ADODB.Recordset
    strQuery = "select count(*) as CONTACT_COUNT from CONTACTS"
    rs.Open strQuery, con, 3, 2, 1
    NumberOfRecords = rs.Fields("CONTACT_COUNT")
    rs.Close
    
    strQuery = "select * from CONTACTS"
    
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
       Set A = frmSend.lstAddressBook.ListItems.Add(, , rs.Fields("CONTACTNAME"))
            A.SubItems(1) = rs.Fields("MOBILE")
       rs.MoveNext
    Wend
    For i = 1 To frmSend.lstAddressBook.ListItems.Count
    frmSend.lstAddressBook.ListItems.item(i).Checked = True
    Next i
    
    frmSend.frameSend = "Message to All"
    frmSend.LabelNoOfRecords = NumberOfRecords
    frmSend.rdSelectFromAddressBook = True
    frmSend.rdGroups.Visible = False
    frmSend.rdSend.Visible = False
    frmSend.Label1.Visible = False
    frmSend.cmbGroups.Enabled = False
    frmSend.cmbGroups.Visible = False
    frmSend.txtSend.Visible = False
    frmSend.cmdAddressBook.Visible = False
    frmSend.rdSelectFromAddressBook.Top = frmSend.rdSend.Top
    'frmSend.lstAddressBook.Top = 600
    'frmSend.txtMessage.Top = frmSend.lstAddressBook.Top + 10
    frmSend.rdSelectFromAddressBook = True
    
    frmSend.Show
End Sub
Public Sub LoadGroupSendList(ByVal item As String)
Dim i As Integer
Dim rs As ADODB.Recordset
Dim NumberOfRecords As Integer
    Dim strQuery As String
    
    frmSend.lstAddressBook.ListItems.Clear
    
    frmSend.lstAddressBook.ColumnHeaders.Add , , "Name"
    frmSend.lstAddressBook.ColumnHeaders.Add , , "Mobile Number"
    
    Set rs = New ADODB.Recordset
    strQuery = "select count(*) as CONTACT_COUNT from CONTACTS where GROUPNAME = '" & item & "'"
    rs.Open strQuery, con, 3, 2, 1
    NumberOfRecords = rs.Fields("CONTACT_COUNT")
    rs.Close
    
    strQuery = "select * from CONTACTS where GROUPNAME = '" & item & "'"
    
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
       Set A = frmSend.lstAddressBook.ListItems.Add(, , rs.Fields("CONTACTNAME"))
           A.SubItems(1) = rs.Fields("MOBILE")
          rs.MoveNext
    Wend
    
    For i = 1 To frmSend.lstAddressBook.ListItems.Count
    frmSend.lstAddressBook.ListItems.item(i).Checked = True
    Next i
    
    
    
'    frmChild.Sendfrm
    frmSend.rdSelectFromAddressBook = False
    frmSend.rdGroups.Visible = True
    frmSend.rdSend.Visible = False
    frmSend.lstAddressBook.Visible = True
    frmSend.rdSelectFromAddressBook.Visible = False
    frmSend.Label1.Visible = True
    frmSend.cmdAddressBook.Visible = False
    
    frmSend.cmbGroups.Visible = True
    frmSend.txtSend.Visible = False
    'frmSend.Label3(0).Left = frmSend.cmbGroups.Left
   ' frmSend.txtMessage.Top = 1200
    'frmSend.txtMessage.Left = frmSend.cmbGroups.Left
    frmSend.rdSelectFromAddressBook.Top = frmSend.rdSend.Top
    frmSend.rdGroups = True
    frmSend.Label1.Caption = "Group"
    frmSend.frameSend.Caption = "Select Group Name"
    frmSend.lstAddressBook.Enabled = True
    
    'frmSend.txtMessage.Left = frmSend.cmbGroups.Left
    'frmSend.txtMessage.Top = frmSend.lstAddressBook.Top + 10
End Sub
Private Sub CheckNode1(ByVal check_node As Node)
Dim D() As String

D = Split(check_node.Tag, "|")
    If D(0) = "GRP" Then
        LoadGroupSendList frmChild.bpp_tree.SelectedItem
       
        frmSend.Show
    End If
    
    If D(0) = "CONT" Then
    mnuSendMessageToContact_Click
    End If
        
End Sub

Private Sub removeAllOfText(ByVal Message As String, ByVal table As String)

Dim strQuery As String
   If table = "Schedule Messages" Then table = "ScheduleMessages"
   If table = "Sent Messages" Then table = "SentMessages"
   If table = "Failed Messages" Then table = "FailedMessages"
    strQuery = "delete from " & table & " where MESSAGE  ='" & Message & "' "
 '   InputBox "", "", strQuery
    
    If MsgBox("Are you sure you want to delete Message from " & table, vbYesNo) = vbYes Then
        con.Execute strQuery
        
    End If
Select Case table
       Case "Outbox"
        frmChild.LoadOutbox
       Case "Inbox"
        frmChild.LoadInbox
       Case "SentMessages"
       frmChild.LoadSentMessages
       Case "ScheduleMessages"
       frmChild.LoadSchedule
       Case "FailedMessages"
       frmChild.LoadFailedMessages
   End Select
'frmChild.LoadSentMessages
End Sub


Private Sub mnuDeleteScheduleList_Click()
frmChild.bpp_list.MultiSelect = True
Dim j
j = 0
On Error Resume Next

    For i = 1 To frmChild.bpp_list.ListItems.Count
        If frmChild.bpp_list.ListItems.item(i).Selected Then
        j = j + 1
        End If
     Next i
     
     If MsgBox("Are you sure you want to delete  '" & j & "' messages from Schedule Messages ? ", vbYesNo) = vbYes Then
         
     For i = 1 To frmChild.bpp_list.ListItems.Count
     
     If frmChild.bpp_list.ListItems.item(i).Selected Then
        
            RemoveScheuleItem (frmChild.bpp_list.ListItems.item(i).SubItems(4))
        End If
     
     Next i
     End If
     frmChild.LoadSchedule
   frmChild.bpp_list.Refresh
   frmChild.bpp_list.SetFocus

End Sub

Private Sub RemoveScheuleItem(ByVal inboxid As String)
On Error GoTo handler
Dim strQuery As String
    
    strQuery = "delete from SCHEDULEMESSAGES where ID = " & inboxid & " "
    
  '   Debug.Print strQuery
        
        con.Execute strQuery
    
    
handler:


End Sub
Private Sub RemoveFailedItem(ByVal id As String)
On Error GoTo handler
Dim strQuery As String
    
    strQuery = "delete from FAILEDMESSAGES where ID = " & id & " "
    
  '   Debug.Print strQuery
        
        con.Execute strQuery
    
    
handler:


End Sub
Public Sub LoadFailedlist(ByVal id As String)
Dim i As Integer
Dim rs As ADODB.Recordset
Dim NumberOfRecords As Integer
Dim TEMPMESSAGE
    Dim strQuery As String
    
  
   
           frmSendSingleMessage.Text1.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(1).Text
          frmSendSingleMessage.Text2.Text = frmChild.bpp_list.SelectedItem.ListSubItems.item(2).Text
         
    
   
End Sub
