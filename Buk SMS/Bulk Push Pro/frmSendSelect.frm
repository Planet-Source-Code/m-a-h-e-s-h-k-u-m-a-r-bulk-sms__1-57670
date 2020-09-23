VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSendSelect 
   Caption         =   "Select Contacts"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewContact 
      Caption         =   "&New Contact"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select >>"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox cmbGroups 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   5355
   End
   Begin MSComctlLib.ListView lstSelectContacts 
      Height          =   3135
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstSelectedContacts 
      Height          =   2055
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Mobile Number"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSendSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbGroups_Click()
    'MsgBox cmbGroups.Text
    LoadListViewContacts cmbGroups.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNewContact_Click()
    frmAddEditContact.Show
    frmAddEditContact.cmbGroups.Text = cmbGroups.Text
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If lstSelectContacts.ListItems.count Then
        Set A = lstSelectedContacts.ListItems.Add(, , lstSelectContacts.SelectedItem.SubItems(1))
        Set A = frmSend.lstAddressBook.ListItems.Add(, , lstSelectContacts.SelectedItem.SubItems(1))
        lstSelectContacts.ListItems.Remove (lstSelectContacts.SelectedItem.Index)
    End If
End Sub

Private Sub Form_Load()
'Dim rs As ADODB.Recordset
'    Dim strQuery As String
'
'    Set rs = New ADODB.Recordset
'    strQuery = "select GROUPNAME from GROUPS"
'
'    rs.Open strQuery, con, 3, 2, 1
'
'    While Not rs.EOF
'        cmbGroups.AddItem rs.Fields("GROUPNAME")
'        rs.MoveNext
'    Wend
'    rs.Close
'    Set rs = Nothing
'
'    cmbGroups.ListIndex = 0
'
'    If SelectedNode <> "" Then
'        While True
'            If cmbGroups = SelectedNode Then
'                GoTo Finish
'            Else
'                cmbGroups.ListIndex = cmbGroups.ListIndex + 1
'            End If
'        Wend
'    End If
'
'Finish:
End Sub

Private Sub LoadListViewContacts(ByVal cmbSelected As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
        
    lstSelectContacts.ListItems.Clear
    lstSelectContacts.ColumnHeaders.Clear
    lstSelectContacts.ColumnHeaders.Add , , "Name"
    lstSelectContacts.ColumnHeaders.Add , , "Mobile No"
    lstSelectContacts.ColumnHeaders.Add , , "Designation"
   
    Set rs = New ADODB.Recordset
    strQuery = "select * from CONTACTS where GROUPNAME = '" & cmbSelected & "'"
    
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
       Set A = lstSelectContacts.ListItems.Add(, , rs.Fields("CONTACTNAME"))
            A.SubItems(1) = rs.Fields("MOBILE")
            A.SubItems(2) = rs.Fields("DESIGNATION")
        rs.MoveNext
    Wend
End Sub

'Private Sub lstSelectContacts_DblClick()
'    Set A = lstSelectedContacts.ListItems.Add(, , lstSelectContacts.SelectedItem.SubItems(1))
'    Set A = frmSend.lstAddressBook.ListItems.Add(, , lstSelectContacts.SelectedItem.SubItems(1))
'    lstSelectContacts.ListItems.Remove (lstSelectContacts.SelectedItem.Index)
'End Sub
