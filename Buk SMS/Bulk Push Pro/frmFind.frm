VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   5295
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter String to Search"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton command3 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtserch 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox cmbserchby 
         Height          =   315
         ItemData        =   "frmFind.frx":030A
         Left            =   1080
         List            =   "frmFind.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Look &in"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Find "
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   6255
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3705
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1665
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   1695
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
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
      Begin VB.Label l3 
         Caption         =   "Designition:"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LabelDesig 
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label labelNum 
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label l2 
         Caption         =   "Mobile No :  "
         Height          =   495
         Left            =   3480
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LabelName 
         Height          =   495
         Left            =   4560
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label ll1 
         Caption         =   "Name"
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As ListItem

Private Sub cmdvi_Click()

End Sub

Private Sub cmbserchby_Change()
    If cmbserchby.ListIndex = 0 Then
        txtserch.Text = "Enter ID Number"
        txtserch.SetFocus
    ElseIf cmbserchby.ListIndex = 1 Then
        txtserch.Text = "Enter Last Name"
    ElseIf cmbserchby.ListIndex = 2 Then
        txtserch.Text = "Enter First Name"
    End If



End Sub

Private Sub Command1_Click()
Unload Me
End Sub



Private Sub Command3_Click()
        
    
    
    
        If cmbserchby.ListIndex = -1 Then
        
        MsgBox "Select search type"
        
        ElseIf cmbserchby.ListIndex = 0 Then
'            If IsNumeric(txtserch) Then
'
'            MsgBox "Invalid Name"
'            Exit Sub
            
'            End If
            
            SearchWithName txtserch
        ElseIf cmbserchby.ListIndex = 1 Then
'            If Not IsNumeric(txtserch) Then
'            MsgBox "Invalid Number"
'            Exit Sub
'            End If
            SearchWithMobile txtserch
       ElseIf cmbserchby.ListIndex = 2 Then
'            If IsNumeric(txtserch) Then
'            MsgBox "Invalid Designition"
'            Exit Sub
'            End If
            SearchWithDesignition txtserch
            
        End If

    
    
    StatusBar1.Visible = True
    frmFind.Height = 6030
End Sub

Private Sub SearchWithName(ByVal searchString As String)
'On Error Resume Next
  Dim strQuery As String
 Dim rs As ADODB.Recordset
 Dim RecordCount
    RecordCount = 0
     lv1.ListItems.Clear
        lv1.ColumnHeaders.Clear
        lv1.ColumnHeaders.Add , , "Name", 1700
        lv1.ColumnHeaders.Add , , "Number", 1500
        lv1.ColumnHeaders.Add , , "Designation", 1500
        lv1.ColumnHeaders.Add , , "Group Name", 1500
        lv1.ColumnHeaders.Add , , "ID"
    
    Set rs = New ADODB.Recordset
    
'     strQuery = "select * from CONTACTS where CONTACTNAME like '" & searchString & "" + "*'" '"
      strQuery = "select * from CONTACTS"
     
     Debug.Print strQuery
      rs.Open strQuery, con, 3, 2, 1
       
    While Not rs.EOF
        If LCase(Mid(rs.Fields("CONTACTNAME"), 1, Len(searchString))) = LCase(searchString) Then
        RecordCount = RecordCount + 1
        Debug.Print rs.Fields("CONTACTNAME")
         Me.Height = 4815
        Set A = lv1.ListItems.Add(, , rs.Fields("CONTACTNAME"))
            A.SubItems(1) = rs.Fields("MOBILE")
'            A.SubItems(2) = rs.Fields("DESIGNATION")
            A.SubItems(3) = rs.Fields("GROUPNAME")
            A.SubItems(4) = rs.Fields("CONTACTID")
        End If
        
        rs.MoveNext
Wend
If RecordCount = 0 Then
   StatusBar1.Panels(1).Text = "No Records Found"
Else
   StatusBar1.Panels(1).Text = " " & RecordCount & "   Records Found"
End If
End Sub
   


Private Sub SearchWithMobile(ByVal searchString As String)
'On Error Resume Next
  Dim strQuery As String
 Dim rs As ADODB.Recordset
     Dim RecordCount
    RecordCount = 0
    
     lv1.ListItems.Clear
        lv1.ColumnHeaders.Clear
        lv1.ColumnHeaders.Add , , "Number", 1500
        lv1.ColumnHeaders.Add , , "Name", 1700
        
        lv1.ColumnHeaders.Add , , "Designation", 1500
        lv1.ColumnHeaders.Add , , "Group Name", 1500
        lv1.ColumnHeaders.Add , , "ID"
    
    Set rs = New ADODB.Recordset
    
'     strQuery = "select * from CONTACTS where CONTACTNAME like '" & searchString & "" + "*'" '"
      strQuery = "select * from CONTACTS"
     
     Debug.Print strQuery
      rs.Open strQuery, con, 3, 2, 1
       
    While Not rs.EOF
        If LCase(Mid(rs.Fields("MOBILE"), 1, Len(searchString))) = LCase(searchString) Then
        RecordCount = RecordCount + 1
        Debug.Print rs.Fields("CONTACTNAME")
         Me.Height = 4815
        Set A = lv1.ListItems.Add(, , rs.Fields("MOBILE"))
            A.SubItems(1) = rs.Fields("CONTACTNAME")
            A.SubItems(2) = rs.Fields("DESIGNATION")
             A.SubItems(3) = rs.Fields("GROUPNAME")
            A.SubItems(4) = rs.Fields("CONTACTID")
        End If
        
        rs.MoveNext
Wend
If RecordCount = 0 Then
   StatusBar1.Panels(1).Text = "No Records Found"
Else
   StatusBar1.Panels(1).Text = " " & RecordCount & "   Records Found"
End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Command4_Click()
DeleteContact
frmChild.RefreshTree
frmChild.bpp_list.Refresh

End Sub

Private Sub Form_Load()
Me.Height = 2670
cmbserchby.ListIndex = 0
StatusBar1.Visible = False
End Sub

Private Sub lv1_ItemClick(ByVal item As MSComctlLib.ListItem)
LabelName.Caption = lv1.SelectedItem.Text
labelNum.Caption = lv1.SelectedItem.ListSubItems.item(1).Text
LabelDesig.Caption = lv1.SelectedItem.ListSubItems.item(2).Text
End Sub

Private Sub SearchWithDesignition(ByVal searchString As String)
'On Error Resume Next
  Dim strQuery As String
 Dim rs As ADODB.Recordset
     Dim RecordCount
    RecordCount = 0
    
     lv1.ListItems.Clear
        lv1.ColumnHeaders.Clear
        
        lv1.ColumnHeaders.Add , , "Name", 1700
        lv1.ColumnHeaders.Add , , "Designation", 1500
        lv1.ColumnHeaders.Add , , "Number", 1500
        lv1.ColumnHeaders.Add , , "Group Name", 1500
        lv1.ColumnHeaders.Add , , "ID"
   
    
    Set rs = New ADODB.Recordset
    
'     strQuery = "select * from CONTACTS where CONTACTNAME like '" & searchString & "" + "*'" '"
      strQuery = "select * from CONTACTS"
     
     Debug.Print strQuery
      rs.Open strQuery, con, 3, 2, 1
       
    While Not rs.EOF
        If LCase(Mid(rs.Fields("DESIGNATION"), 1, Len(searchString))) = LCase(searchString) Then
          RecordCount = RecordCount + 1
        Debug.Print rs.Fields("CONTACTNAME")
         Me.Height = 4815
        Set A = lv1.ListItems.Add(, , rs.Fields("CONTACTNAME"))
             A.SubItems(1) = rs.Fields("DESIGNATION")
            
            A.SubItems(2) = rs.Fields("MOBILE")
             A.SubItems(3) = rs.Fields("GROUPNAME")
            A.SubItems(4) = rs.Fields("CONTACTID")
           
        End If
      
        rs.MoveNext
Wend

If RecordCount = 0 Then
   StatusBar1.Panels(1).Text = "No Records Found"
Else
   StatusBar1.Panels(1).Text = " " & RecordCount & "   Records Found"
End If
End Sub

Private Sub Command2_Click()
If IsNumeric(lv1.SelectedItem.Text) Then
frmContactEdit.FillForm lv1.SelectedItem.ListSubItems.item(3).Text, lv1.SelectedItem.ListSubItems.item(1).Text
frmContactEdit.Show
Else
frmContactEdit.FillForm lv1.SelectedItem.ListSubItems.item(3).Text, lv1.SelectedItem.Text
frmContactEdit.Show
End If
End Sub


Private Sub DeleteContact()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
   'On Error Resume Next
    Dim id
    id = lv1.SelectedItem.ListSubItems.item(4).Text
If MsgBox("Are you sure you want to delete  '" & lv1.SelectedItem.Text & "'   from contacts ? ", vbYesNo) = vbYes Then
     lv1.ListItems.Remove (lv1.SelectedItem.Index)
    strQuery = "delete from CONTACTS where CONTACTID = " & id & ""
    Debug.Print strQuery
    con.Execute strQuery
End If
End Sub
