VERSION 5.00
Begin VB.Form frmAddEditContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Contact"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmAddEditContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contact Information"
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtDesigantion 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtMobile 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   13
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desigantion"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cell:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group"
      Height          =   795
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cmbGroups 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   5355
      End
   End
End
Attribute VB_Name = "frmAddEditContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedNode As String



Private Sub cmdAddEdit_Click()

    On Error GoTo handler


    Dim rs As ADODB.Recordset
    Dim strQuery As String
    If Not IsNumeric(txtMobile) Then
    MsgBox "Invalid number"
    Exit Sub
    End If
    If txtName = "" Then
        MsgBox "Contact name cannot be empty"
        Exit Sub
    ElseIf txtMobile = "" Then
        MsgBox "Mobile number cannot be empty"
        Exit Sub
    Else
        If txtDesigantion = "" Or txtEmail = "" Then
        
        txtDesigantion = "Unknown"
        txtEmail = "Unknown"
        End If
        
        Set rs = New ADODB.Recordset
        strQuery = "select count(CONTACTNAME) as CONTACT_COUNT from CONTACTS where CONTACTNAME = '" & txtName & "' and GROUPNAME = '" & cmbGroups & "'"
        rs.Open strQuery, con, 3, 2, 1
        If rs.Fields("CONTACT_COUNT") Then
            MsgBox "Contact name '" & txtName & "' already exists in this group, Please choose another", vbInformation
        Else
            rs.Close
            strQuery = "select count(MOBILE) as MOB_COUNT from CONTACTS where MOBILE = '" & txtMobile & "' "
            
            rs.Open strQuery, con, 3, 2, 1
            
            If rs.Fields("MOB_COUNT") Then
                MsgBox "A contact with mobile number '" & txtMobile & "' already exists, Please choose another", vbInformation
                Exit Sub
            Else
                strQuery = "insert into CONTACTS(GROUPNAME, CONTACTNAME, MOBILE, DESIGNATION, EMAIL) values ('" & cmbGroups & "', '" & txtName & "', '" & txtMobile & "', '" & txtDesigantion & "', '" & txtEmail & "')"
                con.Execute strQuery
                frmChild.RefreshTree
                Unload Me
            End If
            rs.Close
            Set rs = Nothing
        End If
    End If
    frmMain.Enabled = True
    frmMain.Toolbar1.Enabled = True

handler:
'MsgBox Err.Number
Debug.Print Err.Number
If Err.Number = -2147217900 Then
MsgBox "Name already in use. Please choose another", vbInformation
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    Set rs = New ADODB.Recordset
    strQuery = "select GROUPNAME from GROUPS"
    
    rs.Open strQuery, con, 3, 2, 1
    
    While Not rs.EOF
        cmbGroups.AddItem rs.Fields("GROUPNAME")
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    cmbGroups.ListIndex = 0
    
    If SelectedNode <> "" Then
        While True
            If cmbGroups = SelectedNode Then
                GoTo finish
            Else
                cmbGroups.ListIndex = cmbGroups.ListIndex + 1
            End If
        Wend
    End If
   
finish:

'frmMain.Enabled = False
'frmMain.Toolbar1.Enabled = False


End Sub
Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub
