VERSION 5.00
Begin VB.Form frmContactEdit 
   Caption         =   "Edit Contact"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   Icon            =   "frmEditContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Group"
      Height          =   795
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   5655
      Begin VB.ComboBox cmbGroups 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   5355
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contact Information"
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   5655
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtMobile 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtDesigantion 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cell:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   315
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desigantion"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   840
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
         TabIndex        =   9
         Top             =   1080
         Width           =   375
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
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "frmContactEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strTempContactName As String
Public strTempGroup As String
Public strTempMobile As String


Private Sub cmdCancel_Click()
    Unload Me
    frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub

Private Sub cmdSave_Click()


    If Not IsNumeric(txtMobile) Then
    MsgBox "Invalid number"
    Exit Sub
    End If


    If txtName = "" Then
        MsgBox "Name cannot be empty", vbCritical
        Exit Sub
    End If
    
    If txtMobile = "" Then
        MsgBox "Mobile Number cannot be empty", vbCritical
        Exit Sub
    End If
    
    If cmbGroups = "" Then
        MsgBox "Please select the group", vbCritical
        Exit Sub
    End If
    If txtDesigantion = "" Or txtEmail = "" Then
        txtDesigantion.Text = "Unknown"
        txtEmail.Text = "Unknown"
    End If
    SaveContact cmbGroups, txtName, txtDesigantion.Text, txtEmail, txtMobile
    frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub

Private Sub Form_Load()
   
'    frmMain.Enabled = False
'    frmMain.Toolbar1.Enabled = False

End Sub

Public Sub FillForm(ByVal GroupName As String, ByVal ContactName As String)
    
    On Error Resume Next
    
    Dim rs As ADODB.Recordset
    Dim tmpGroupName As String
    Dim strQuery As String
    

    strQuery = "select * from CONTACTS where GROUPNAME = '" & GroupName & "' and CONTACTNAME = '" & ContactName & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open strQuery, con, 3, 2, 1
    
    If Not rs.EOF Then
            
        txtName = rs.Fields("CONTACTNAME")
        txtDesigantion = rs.Fields("DESIGNATION")
        txtMobile = rs.Fields("MOBILE")
        txtEmail = rs.Fields("EMAIL")
        
        tmpGroupName = rs.Fields("GROUPNAME")
        strTempContactName = rs.Fields("CONTACTNAME")
        strTempGroup = rs.Fields("GROUPNAME")
        strTempMobile = rs.Fields("MOBILE")
        rs.Close
        
        strQuery = "select GROUPNAME from GROUPS"
        rs.Open strQuery, con, 3, 2, 1
        
        While Not rs.EOF
            cmbGroups.AddItem rs.Fields("GROUPNAME")
            rs.MoveNext
        Wend
        
        cmbGroups.ListIndex = 0
        While True
            If cmbGroups = tmpGroupName Then
                GoTo finish
            Else
                 cmbGroups.ListIndex = cmbGroups.ListIndex + 1
            End If
        Wend
        
finish:
    End If
        
    'RefreshGroups frmMainGUI.TreeView1, rs

End Sub



Public Sub SaveContact(ByVal GroupName As String, ByVal Username As String, ByVal Designation As String, ByVal Email As String, ByVal Mobileno As String)
    Dim rs As ADODB.Recordset
    Dim strQuery As String
    
    Set rs = New ADODB.Recordset
    
    If GroupName = strTempGroup Then
        If Username = strTempContactName Then
            If strTempMobile = Mobileno Then
                'same group, same contact name, same mobile
'                MsgBox "update same group, same contact name, same mobile"
                UpdateContactInfo GroupName, Username, Designation, Email, Mobileno
                Unload Me
            Else
                strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & Mobileno & "' and GROUPNAME = '" & GroupName & "'"
                rs.Open strQuery, con, 3, 2, 1

                If rs.RecordCount Then
                    If rs.Fields("CONTACTNAME") <> strTempContactName Then
                        MsgBox "The contact '" & rs.Fields("CONTACTNAME") & "' with mobile number '" & Mobileno & "' already exists in '" & GroupName & "'"
                        rs.Close     '-----
                    End If
                Else
                    'same group, same contact name, NEW mobile
 '                   MsgBox "update same group, same contact name, NEW mobile"
                    UpdateContactInfo GroupName, Username, Designation, Email, Mobileno
                    Unload Me
              End If
            End If
        Else
            strQuery = "select CONTACTNAME from CONTACTS where CONTACTNAME = '" & Username & "' and GROUPNAME = '" & GroupName & "'"
            rs.Open strQuery, con, 3, 2, 1
            
            If rs.RecordCount Then
                MsgBox "A contact with name '" & Username & "' already exists in '" & GroupName & "'"
                rs.Close
            Else
                rs.Close
                If strTempMobile = Mobileno Then
                    'same group, NEW contact name, same mobile
                    'MsgBox "update same group, NEW contact name, same mobile"
                    UpdateContactInfo GroupName, Username, Designation, Email, Mobileno
                    Unload Me
                Else
                    strQuery = "select CONTACTNAME from CONTACTS where MOBILE = '" & Mobileno & "' and GROUPNAME = '" & GroupName & "'"
                    rs.Open strQuery, con, 3, 2, 1
    
                    If rs.RecordCount Then
                        If rs.Fields("CONTACTNAME") <> strTempContactName Then
                            MsgBox "The contact '" & rs.Fields("CONTACTNAME") & "' with mobile number '" & Mobileno & "' already exists in '" & GroupName & "'"
                            rs.Close '-----
                        End If
                    Else
                        'same group, NEW contact name, NEW mobile
                        'MsgBox "update same group, NEW contact name, NEW mobile"
                        UpdateContactInfo GroupName, Username, Designation, Email, Mobileno
                        Unload Me
                    End If
                End If
            End If
        End If
    Else
        strQuery = "select CONTACTNAME from CONTACTS where CONTACTNAME = '" & Username & "' and GROUPNAME = '" & GroupName & "'"
        rs.Open strQuery, con, 3, 2, 1
        
        If rs.RecordCount Then
            rs.Close
            MsgBox "The contact '" & Username & "' already exists in group '" & GroupName & "'"
        Else
            rs.Close
            
            strQuery = "select MOBILE from CONTACTS where MOBILE = '" & Mobileno & "' and GROUPNAME = '" & GroupName & "'"
            rs.Open strQuery, con, 3, 2, 1
            
            If rs.RecordCount Then
                rs.Close
                MsgBox "A contact with mobile number '" & Mobileno & "' exists in the destination group"
            Else
                rs.Close
                UpdateContactInfo GroupName, Username, Designation, Email, Mobileno
                Unload Me
           End If
        End If
        
    End If
    
    
End Sub

Private Sub UpdateContactInfo(ByVal GroupName As String, ByVal ContactName As String, ByVal Designation As String, ByVal Email As String, ByVal Mobileno As String)
    Dim strQuery As String
On Error GoTo handler
    strQuery = "update CONTACTS set CONTACTNAME = '" & ContactName & "', DESIGNATION = '" & Designation & "', EMAIL = '" & Email & "', MOBILE = '" & Mobileno & "', GROUPNAME = '" & GroupName & "' where GROUPNAME = '" & strTempGroup & "' and CONTACTNAME = '" & strTempContactName & "'"
    con.Execute strQuery
    
    frmChild.RefreshTree
handler:
Debug.Print Err.Number
If Err.Number = -2147217900 Then
MsgBox "Name already in use."
Exit Sub
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.Toolbar1.Enabled = True

End Sub
