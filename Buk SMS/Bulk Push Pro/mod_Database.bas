Attribute VB_Name = "mod_Database"
Public con As ADODB.Connection
Public rs As ADODB.Recordset

Public Sub ConnectDB()
    Set con = New ADODB.Connection
    con.Open "DSN=BPP;UID=admin;PWD=12345"
End Sub

Public Sub DisconnectDB()
    'If con.State = adOpenStatic Then
        con.Close
   ' End If
End Sub
