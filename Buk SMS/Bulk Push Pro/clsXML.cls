VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################
'Programmer Name: Ronny Febrichrisanto, s.kom
'Email Address  : ilovejpop2002@yahoo.com
'Contact Number : 08561025823 - sms only
   
'syarat jika anda ingin menggunakan source ini yaitu :
'1. program ini jangan di jual dalam bentuk apapun!
'   Karena program ini gratis!
'2. jangan merubah nama pembuatnya : yaitu saya :-)
'3. berikan kredit, dengan tidak merubah nama pembuatnya :-)
'4. anda bebas memberikan source ini akan tetapi saya tidak
'   bertanggung jawab atas kerusakan yang terjadi pada computer user atau
'   pun masalah lisensi
'5. point no 4 di tanggung oleh user sendiri

'file ini bisa juga anda download di
'http://www.geocities.com/ronny_fch/download.htm
'#################################################

Option Explicit

Dim oDoc As MSXML2.DOMDocument
Dim bLoaded As Boolean

Private m_FileName As String
Private m_error As Integer
Public M_errCode As String

'Private str As String
'Private nilai As String

'Private urut As Integer

Public Event clserror(ByVal errorno As Integer)

'##################################################
'Fungsi untuk menampilkan data pembuat dll ini
Public Property Get About() As String
    About = "Ronny Febrichrisanto"
End Property

Public Property Let FileName(ByVal param As String)
    m_FileName = param
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Function KeySearch(param1 As String) As String
    On Error GoTo errs

    'xml object loaded
    Set oDoc = New DOMDocument
    oDoc.async = False
    oDoc.validateOnParse = True
    oDoc.Load "file:///" & m_FileName

    If oDoc.parseError.errorCode = 0 Then
        bLoaded = True
    Else
        MsgBox oDoc.parseError.reason & vbCrLf & oDoc.parseError.Line & vbCrLf & oDoc.parseError.srcText
    End If

    Dim result As IXMLDOMNodeList
    Set result = oDoc.selectNodes(param1)
'    MsgBox result.nextNode.nodeName
    KeySearch = "0"
    If result.length > 0 Then
        KeySearch = result.item(0).nodeTypedValue
    End If
    
    Exit Function
errs:
    m_error = Err.Number
    RaiseEvent clserror(m_error)
End Function

Public Function KeySearchADV(param, param1 As String) As String
    On Error GoTo errs

    'xml object loaded
    Set oDoc = New DOMDocument
    oDoc.async = False
    oDoc.validateOnParse = True
    oDoc.loadXML param

    If oDoc.parseError.errorCode = 0 Then
        bLoaded = True
    Else
        MsgBox oDoc.parseError.reason & vbCrLf & oDoc.parseError.Line & vbCrLf & oDoc.parseError.srcText
    End If

    Dim result As IXMLDOMNodeList
    Set result = oDoc.selectNodes(param1)
    KeySearchADV = "0"
    If result.length > 0 Then
        KeySearchADV = result.item(0).nodeTypedValue
    End If
    
    Exit Function
errs:
    m_error = Err.Number
    RaiseEvent clserror(m_error)
End Function

Public Function cekXML() As Boolean
    Set oDoc = New DOMDocument
    oDoc.async = False
    oDoc.validateOnParse = True
    oDoc.Load "file:///" & m_FileName
    cekXML = True
    If oDoc.parseError.errorCode <> 0 Then
        cekXML = False
        M_errCode = oDoc.parseError.reason & vbCrLf & oDoc.parseError.Line & vbCrLf & oDoc.parseError.srcText
    End If
End Function
