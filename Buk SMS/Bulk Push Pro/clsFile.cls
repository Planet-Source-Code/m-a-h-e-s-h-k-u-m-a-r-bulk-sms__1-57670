VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
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

Const Buf As Integer = 30000
Dim FileLen As Long 'To keep file lenght information
Dim Multiply As Long 'It is required To find how many Buf
Dim Temp As String * Buf 'Temporary string block
Dim Content As String 'Content is the file content,
Dim Plus As Long 'The plus part of the file after dividing
Dim Point As Long 'Point shows on which Byte the content is.
Dim FileNo As Byte 'To find a free file number
Dim Counter As Long 'Is required For loops

Public Function ReadFile(FileName As String) As String 'Returns STRING variable!

    FileNo = FreeFile 'Find a free file number
    Open FileName For Random As #FileNo Len = Buf 'Open the file as Random, Each record will have the lenght of Buf
    FileLen = LOF(FileNo) 'File lenght
    Multiply = Int(FileLen \ Buf) 'How many loops required To read the file
    Content = Space(FileLen) 'Allocate a space For file content In the memory
    Plus = FileLen - (Multiply * Buf) 'After this loops, there might be also some bytes To read
    Point = 1 'Content is in this byte: 1

    If Multiply = 0 Then 'If the file is smaller than Buf (30000 bytes here, you can change it)
        Plus = FileLen: Counter = 1: GoTo Jump1
    End If

    For Counter = 1 To Multiply
        Get #FileNo, Counter, Temp
        Mid(Content, Point, Buf) = Temp
        Point = Point + Buf
    Next Counter

Jump1:

    If Plus > 0 Then
        Get #FileNo, Counter, Temp
        Mid(Content, Point, Plus) = Left(Temp, Plus)
    End If

    Close #FileNo
    ReadFile = Content
End Function


Public Sub WriteFile(FileName As String, Content As String)

    FileNo = FreeFile
    Open FileName For Output As #FileNo
    Print #FileNo, Content; '; is required For Vb To Not write another 2 charachters of new line in the file
    Close #FileNo
End Sub


'checking exists file ada atau tidak
Public Function FileExists(FullFileName As String) As Boolean

    On Error GoTo MakeF
     'If file does Not exist, there will be an Error
     Open FullFileName For Input As #1
     Close #1
     'no error, file exists
     FileExists = True
    Exit Function
MakeF:
     'error, file does Not exist
     FileExists = False
    Exit Function
End Function

