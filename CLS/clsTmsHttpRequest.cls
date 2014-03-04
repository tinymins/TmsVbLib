VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim xmlHTTP_ As Object

Public WithEvents sck_ As Winsock
Attribute sck_.VB_VarHelpID = -1

'���캯��
Private Sub Class_Initialize()
    Set xmlHTTP_ = CreateObject("Microsoft.XMLHTTP")
End Sub

'��������
Private Sub Class_Terminate()
    Set xmlHTTP_ = Nothing
End Sub

'˵������ȡ��ҳԴ����
'������
'   url: ���ӵ�ַ
'   encoding: ҳ�����,gb2312��utf-8��
Public Function HttpGet(ByVal url As String, ByVal encoding As String) As String
    Dim content As Variant
    On Error Resume Next
    'Set xmlHTTP_ = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP_.Open "GET", url, True
    xmlHTTP_.send
    While xmlHTTP_.readyState <> 4
        DoEvents
    Wend
    content = xmlHTTP_.responseBody
    If CStr(content) <> "" Then HttpGet = EncodingConvertor(content, encoding)
    'Set xmlHTTP_ = Nothing
    If Err.Number <> 0 Then
        HttpGet = ""
    End If
    On Error GoTo 0
End Function

'˵������ȡ��ҳԴ����
'������
'   url: ���ӵ�ַ
'   encoding: ҳ�����,gb2312��utf-8��
Public Function HttpPost(ByVal url As String, ByVal postData As String, ByVal encoding As String) As String
    Dim content As Variant
    On Error Resume Next
    'Set xmlHTTP_ = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP_.Open "POST", url, True
    xmlHTTP_.setRequestHeader "Content-Length", Len(postData)
    xmlHTTP_.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"  '�����������ݱ����ʽ
    xmlHTTP_.send (postData)
    While xmlHTTP_.readyState <> 4
        DoEvents
    Wend
    content = xmlHTTP_.responseBody
    If CStr(content) <> "" Then HttpPost = EncodingConvertor(content, encoding)
    'Set xmlHTTP_ = Nothing
    If Err.Number <> 0 Then
        HttpPost = ""
    End If
    On Error GoTo 0
End Function

'˵������ȡ��ҳԴ����
'������
'   url: ���ӵ�ַ
'   encoding: ҳ�����,gb2312��utf-8��
Public Function PostViaSocket(ByVal url As String, ByVal postData As String, ByVal encoding As String) As String
    Dim content As Variant
    On Error Resume Next
    'Set xmlHTTP_ = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP_.Open "POST", url, True
    xmlHTTP_.setRequestHeader "Content-Length", Len(postData)
    xmlHTTP_.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"  '�����������ݱ����ʽ
    xmlHTTP_.send (postData)
    While xmlHTTP_.readyState <> 4
        DoEvents
    Wend
    content = xmlHTTP_.responseBody
    If CStr(content) <> "" Then HttpPost = EncodingConvertor(content, encoding)
    'Set xmlHTTP_ = Nothing
    If Err.Number <> 0 Then
        HttpPost = ""
    End If
    On Error GoTo 0
End Function

'˵�����ַ�������ת��
'������
'   content: �ı�
'   encoding:����
Public Function EncodingConvertor(ByVal content As Variant, ByVal encoding As String) As String
    Dim objStream As Object
    On Error Resume Next
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
        .Type = 1
        .Mode = 3
        .Open
        .Write content
        .Position = 0
        .Type = 2
        .Charset = encoding
        EncodingConvertor = .ReadText
        .Close
    End With
    Set objStream = Nothing
    If Err.Number <> 0 Then
        EncodingConvertor = ""
    End If
    On Error GoTo 0
End Function

