Attribute VB_Name = "ModTmsHttpRequest"
Option Explicit
'˵������ȡ��ҳԴ����
'������
'   url: ���ӵ�ַ
'   encoding: ҳ�����,gb2312��utf-8��
Public Function GetHttpResponse(ByVal url As String, ByVal encoding As String) As String
    Dim xmlHTTP As Object
    Dim content As Variant
    On Error Resume Next
    Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP.Open "GET", url, True
    xmlHTTP.send
    While xmlHTTP.readyState <> 4
        DoEvents
    Wend
    content = xmlHTTP.responseBody
    If CStr(content) <> "" Then GetHttpResponse = EncodingConvertor(content, encoding)
    Set xmlHTTP = Nothing
    If Err.Number <> 0 Then
        GetHttpResponse = ""
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

