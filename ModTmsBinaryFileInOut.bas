Attribute VB_Name = "BinaryFileInOut"
Public byteFile() As Byte
Public Function BinaryFileRead(ByRef byteFile() As Byte, ByVal sFilePath As String) As Boolean
    Dim fso As New FileSystemObject  '����Microsoft Scripting Runtime
    If Not fso.FileExists(sFilePath) Then
err1:   BinaryFileRead = False
        Exit Function
    End If
    On Error GoTo err1

    Dim intFile As Integer      '�ļ���
    Dim lngDatLength As Single  '�ļ����ȣ��ֽڣ�
    intFile = FreeFile()        '�����µ��ļ���
    Open sFilePath For Binary As intFile
        'Seek #1, 22
        lngDatLength = LOF(intFile) '�ļ����ȣ��ֽ�����
        ReDim byteFile(1 To lngDatLength)
        Get #1, , byteFile
    Close intFile
    
    BinaryFileRead = False
End Function
    
Public Function BinaryFileWrite(ByRef byteFile() As Byte, ByVal sFilePath As String) As Boolean
    If 1 = 2 Then
err1:   BinaryFileWrite = False
        Exit Function
    End If
    On Error GoTo err1

    Dim intFile As Integer      '�ļ���
    Dim lngDatLength As Single  '�ļ����ȣ��ֽڣ�
    intFile = FreeFile()        '�����µ��ļ���
    Open sFilePath For Binary As intFile
        Put intFile, , byteFile
    Close intFile
    
    BinaryFileWrite = True
End Function
