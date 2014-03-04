Attribute VB_Name = "ModShellEx"
'*************************************************************************
'**ģ �� ����ModShellEx
'**˵    ������ǿSHELL����
'**�� �� �ˣ�����
'**��    ������ˮ�������� http://www.m5home.com/
'**��    �ڣ�2007��4��24��
'**��    ����V1.0
'*************************************************************************
Option Explicit

Private Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ShellMod(ByVal FileName As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal DelayTime As Long = -1)
    '��SHELL����һ���Ĳ���,����������ִ��.(ͬ��)
    'FileName - Ŀ���ļ���
    'WindowStyle - ��������ʱ���ڵ���ʽ
    'DelayTime - �ȴ���ʱ��,��λΪms
    '��ע:
    '       DelayTime����Ϊ-1ʱ��ʾһֱ�ȴ�,ֱ��Ŀ��������н���
    Dim I As Long, J As Long
    
    I = Shell(FileName, WindowStyle)
    Do
        If GetProcessVersion(I) = 0 Then Exit Do
        Sleep 10
        J = J + 1
        If DelayTime <> -1 And J > DelayTime \ 10 Then Exit Do
    Loop
End Function

Public Function ShellModEx(ByVal FileName As String, Optional ByVal lpParameters As String = vbNullString, Optional ByVal DelayTime As Long = -1)
    '��SHELL����һ���Ĳ���,����������ִ��.(ͬ��)
    'FileName - Ŀ���ļ���
    'WindowStyle - ��������ʱ���ڵ���ʽ
    'DelayTime - �ȴ���ʱ��,��λΪms
    '��ע:
    '       DelayTime����Ϊ-1ʱ��ʾһֱ�ȴ�,ֱ��Ŀ��������н���
    Dim I As Long, J As Long
    I = ShellExecute(0, "open", FileName, lpParameters, vbNullString, 1)
    Do
        If GetProcessVersion(I) = 0 Then Exit Do
        Sleep 10
        J = J + 1
        If DelayTime <> -1 And J > DelayTime \ 10 Then Exit Do
    Loop
End Function

Public Function ShellOnce(ByVal FileName As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus)
    '��SHELL����һ���Ĳ���,��ֻ��Ŀ��ִ��һ��
    'FileName - Ŀ���ļ���
    'WindowStyle - ��������ʱ���ڵ���ʽ
    Static I As Long
    
    If I <> 0 Then          '�����PIDֵ���ж����Ƿ�����ִ��
        If GetProcessVersion(I) <> 0 Then Exit Function       '�������ִ��,��������
    End If
    I = Shell(FileName, WindowStyle)
End Function

