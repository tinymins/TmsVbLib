Attribute VB_Name = "ModTmsTranslucentForm"
Option Explicit
'*************************************************************************
'**ģ �� ����ModTmsTranslucentForm
'**˵    �������棨�룩͸����
'**�� �� �ˣ���һ�� tinymins
'**��    վ��ZhaiYiMing.CoM
'**��    �ڣ�2013��5��17��
'**��    ע: ��Ȩľ�У���¼������ת���뱣���������֡�
'*************************************************************************
'*********************************************************
                    '����͸��
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Boolean
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'*********************************************************
'����״̬
Private crKey As Long
Private crKeySet As Long
Private bAlpha As Byte
Private bAlphaSet As Long
Public Function TranslucentForm(frmhwnd As Long, TranslucenceLevel As Byte) As Boolean
    '����1ΪĿ�괰��ľ��,����2Ϊ͸����0-255
    bAlpha = TranslucenceLevel
    bAlphaSet = LWA_ALPHA
    
    SetWindowLong frmhwnd, GWL_EXSTYLE, GetWindowLong(frmhwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmhwnd, crKey, bAlpha, (bAlphaSet And LWA_ALPHA) Or (crKeySet And LWA_COLORKEY)
    TranslucentForm = Err.LastDllError = 0
End Function
Public Function TranslucentColor(frmhwnd As Long, TranslucenceColor As Long) As Boolean
    '����1ΪĿ�괰��ľ��,����2Ϊ͸������ɫ
    crKey = TranslucenceColor
    crKeySet = LWA_COLORKEY
        
    SetWindowLong frmhwnd, GWL_EXSTYLE, GetWindowLong(frmhwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmhwnd, crKey, bAlpha, (bAlphaSet And LWA_ALPHA) Or (crKeySet And LWA_COLORKEY)
        
    TranslucentColor = Err.LastDllError = 0
End Function
Public Function getTranslucentLevel() As Byte
    getTranslucentLevel = bAlpha
End Function
Public Function getTranslucentColor() As Long
    getTranslucentColor = crKey
End Function
'*********************************************************
'TranslucentForm Me.hwnd, 180  '��Form��������Ϊ180��͸����
