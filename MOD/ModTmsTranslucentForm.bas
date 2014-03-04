Attribute VB_Name = "ModTmsTranslucentForm"
Option Explicit
'*************************************************************************
'**模 块 名：ModTmsTranslucentForm
'**说    明：界面（半）透明类
'**创 建 人：翟一鸣 tinymins
'**网    站：ZhaiYiMing.CoM
'**日    期：2013年5月17日
'**备    注: 版权木有，翻录不究。转载请保留本段文字。
'*************************************************************************
'*********************************************************
                    '窗体透明
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Boolean
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'*********************************************************
'窗体状态
Private crKey As Long
Private crKeySet As Long
Private bAlpha As Byte
Private bAlphaSet As Long
Public Function TranslucentForm(frmhwnd As Long, TranslucenceLevel As Byte) As Boolean
    '参数1为目标窗体的句柄,参数2为透明度0-255
    bAlpha = TranslucenceLevel
    bAlphaSet = LWA_ALPHA
    
    SetWindowLong frmhwnd, GWL_EXSTYLE, GetWindowLong(frmhwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmhwnd, crKey, bAlpha, (bAlphaSet And LWA_ALPHA) Or (crKeySet And LWA_COLORKEY)
    TranslucentForm = Err.LastDllError = 0
End Function
Public Function TranslucentColor(frmhwnd As Long, TranslucenceColor As Long) As Boolean
    '参数1为目标窗体的句柄,参数2为透明的颜色
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
'TranslucentForm Me.hwnd, 180  '将Form窗口设置为180的透明度
