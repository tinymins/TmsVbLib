VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsExeUpgrader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************
'**模 块 名：clsTmsExeUpgrader
'**说    明：EXE版本更新类
'**创 建 人：翟一鸣 tinymins
'**网    站：ZhaiYiMing.CoM
'**日    期：2013年5月17日
'**备    注: 要求clsMD5类
'*************************************************************************
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'*********************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_HIDE = 0            '{隐藏}
Private Const SW_SHOWNORMAL = 1      '{用最近的大小和位置显示, 激活}
Private Const SW_NORMAL = 1          '{同 SW_SHOWNORMAL}
Private Const SW_SHOWMINIMIZED = 2   '{最小化, 激活}
Private Const SW_SHOWMAXIMIZED = 3   '{最大化, 激活}
Private Const SW_MAXIMIZE = 3        '{同 SW_SHOWMAXIMIZED}
Private Const SW_SHOWNOACTIVATE = 4  '{用最近的大小和位置显示, 不激活}
Private Const SW_SHOW = 5            '{同 SW_SHOWNORMAL}
Private Const SW_MINIMIZE = 6        '{最小化, 不激活}
Private Const SW_SHOWMINNOACTIVE = 7 '{同 SW_MINIMIZE}
Private Const SW_SHOWNA = 8          '{同 SW_SHOWNOACTIVATE}
Private Const SW_RESTORE = 9         '{同 SW_SHOWNORMAL}
Private Const SW_SHOWDEFAULT = 10    '{同 SW_SHOWNORMAL}
Private Const SW_MAX = 10            '{同 SW_SHOWNORMAL}
'*********************************************
' 发现新版本，更新开始。
Public Event OnUpgradeStart(ByRef bUpgradeContinue As Boolean)
' 没有新版本，更新终止
Public Event OnUpgradeAbort()
' 下载器就绪，准备更新
Public Event OnUpgradeReady()
' 更新过程出错
Public Event OnUpgradeError(ByVal iErrCode As Integer, ByVal sErrString As String)
'iErrCode    0.网络连接失败
'            1.最新VER文件格式错误
'            2.最新EXE文件下载失败
'           11.更新器VER文件错误
'           12.更新器EXE文件下载错误
'*********************************************
Public LatestInfo, LatestVersion, LatestExeUrl, LatestExeMd5, OnErrorUrl, UpgraderFilePath As String, LatestMajor, LatestMinor, LatestRevision As Integer
Private m_iUpgradeState

'构造函数
Private Sub Class_Initialize()
    LatestInfo = ""
    LatestVersion = ""
    LatestExeMd5 = ""
    LatestExeUrl = ""
    OnErrorUrl = ""
    m_iUpgradeState = 0
End Sub

'析构函数
Private Sub Class_Terminate()
    
End Sub

'说明：开始检测更新
'参数：
'   sVerUrl: 最新版本信息获取地址
'   sOnErrorUrl: 更新下载器过程中网络错误时弹出的页面
'   sUpgraderVerUrl: 下载器版本信息地址
Public Sub Start(ByVal sVerUrl As String, ByVal sOnErrorUrl As String, Optional sUpgraderVerUrl As String = "http://www.zhaiyiming.com/soft/upgrader/version.html")
    'On Error Resume Next
    m_iUpgradeState = 1
    OnErrorUrl = sOnErrorUrl
    
    Dim i As Integer, arrVer() As String
    
    Me.LatestInfo = Me.GetHttpResponse(sVerUrl & IIf(InStr(sVerUrl, "?") = 0, "?", "&") & "rnd=" + CStr(Now), "GBK")
    Me.LatestVersion = Me.Trim(Me.TmsExeVersionDecode(Me.LatestInfo, "ver"))
    Me.LatestExeMd5 = Me.Trim(Me.TmsExeVersionDecode(Me.LatestInfo, "md5"))
    Me.LatestExeUrl = Me.Trim(Me.TmsExeVersionDecode(Me.LatestInfo, "url"))
    
    If Me.LatestInfo = "" Then RaiseEvent OnUpgradeError(0, "获取最新版本号失败：网络连接错误"): Exit Sub
    If InStr(LCase(Me.LatestExeUrl), "http://") = 0 And InStr(LCase(Me.LatestExeUrl), "https://") = 0 Then Me.LatestExeUrl = Mid(sVerUrl, 1, InStrRev(sVerUrl, "/")) & Me.LatestExeUrl
    
    arrVer = Split(Me.LatestVersion, ".")
    If UBound(arrVer) - LBound(arrVer) + 1 < 3 Then
        RaiseEvent OnUpgradeError(1, "获取最新版本号失败：Version文件格式错误")
        Exit Sub
    Else
        Me.LatestMajor = CInt(arrVer(0))
        Me.LatestMinor = CInt(arrVer(1))
        Me.LatestRevision = CInt(arrVer(3))
    End If
    
    If Me.LatestMajor > App.Major Then GoTo startUpgrader
    If Me.LatestMajor = App.Major And Me.LatestMinor > App.Minor Then GoTo startUpgrader
    If Me.LatestMajor = App.Major And Me.LatestMinor = App.Minor And Me.LatestRevision > App.Revision Then GoTo startUpgrader
    GoTo abortUpgrader
    
startUpgrader:
    Dim bUpgradeContinue As Boolean
    bUpgradeContinue = True
    RaiseEvent OnUpgradeStart(bUpgradeContinue)
    If Not bUpgradeContinue Then Exit Sub
    ' 产生更新器绝对路径 Me.UpgraderFilePath
    Me.UpgraderFilePath = Me.AddStrToStrEnd(App.Path, "\") & "upgrader.exe"
    ' 验证upgrader已存在
    Dim fso As New FileSystemObject
    If fso.FileExists(Me.UpgraderFilePath) Then
        Dim sUpgraderInfo, sUpgraderMd5 As String
        sUpgraderInfo = Me.GetHttpResponse(sUpgraderVerUrl & IIf(InStr(sUpgraderVerUrl, "?") = 0, "?", "&") & "rnd=" + CStr(Now), "GBK")
        sUpgraderMd5 = Me.Trim(Me.TmsExeVersionDecode(sUpgraderInfo, "md5"))
        
        If Len(sUpgraderMd5) <> 32 Then GoTo shellUpgrader
        
        Dim md5 As New clsMD5
        If sUpgraderMd5 = md5.Md5_File_Calc(Me.UpgraderFilePath) Then
            GoTo shellUpgrader
        Else
            fso.DeleteFile Me.UpgraderFilePath
        End If
    End If
    ' 不存在则下载
    If Not GetInternetFile(Mid(sUpgraderVerUrl, 1, InStrRev(sUpgraderVerUrl, "/")) & "/upgrader.exe?rand=" & Now, Me.UpgraderFilePath) Then
        ShellExecute 0, "open", sOnErrorUrl, vbNullString, vbNullString, SW_SHOWNORMAL
        End
    End If
shellUpgrader:
    RaiseEvent OnUpgradeReady
    Exit Sub
abortUpgrader:
    RaiseEvent OnUpgradeAbort
End Sub

Public Function ShellEx(ByVal sPath As String, ByVal sParam As String)
    ShellEx = ShellExecute(0, "open", sPath, sParam, vbNullString, SW_SHOWNORMAL)
End Function

Public Function AddStrToStrEnd(ByVal Str1 As String, ByVal Str2 As String) As String
    '自动添加字符串到目标字符串结尾
    If LCase(Right(Str1, Len(Str2))) = LCase(Str2) Then
        AddStrToStrEnd = Str1
    Else
        AddStrToStrEnd = Str1 & Str2
    End If
End Function

Public Function RemoveStrFromStrStart(ByVal Str1 As String, ByVal Str2 As String) As String
    '从目标字符串首部删除字符串
    While LCase(Left(Str1, Len(Str2))) = LCase(Str2)
        Str1 = Mid(Str1, Len(Str2) + 1)
    Wend
    RemoveStrFromStrStart = Str1
End Function

'说明：获取网页源代码
'参数：
'   url: 链接地址
'   encoding: 页面编码,gb2312、utf-8等
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

'说明：字符串编码转换
'参数：
'   content: 文本
'   encoding:编码
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

'说明：下载网络文件
'参数：
'   SourceURL: 文件链接地址
'   DestFilePath: 文件保存路径
Public Function GetInternetFile(SourceURL As String, DestFilePath As String) As Boolean
    GetInternetFile = IIf(0 = URLDownloadToFile(0, SourceURL, DestFilePath, 0, 0), True, False)
    Exit Function
End Function

'说明：从Version.html中获取信息
'参数：
'   s_info: Version文件内容 字符串信息
'   s_seek: 寻找的字段
Public Function TmsExeVersionDecode(ByVal s_info As String, ByVal s_seek As String) As String
If 1 = 0 Then
notFound:
    TmsExeVersionDecode = ""
    Exit Function
End If
    Dim s_rtn As String
    If InStr(s_info, "|" + s_seek + "=") = 0 Then GoTo notFound
    s_rtn = Mid(s_info, InStr(s_info, "|" + s_seek + "=") + Len("|" + s_seek + "="))
    If InStr(s_rtn, "|") <> 0 Then s_rtn = Mid(s_rtn, 1, InStr(s_rtn, "|") - 1)
    TmsExeVersionDecode = s_rtn
End Function

'说明：去掉字符串左右的换行符空格制表以及\0
'参数：
'   str: 字符串
Public Function Trim(ByVal str As String) As String
    Dim i As Integer
    For i = 1 To Len(str) Step 1
        Select Case Mid(str, 1, 1)
        Case vbCr
        Case vbLf
        Case vbCrLf
        Case Chr(0)
        Case " "
        Case "  "
        Case Else
            Exit For
        End Select
        str = Mid(str, 2)
    Next
    For i = 1 To Len(str) Step 1
        Select Case Mid(str, Len(str), 1)
        Case vbCr
        Case vbLf
        Case vbCrLf
        Case Chr(0)
        Case " "
        Case "  "
        Case Else
            Exit For
        End Select
        str = Mid(str, 1, Len(str) - 1)
    Next
    Trim = str
End Function
