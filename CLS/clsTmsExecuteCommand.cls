VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsExecuteCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nsize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Declare Function GetSystemDirectoryW Lib "kernel32" (ByVal lpBuffer As Long, ByVal nsize As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'注释: 用于获取自windows启动以来经历的时间长度（毫秒） Long，返回值:以毫秒为单位的windows运行时间 2006-02-19
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Public Function ExecuteCommand(CommandLine As String) As String
'管道方式执行DOS命令并返回执行结果
    Const BuffSize As Long = 1024
    Dim proc As PROCESS_INFORMATION
    Dim Ret As Long
    Dim Start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long, lngBytesread As Long
    Dim strBuff As String * BuffSize, mOutputs As String
   
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    Ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
   
    If Ret = 0 Then ExecuteCommand = "CreatePipe Failed. Error: " & Err.LastDllError & vbCrLf: Exit Function
   
    Start.cb = Len(Start)
    Start.dwFlags = &H100 Or 1
    Start.hStdOutput = hWritePipe
    Start.hStdError = hWritePipe
    Ret& = CreateProcessA(0&, CommandLine, sa, sa, 1, &H20, 0, 0, Start, proc)
    If Ret <> 1 Then ExecuteCommand = "错误: '" & CommandLine & "' 不是可运行的程序" & vbCrLf: Exit Function
    Ret = CloseHandle(hWritePipe)

    mOutputs = ""
    Do
        Ret = ReadFile(hReadPipe, strBuff, BuffSize, lngBytesread, 0&)
        If Ret = 0 Then Exit Do
        mOutputs = mOutputs & LeftB(StrConv(strBuff, vbFromUnicode), lngBytesread)
    Loop
   
    Ret = CloseHandle(proc.hProcess)
    Ret = CloseHandle(proc.hThread)
    Ret = CloseHandle(hReadPipe)
   
    ExecuteCommand = Replace(StrConv(mOutputs, vbUnicode), Chr(0), "")
End Function


Public Function ExeCMD(ByVal CmdLine As String) As String
'2006-7-27 by VB爬虫
'以Shell方式执行Dos命令并输出结果
    Randomize
    Dim sTempFile As String
    sTempFile = UCase(Left$(GetSystemDirectory, 1))
    sTempFile = IIf(Asc(sTempFile) >= Asc("A") And Asc(sTempFile) <= Asc("Z"), Chr(Asc(sTempFile)), "C") & ":\ExeDosCMD." & DateDiff("s", "1970-01-01 00:00:00", Now) & CStr(Int(Rnd * 899) + 100) & ".tmp"
    
    Dim sOutPut As String, Temp As String
    Dim ErNum As Integer
    '删除临时文件 如果存在
    Shell "CMD /C Del " & sTempFile, vbHide 'vbMinimizedNoFocus
    '打开错误陷阱 以便捕捉错误
    On Error Resume Next
    '清除错误标记
    Err.Clear
    '检测命令行
    'ExeCMD = "您输入的命令是: [" & CmdLine & "]" & vbCrLf & String(50, "-") & vbCrLf
    If InStr(LCase(CmdLine), "cmd ") > 0 Then
        CmdLine = CmdLine & " > " & sTempFile
    Else
        CmdLine = "CMD /C " & CmdLine & " > " & sTempFile
    End If
    '以Shell方式执行命令行,并把输出结果重定向到临时文件
    Call Shell(CmdLine, vbMinimizedNoFocus)
    
    If Err.Number <> 0 Then
        ExeCMD = ExeCMD & "错误: " & Err.Number & " - " & Err.Description
        Exit Function
    End If
    
    ErNum = 0
REOPEN:
    Err.Clear
    Open sTempFile For Input Lock Read Write As #1
    If Err.Number = 53 Or Err.Number = 70 Then ' 文件未找到 / 文件正在被其他程序打开(没有权限)
        ErNum = ErNum + 1
        '记录并判断,防止程序无限期等待下去
        If ErNum > 65535 Then
            ExeCMD = "错误: 操作已超时!"
            Exit Function
        End If
        DoEvents
        Sleep 50
        GoTo REOPEN
    ElseIf Err.Number <> 0 Then
        ExeCMD = ExeCMD & "错误: " & Err.Number & " - " & Err.Description
        Exit Function
    End If
    Err.Clear
    '开始读取临时文件的内容(命令的输出结果)
    While Not EOF(1)
        '读取一行的内容
        Line Input #1, Temp
        sOutPut = sOutPut + Temp + vbCrLf
    Wend
    '关闭临时文件
    Close #1
    '删除临时文件
    Shell "CMD /C Del " & sTempFile, vbMinimizedNoFocus

    If Err.Number <> 0 Then
        ExeCMD = ExeCMD & "错误: " & Err.Number & " - " & Err.Description
        Exit Function
    Else
        ExeCMD = ExeCMD & sOutPut
    End If
End Function

Public Function GetSystemDirectory() As String
    Dim sSystemDirectory As String, lRet As Long
    sSystemDirectory = Space(255)
    lRet = GetSystemDirectoryW(StrPtr(sSystemDirectory), 255)
    sSystemDirectory = Left$(sSystemDirectory, lRet)
    GetSystemDirectory = sSystemDirectory
End Function

