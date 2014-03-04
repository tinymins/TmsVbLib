Attribute VB_Name = "TmsMultiThreading"
Option Explicit
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private exitSignal As Integer
Private threadState() As Integer
Private thread(), ThreadID() As Long
Private threadCount As Long
Type TmsThreadParam
    addr As Long    '多线程地址
    i As Integer    '线程编号
End Type

' ::退出程序（主线程）
Public Sub TmsExitMainThread()
    exitSignal = 1
    With TmsThreadTimerForm.TmsMultiThreadTimer
        .Interval = 100
        .Enabled = True
    End With
End Sub

' ::创建线程 notice:线程结束务必设置对应threadState()为1
Public Sub TmsCreateThread(ByVal subThreadAddress As Long)
    threadCount = threadCount + 1
    ReDim Preserve threadState(threadCount)
    ReDim Preserve thread(threadCount)
    ReDim Preserve ThreadID(threadCount)
    
    Dim ttp As TmsThreadParam
    ttp.i = threadCount - 1
    ttp.addr = subThreadAddress
    
    Dim mp As TmsThreadParam
    mp.addr = 34259874
    mp.i = 1
    Dim ThreadID1 As Long
    Call CreateThread(ByVal 0&, ByVal 0&, AddressOf thMyFun, ByVal VarPtr(mp), ByVal 0&, ThreadID1)
    'thread(ttp.i) = CreateThread(ByVal 0&, ByVal 0&, AddressOf subThread, ByVal VarPtr(ttp), ByVal 0&, ThreadID(ttp.i))
    exitSignal = 0
    threadState(ttp.i) = 1
End Sub

Function thMyFun(p As TmsThreadParam) As Long
    MessageBox 0, p.i, p.addr, 0
End Function
Private Function subThread(ttp As TmsThreadParam) As Long
    'MsgBox ttp.i
   ' MsgBox ttp.addr
    'MsgBox ttp.s
    MessageBox 0, ttp.i, ttp.addr, 0
End Function

' ::获取对应线程运行状态
Public Function TmsGetThreadState(Optional i As Integer = -1)
    If i < 0 Or i > threadCount Then
        TmsGetThreadState = threadState
    Else
        TmsGetThreadState = threadState(i)
    End If
End Function
