VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsThreading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Sub SleepA Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private iDoEventsCount As Integer

Public Sub Sleep(ByVal dwMilliseconds As Long)
    SleepA dwMilliseconds
    DoEvent
End Sub

Public Sub DoEvent()
    '来代替doevents可使循环运行更快
    If GetInputState() Or iDoEventsCount > 10 Then
        DoEvents
        iDoEventsCount = 0
    Else
        iDoEventsCount = iDoEventsCount + 1
    End If
End Sub

