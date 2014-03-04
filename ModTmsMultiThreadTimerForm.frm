VERSION 5.00
Begin VB.Form TmsThreadTimerForm 
   Caption         =   "TmsThreadTimerForm"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer TmsMultiThreadTimer 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "TmsThreadTimerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TmsMultiThreadTimer_Timer()
    'Dim i As Integer, threadState() As Long
    'threadState = TmsGetThreadState
    'For Each i In threadState
    '    If 1 = i Then Exit Sub
    'Next
    'End
End Sub
