VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TmsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Function LTrim(ByVal str As String) As String
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
    LTrim = str
End Function

Public Function RTrim(ByVal str As String) As String
    Dim i As Integer
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
    RTrim = str
End Function

Public Function TrimEx(ByVal str As String, ByVal sTrim As String) As String
    Dim i As Integer
    For i = 1 To Len(str) Step 1
        Select Case Mid(str, 1, Len(sTrim))
        Case sTrim
        Case Else
            Exit For
        End Select
        str = Mid(str, Len(sTrim) + 1)
    Next
    For i = 1 To Len(str) Step 1
        Select Case Mid(str, Len(str) - Len(sTrim) + 1, Len(sTrim))
        Case sTrim
        Case Else
            Exit For
        End Select
        str = Mid(str, 1, Len(str) - Len(sTrim))
    Next
    TrimEx = str
End Function

Public Function LTrimEx(ByVal str As String, ByVal sTrim As String) As String
    Dim i As Integer
    For i = 1 To Len(str) Step 1
        Select Case Mid(str, 1, Len(sTrim))
        Case sTrim
        Case Else
            Exit For
        End Select
        str = Mid(str, Len(sTrim) + 1)
    Next
    LTrimEx = str
End Function

Public Function RTrimEx(ByVal str As String, ByVal sTrim As String) As String
    Dim i As Integer
    For i = 1 To Len(str) Step 1
        Select Case Mid(str, Len(str) - Len(sTrim) + 1, Len(sTrim))
        Case sTrim
        Case Else
            Exit For
        End Select
        str = Mid(str, 1, Len(str) - Len(sTrim))
    Next
    RTrimEx = str
End Function

