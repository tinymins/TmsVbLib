Attribute VB_Name = "ModTmsByteOperator"
Option Explicit

Public Function strpos(ByRef buf_src() As Byte, ByVal str_tar As String)
    Dim buf_tar() As Byte
    buf_tar = str2byte(str_tar)
    
    Dim i As Long
    For i = 0 To UBound(buf_src) - UBound(buf_tar)
        Dim j As Integer
        For j = 0 To UBound(buf_tar)
            If buf_src(i + j) <> buf_tar(j) Then Exit For
        Next
        If j >= UBound(buf_tar) Then
            strpos = i
            Exit Function
        End If
    Next
    strpos = -1
End Function

Public Function bytes_str_replace(ByRef buf_src() As Byte, ByVal str_tar As String, ByVal l_start_pos As Long, Optional l_replace_len As Long = &H7FFFFFFF)
    If l_replace_len > Len(str_tar) Then l_replace_len = Len(str_tar)
    If l_replace_len < 0 Then l_replace_len = 0
    
    Dim buf_tar() As Byte
    buf_tar = str2byte(str_tar)
    
    Dim i As Long
    For i = 0 To l_replace_len - 1
        buf_src(l_start_pos + i) = buf_tar(i)
    Next
End Function



Public Function str2byte(ByVal str_src) As Byte()
    Dim buf_src() As Byte
    
    ReDim buf_src(Len(str_src))
    
    Dim i As Long
    For i = 0 To UBound(buf_src) - 1
        buf_src(i) = Asc(Mid(str_src, i + 1, 1))
    Next
    
    str2byte = buf_src
End Function
