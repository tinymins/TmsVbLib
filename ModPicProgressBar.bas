Attribute VB_Name = "ModTmsPicProgressBar"
Option Explicit

Public Function setPicProgressBarStatus(picProgressBar As PictureBox, ByVal sStatus As String)
    With picProgressBar
        .ToolTipText = Left(.ToolTipText, InStr(.ToolTipText, "%") - 1) & "%" & sStatus
    End With
    redrawPicProgressBar picProgressBar
End Function

Public Sub setPicProgressBarPercent(picProgressBar As PictureBox, ByVal percent As Integer)
    If percent < 0 Then percent = 0
    If percent > 100 Then percent = 100
    With picProgressBar
        .ToolTipText = str(percent) & "%" & Mid(.ToolTipText, InStr(.ToolTipText, "%") + 1)
    End With
    redrawPicProgressBar picProgressBar
End Sub

Private Sub redrawPicProgressBar(picProgressBar As PictureBox)
    Dim progressPercent As Integer
    Dim progressCaption As String
    With picProgressBar
        progressPercent = Int(Left(.ToolTipText, InStr(.ToolTipText, "%") - 1))
        progressCaption = Mid(.ToolTipText, InStr(.ToolTipText, "%") + 1)
        .AutoRedraw = True
        picProgressBar.Scale (0, 0)-(100, 10)  ' 設定坐標系
        '.Font.Name = "System"               ' 設定字體
        '.Font.Size = 12
        
        ' 畫出色塊
        .Cls
        picProgressBar.Line (0, 0)-(progressPercent, 10), RGB(0, 200, 200), BF
        ' 顯示文字
        .CurrentX = (100 - .TextWidth(progressCaption)) / 2
        .CurrentY = (10 - .TextHeight(progressCaption)) / 2 - .BorderStyle * (40 / .Height)
        picProgressBar.Print progressCaption
    End With
End Sub

