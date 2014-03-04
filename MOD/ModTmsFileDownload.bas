Attribute VB_Name = "ModTmsFileDownload"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function GetInternetFile(SourceURL As String, DestFilePath As String) As Boolean
    GetInternetFile = IIf(0 = URLDownloadToFile(0, SourceURL, DestFilePath, 0, 0), True, False)
    Exit Function
End Function
