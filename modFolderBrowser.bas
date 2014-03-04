Attribute VB_Name = "modFolderBrowser"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Const MAX_PATH = 260

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 3

Private Const WM_USER = &H400

Private Const BFFM_SETSTATUSTEXT As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)
   
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Function GetFolder(ByVal title As String, ByVal start As String, ByVal newfolder As Boolean) As String
    Dim BI As BROWSEINFO, pidl As Long, lpSelPath As Long
    Dim spath As String * MAX_PATH
    
    'fill in the info it needs
    With BI
        .hOwner = GetForegroundWindow
        .pidlRoot = 0
        .lpszTitle = title
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .ulFlags = BIF_RETURNONLYFSDIRS
        If newfolder = True Then .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
        If start <> "" Then
            lpSelPath = LocalAlloc(LPTR, Len(start) + 1)
            CopyMemory ByVal lpSelPath, ByVal start, Len(start) + 1
            .lParam = lpSelPath
        End If
    End With
    
    'get the idlist long from the returned folder
    pidl = SHBrowseForFolder(BI)
    
    'do then if they clicked ok
    If pidl Then
        If SHGetPathFromIDList(pidl, spath) Then
            'next line is the returned folder
            GetFolder = Left$(spath, InStr(spath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(pidl)
    Else
        'user clicked cancel
    End If
    
    Call LocalFree(lpSelPath)
    
End Function

'this seems to happen before the box comes up and when a folder is clicked on within it
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim spath As String, bFlag As Long
                                       
    spath = Space$(MAX_PATH)
        
    Select Case uMsg
        Case BFFM_INITIALIZED
            'browse has been initialized, set the start folder
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal lpData)
        Case BFFM_SELCHANGED
            If SHGetPathFromIDList(lParam, spath) Then
                spath = Left(spath, InStr(1, spath, Chr(0)) - 1)
            End If
    End Select
          
End Function
          
Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn

End Function

