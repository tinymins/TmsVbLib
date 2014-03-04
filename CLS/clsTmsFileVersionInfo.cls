VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsFileVersionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ===== From Win32 Ver.h =================
' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&
' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20
' ----- VS_VERSION.dwFileOS -----
Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000
Private Const VOS__BASE = &H0
Private Const VOS__WINDOWS16 = &H1
Private Const VOS__PM16 = &H2
Private Const VOS__PM32 = &H3
Private Const VOS__WINDOWS32 = &H4
Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004
' ----- VS_VERSION.dwFileType -----
Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7
' ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA
Private Const VFT_FONT_RASTER = &H1&
Private Const VFT_FONT_VECTOR = &H2&
Private Const VFT_FONT_TRUETYPE = &H3&
Private Const OF_READ = &H0
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

Public Filename As String, Directory As String, FullFileName As String
Public StrucVer As String, FileVer As String, ProdVer As String
Public FileFlags As String, FileOS As String, FileType As String, FileSubType As String

Public Function GetFileInfo(ByVal sFullFileName As String)
    FullFileName = sFullFileName
    Dim rc As Long, lDummy As Long, sBuffer() As Byte
    Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
    Dim lVerbufferLen As Long
    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
    If lBufferLen < 1 Then
        'MsgBox "No Version Info available!"
        GetFileInfo = False
        Exit Function
    Else
        GetFileInfo = True
    End If
    '**** Store info to udtVerBuffer struct ****
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    '**** Determine Structure Version number - NOT USED ****
    StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)
    '**** Determine File Version number ****
    FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
    '**** Determine Product Version number ****
    ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
    '**** Determine Boolean attributes of File ****
    FileFlags = ""
    If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then FileFlags = "Debug "
    If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE Then FileFlags = FileFlags & "PreRel "
    If udtVerBuffer.dwFileFlags And VS_FF_PATCHED Then FileFlags = FileFlags & "Patched "
    If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD Then FileFlags = FileFlags & "Private "
    If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRED Then FileFlags = FileFlags & "Info "
    If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD Then FileFlags = FileFlags & "Special "
    If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN Then FileFlags = FileFlags + "Unknown "
    '**** Determine OS for which file was designed ****
    Select Case udtVerBuffer.dwFileOS
        Case VOS_DOS
            FileOS = "MS-DOS"
        Case VOS_OS216
            FileOS = "16-bit OS/2"
        Case VOS_OS232
            FileOS = "32-bit OS/2"
        Case VOS_NT
            FileOS = "Windows NT / Windows 2000"
        Case VOS__BASE
            FileOS = "BASE"
        Case VOS__WINDOWS16
            FileOS = "16-bit Windows"
        Case VOS__PM16
            FileOS = "16-bit Presentation Manager"
        Case VOS__PM32
            FileOS = "32-bit Presentation Manager"
        Case VOS__WINDOWS32
            FileOS = "Win32 API"
        Case VOS_DOS_WINDOWS16
            FileOS = "16-bit Windows running on MS-DOS"
        Case VOS_DOS_WINDOWS32
            FileOS = "32-bit Windows running on MS-DOS"
        Case VOS_OS216_PM16
            FileOS = "16-bit Presentation Manager running on 16-bit OS/2"
        Case VOS_OS232_PM32
            FileOS = "32-bit Presentation Manager running on 32-bit OS/2"
        Case VOS_NT_WINDOWS32
            FileOS = "Windows NT / Windows 2000"
        Case Else
            FileOS = "Unknown"
    End Select
    Select Case udtVerBuffer.dwFileType
        Case VFT_APP
            FileType = "App"
        Case VFT_DLL
            FileType = "DLL"
        Case VFT_DRV
            FileType = "Driver"
            Select Case udtVerBuffer.dwFileSubtype
                Case VFT2_DRV_PRINTER
                    FileSubType = "Printer drv"
                Case VFT2_DRV_KEYBOARD
                    FileSubType = "Keyboard drv"
                Case VFT2_DRV_LANGUAGE
                    FileSubType = "Language drv"
                Case VFT2_DRV_DISPLAY
                    FileSubType = "Display drv"
                Case VFT2_DRV_MOUSE
                    FileSubType = "Mouse drv"
                Case VFT2_DRV_NETWORK
                    FileSubType = "Network drv"
                Case VFT2_DRV_SYSTEM
                    FileSubType = "System drv"
                Case VFT2_DRV_INSTALLABLE
                    FileSubType = "Installable"
                Case VFT2_DRV_SOUND
                    FileSubType = "Sound drv"
                Case VFT2_DRV_COMM
                    FileSubType = "Comm drv"
                Case VFT2_UNKNOWN
                    FileSubType = "Unknown"
            End Select
        Case VFT_FONT
            FileType = "Font"
            Select Case udtVerBuffer.dwFileSubtype
                Case VFT_FONT_RASTER
                    FileSubType = "Raster Font"
                Case VFT_FONT_VECTOR
                    FileSubType = "Vector Font"
                Case VFT_FONT_TRUETYPE
                    FileSubType = "TrueType Font"
            End Select
        Case VFT_VXD
            FileType = "VxD"
        Case VFT_STATIC_LIB
            FileType = "Lib"
        Case Else
            FileType = "Unknown"
    End Select
End Function

