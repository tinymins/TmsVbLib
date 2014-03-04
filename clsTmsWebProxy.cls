VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsWebProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents sck_ As Winsock
Attribute sck_.VB_VarHelpID = -1
Dim bRecieveComplete As Boolean
Dim sRecievedData As String

'构造函数
Private Sub Class_Initialize()
    
End Sub

'析构函数
Private Sub Class_Terminate()
    Set sck_ = Nothing
End Sub


Public Function GetHttpRequestViaSocket(ByVal sHost As String, ByVal sData As String, Optional iPort As Integer = 80) As String
    Dim iErrCount As Integer
    bRecieveComplete = False
    sRecievedData = ""
    
    sck_.Protocol = sckTCPProtocol
    sck_.RemoteHost = sHost
    sck_.RemotePort = iPort
    sck_.Connect
    
    While sck_.State <> sckConnected And iErrCount < 65535
        DoEvents
        iErrCount = iErrCount + 1
    Wend
    
    sck_.SendData sData
    
    While bRecieveComplete <> True And iErrCount < 65535
        DoEvents
        iErrCount = iErrCount + 1
    Wend
    
    GetHttpRequestViaSocket = sRecievedData
End Function

Public Function PostHttpRequestViaSocket(ByVal sUrl As String, ByVal sData As String)

End Function

Private Sub sck__DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    sck_.GetData sData
    sRecievedData = sRecievedData & sData
End Sub

