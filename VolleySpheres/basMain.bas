Attribute VB_Name = "basMain"
Option Explicit

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal fdwError As Long, ByVal lpszErrorText As String, ByVal cchErrorText As Long) As Long

Public Const AppGuid = "{AC330441-9B71-11D2-9AAB-0020781461AC}"
Public Const AppName = "Direct Chat"
Public Const ChatMessage = 411

Public DX7 As New DxVBLib.DirectX7
Public DPlay As DxVBLib.DirectPlay4
Public DPEnumSessions As DxVBLib.DirectPlayEnumSessions
Public DPEnumConnections As DxVBLib.DirectPlayEnumConnections
Public DPEnumPlayers As DxVBLib.DirectPlayEnumPlayers

Public MyPlayerID As Long
Public NotificationID As Long

Public PlayerHandle As String
Public PlayerName As String

Public Sub Init()
    
On Error GoTo Handler
    
    Set DPlay = DX7.DirectPlayCreate("")
    Set DPEnumConnections = DPlay.GetDPEnumConnections("", DPCONNECTION_DIRECTPLAY)
    
 
    Exit Sub
    
Handler:
    MsgBox "Failed to initialize DirectPlay.", vbCritical, AppName
    Form2.Command2.Enabled = False
    End
End Sub


Public Function SendMCIString(Cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(Cmd, 0, 0, Form2.hWnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function



