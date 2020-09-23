VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MsWinSck.ocx"
Begin VB.UserControl Proxy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   5145
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2445
      Top             =   525
   End
   Begin MSWinsockLib.Winsock Outgoing 
      Index           =   0
      Left            =   1425
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Incomming 
      Index           =   0
      Left            =   975
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1995
      Top             =   525
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   480
      Picture         =   "UserControl1.ctx":0312
      Top             =   480
      Width           =   420
   End
End
Attribute VB_Name = "Proxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type var_iables
      IncommingState      As Integer
      OutgoingState       As Integer
      lServIndex          As Long
      RemotePort          As String
      httpRemAddr         As String
      incomingData        As String
      outgoingData        As String
      arrProxyParse()     As String
      bExitLoop           As Boolean
      bAlterOutgoing      As Boolean
End Type

Public AddressPOP$, AddressSMTP$
Public bHttpProxyAddrExtracted    As Boolean

Private Const WM_SYSCOMMAND As Long = &H112
Private Const IE_REGPATH = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"

Public Enum CT
    KeepAlive = 0
    Closed = 1
    DontDisplay = 2
End Enum

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()

Public Enum SockStatus
    i_Closed = 0
    i_Opened = 1
    i_Listening = 2
    i_Pending = 3
    i_Resolving = 4
    i_Resolved = 5
    i_Connecting = 6
    i_Connected = 7
    i_Closing = 8
    i_Error = 9
End Enum

Private Type IncommingConns
    OutGoingIndex As Integer
    ReceivedData As String
    Active As Boolean
    Tick As Long
End Type

Private Type OutgoingConns
    IncommingIndex As Integer
    ReceivedData As String
    Active As Boolean
    Tick As Long
End Type

Public Enum GP
      HKEY_CURRENT_USER = &H80000001
      HKEY_LOCAL_MACHINE = &H80000002
      HKEY_USERS = &H80000003
End Enum

Public Enum RV
      REG_SZ = 1
      REG_BINARY = 3
      REG_DWORD = 4
End Enum

Public Enum FindIndex
    i_Incomming = 0
    i_Outgoing = 1
End Enum

Public Event StatusChangeOutgoing(Current_State As SockStatus)
Public Event StatusChangeIncomming(Current_State As SockStatus)
Public Event FowaredRequest(RequestID As Integer, SockType As FindIndex, Data As String)
Public Event SockOpen(SockType As FindIndex, Index As Integer)
Public Event SockClosed(SockType As FindIndex, Index As Integer)
Public Event PageChange(Page As String)

Private Var                 As var_iables
Private IncommingConns(255) As IncommingConns
Private OutgoingConns(255)  As OutgoingConns

Function FindOpenIndex(IndexType As FindIndex) As Integer
Dim y
If IndexType = i_Incomming Then
    For y = 1 To 255
    If IncommingConns(y).Active = False Then
    FindOpenIndex = y
    Exit Function
    End If
    Next
ElseIf IndexType = i_Outgoing Then
    For y = 1 To 255
    If OutgoingConns(y).Active = False Then
    FindOpenIndex = y
    Exit Function
    End If
    Next
End If
End Function

Private Sub Incomming_Close(Index As Integer)
RaiseEvent SockClosed(i_Incomming, Index)
End Sub

Function AcceptRequest(RequestID As Integer, SockType As FindIndex, Data As String)
If SockType = i_Incomming Then
    If Incomming(RequestID).State = 7 Then Incomming(RequestID).SendData Data & vbCrLf
    RaiseEvent SockOpen(i_Incomming, RequestID)
ElseIf SockType = i_Outgoing Then
    If Outgoing(RequestID).State = 7 Then Outgoing(RequestID).SendData Data
End If
End Function

Private Sub Incomming_Connect(Index As Integer)
        RaiseEvent FowaredRequest(Index, i_Incomming, OutgoingConns(Index).ReceivedData)
End Sub

Private Sub Incomming_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   On Error Resume Next
   Incomming(Index).GetData Var.incomingData
   IncommingConns(Index).ReceivedData = Var.incomingData
   RaiseEvent FowaredRequest(Index, i_Outgoing, IncommingConns(Index).ReceivedData)
End Sub

Private Sub Outgoing_Close(Index As Integer)
RaiseEvent SockClosed(i_Outgoing, Index)
End Sub

Private Sub Outgoing_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
    Dim trimRem$
    If Index = 0 Then
        trimRem = Trim(Outgoing(0).RemoteHostIP)
        If trimRem = "127.0.0.1" Or trimRem = Outgoing(0).LocalIP Then
            Var.bExitLoop = False
            Dim f As Integer: f = FindOpenIndex(i_Outgoing)
            Outgoing(f).Close
            Outgoing(f).Accept RequestID
            RaiseEvent SockOpen(i_Outgoing, f)
            Outgoing(0).Close
            Outgoing(0).Listen
            Do
                DoEvents
            Loop Until Outgoing(f).State = 7 Or Var.bExitLoop = True
       Else
          Outgoing(0).Close
          Outgoing(0).Listen
       End If
    End If
End Sub

Private Sub Outgoing_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Ret
    Ret = GetTickCount&
    OutgoingConns(Index).Tick = Int(Ret / 600)
    OutgoingConns(Index).Active = True
    Outgoing(Index).GetData Var.outgoingData
    Var.arrProxyParse = ExtractHttp(Var.outgoingData, DontDisplay)
    OutgoingConns(Index).ReceivedData = Var.arrProxyParse(2)
    Incomming(Index).Close
    Incomming(Index).Connect Var.arrProxyParse(0), 80
    RaiseEvent PageChange(Var.arrProxyParse(0))
End Sub

Private Sub Timer1_Timer()
    '--------------------------------DISPLAYS STATUS-------------------------------'
    UpdateStates
End Sub

Private Sub Timer2_Timer()
    '--------------------------------CLOSES UNNEEDED SOCKS-------------------------------'
Dim p As Integer, Ret
For p = 1 To 255
  Ret = GetTickCount&
    Ret = Int(Ret / 600)
If Incomming(p).State = 8 Or Incomming(p).State = 9 Then
    Incomming(p).Close
    RaiseEvent SockClosed(i_Incomming, p)
End If

If Outgoing(p).State = 8 Or Outgoing(p).State = 9 Then
    Outgoing(p).Close: Incomming(p).Close
    OutgoingConns(p).Tick = 0
    OutgoingConns(p).Active = False
    IncommingConns(p).Active = False
    RaiseEvent SockClosed(i_Incomming, p)
    RaiseEvent SockClosed(i_Outgoing, p)
End If

If OutgoingConns(p).Tick <> 0 And Ret - OutgoingConns(p).Tick > 15 Then
    OutgoingConns(p).Tick = 0
    Outgoing(p).Close: Incomming(p).Close
    OutgoingConns(p).Active = False
    IncommingConns(p).Active = False
    RaiseEvent SockClosed(i_Incomming, p)
    RaiseEvent SockClosed(i_Outgoing, p)
End If
Next

End Sub

Private Sub UserControl_Initialize()
    '--------------------------------ALIGNS TO PICTURE-------------------------------'
    Image1.Top = 0
    Image1.Left = 0
    UserControl.Height = Image1.Height
    UserControl.Width = Image1.Width
    '--------------------------------PRELOADS WINSOCKS-------------------------------'
    Dim x
    For x = 1 To 255
        Load Outgoing(x)
        Load Incomming(x)
    Next
End Sub

Private Sub UserControl_Resize()
    '--------------------------------ALIGNS TO PICTURE-------------------------------'
    Image1.Top = 0
    Image1.Left = 0
    UserControl.Height = Image1.Height
    UserControl.Width = Image1.Width
End Sub

Private Sub UserControl_Terminate()
    '--------------------------------STOPS ALL-------------------------------'
  Var.bExitLoop = True
  DisableProxy
End Sub
Function Start()
    '--------------------------------STARTS ALL-------------------------------'
    Call EnableProxy(80)
    Call StartProxy
End Function

Function StopProxy()
    '--------------------------------STOPS ALL-------------------------------'
    Var.bExitLoop = True
    DisableProxy
    Incomming(0).Close
    Outgoing(0).Close
    Timer1.Enabled = False
    Timer2.Enabled = False
End Function

Function Get_Socket_State(Index As Integer, SockType As FindIndex) As Integer
    '--------------------------------RETURNS SOCKET STATE-------------------------------'
    If SockType = i_Outgoing Then
        Get_Socket_State = Outgoing(Index).State
    ElseIf SockType = i_Incomming Then
        Get_Socket_State = Incomming(Index).State
    End If
End Function

Function UpdateStates()
    '--------------------------------OUTGOING CHECK-------------------------------'
    If Var.OutgoingState <> Outgoing(0).State Then
        Var.OutgoingState = Outgoing(0).State: RaiseEvent StatusChangeOutgoing(Outgoing(0).State)
    End If
    If Outgoing(0).State = 9 Or Outgoing(0).State = 8 Then Call StartProxy
    
    '--------------------------------INCOMMING CHECK-------------------------------'
    If Var.IncommingState <> Incomming(0).State Then
        Var.IncommingState = Incomming(0).State: RaiseEvent StatusChangeIncomming(Incomming(0).State)
    End If
    If Incomming(0).State = 9 Or Incomming(0).State = 8 Then Call StartProxy
End Function

Private Sub StartProxy()
    '--------------------------------RESTART ALL WINSOCKS-------------------------------'
    Var.RemotePort = 80
    Timer1.Enabled = True
    Timer2.Enabled = True
    Outgoing(0).Close
    Outgoing(0).LocalPort = Var.RemotePort
    Outgoing(0).Listen
    Incomming(0).Close
End Sub

Private Sub DisableProxy()
    '--------------------------------DISABLE PROXY-------------------------------'
    RegSaveDword HKEY_CURRENT_USER, IE_REGPATH, "ProxyEnable", "0"
End Sub

Private Sub EnableProxy(sProxyPort$)
    '--------------------------------ENABLE PROXY-------------------------------'
    RegSaveDword HKEY_CURRENT_USER, IE_REGPATH, "ProxyEnable", "1"
    RegSaveString HKEY_CURRENT_USER, IE_REGPATH, "ProxyServer", "127.0.0.1:" & sProxyPort
End Sub

Sub RegSaveString(hKey As GP, strPath$, strValName$, strNewVal$)
    '--------------------------------SAVE STRING-------------------------------'
    Dim Ret
    RegOpenKey hKey, strPath, Ret
    RegSetValueEx Ret, strValName, 0, REG_SZ, ByVal strNewVal, Len(strNewVal)
    RegCloseKey Ret
End Sub
Sub RegSaveBinaray(hKey As GP, strPath$, strValName$, strNewVal$)
    '--------------------------------SAVE BINARY-------------------------------'
    Dim Ret
    RegOpenKey hKey, strPath, Ret
    RegSetValueEx Ret, strValName, 0, REG_BINARY, ByVal strNewVal, Len(strNewVal)
    RegCloseKey Ret
End Sub
Sub RegSaveDword(hKey As GP, strPath$, strValName$, strNewVal$)
    '--------------------------------SAVE DWROD-------------------------------'
    Dim Ret
    RegOpenKey hKey, strPath, Ret
    RegSetValueEx Ret, strValName, 0, REG_DWORD, CLng(strNewVal), 4
    RegCloseKey Ret
End Sub
Function GetString(hKey As GP, strPath$, strValName As String) As String
    '--------------------------------GET REG VALUE-------------------------------'
    Dim Ret
    RegOpenKey hKey, strPath, Ret
    GetString = RegQueryStringValue(Ret, strValName)
    RegCloseKey Ret
End Function

Function ExtractHttp(sHTTPrequest$, Optional ConnectionType As CT = KeepAlive) As String()
    '--------------------------------GET HTTP VALUE-------------------------------'
    Dim sParts() As String, arr(2) As String
    Dim sUrl$, sEndLine$, line1$, topLineReassembled$
    Dim L&
    If Len(Trim(sHTTPrequest$)) = 0 Then
        Err.Raise 2233, "Function func_ExtractHttp", "argument: sData$ not provided"
        Exit Function
    End If
    line1 = Split(sHTTPrequest, vbCrLf)(0)
    sParts = Split(line1, "/")
    For L = 2 To UBound(sParts) - 1
         If L = 2 Then
            arr(0) = sParts(L)
         Else
             If Trim(sParts(L)) <> "HTTP" Then
               arr(1) = (arr(1) & "/" & sParts(L))
            End If
         End If
    Next L
    arr(1) = Trim(Replace(arr(1), "HTTP", ""))
    topLineReassembled = "GET " & arr(1) & "/ HTTP/1.0"
    sParts = Split(sHTTPrequest$, vbCrLf)
    For L = 1 To UBound(sParts)
       If L = 3 Then
           arr(2) = (arr(2) & "'~~~~~~~~~~~~~~~: ~~~~~ ~~~~~~~" & vbCrLf)
       End If
       
       If Trim(Left(sParts(L), 6)) <> "Proxy-" Then
          If Trim(sParts(L)) <> "" Then
              
              arr(2) = (arr(2) & sParts(L) & vbCrLf)
          End If
       End If
    Next L
    If ConnectionType = KeepAlive Then
        sEndLine = "Connection: Keep -Alive"
    ElseIf ConnectionType = Closed Then
        sEndLine = "Connection: close"
    Else
        sEndLine = ""
    End If
    arr(2) = (topLineReassembled & vbCrLf & arr(2) & sEndLine)
    ExtractHttp = arr
End Function


