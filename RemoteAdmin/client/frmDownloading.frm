VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDownloading 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Downloading File"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5340
   Icon            =   "frmDownloading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5220
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblBytes 
         Height          =   285
         Left            =   1065
         TabIndex        =   4
         Top             =   825
         Width           =   3225
      End
      Begin VB.Label Label1 
         Caption         =   "Filename:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Bytes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   255
         TabIndex        =   2
         Top             =   795
         Width           =   1305
      End
      Begin VB.Label lblFIleName 
         Height          =   285
         Left            =   1515
         TabIndex        =   1
         Top             =   360
         Width           =   3225
      End
   End
End
Attribute VB_Name = "frmDownloading"
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

Option Explicit

Private Sub Form_Load()
    On Error GoTo Err
        Me.Refresh
        DoEvents

Form_Load_Exit:
        Exit Sub
    
Err:
        MsgBox Err.Description, vbCritical, "Remote File Explorer!"
        Exit Sub
End Sub
