VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4140
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Height          =   3375
      Left            =   8880
      TabIndex        =   1
      Top             =   360
      Width           =   3255
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdopen 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdstop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdpause 
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdplay 
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      Begin VB.Label lnama 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   2640
         Width           =   8175
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmp1 
         Height          =   3375
         Left            =   -120
         TabIndex        =   7
         Top             =   0
         Width           =   8535
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   15055
         _cy             =   5953
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdexit_Click()
If MsgBox("apakah ingin keluar", vbYesNo + vbQuestion, "warning") = vbYes Then
End
End If
End Sub

Private Sub cmdopen_Click()
cd1.DefaultExt = "*.mp3"
cd1.Filter = "audio file (*.mp3)|*.mp3| Wav file (*.wav)|*.wav | all file(*.*)|*.*"
cd1.ShowOpen
If cd1.FileName <> "" Then
wmp1.Controls.stop
wmp1.URL = cd1.FileName
wmp1.settings.autoStart = True

lnama.Caption = wmp1.Controls.currentItem.Name
End If
End Sub

Private Sub cmdpause_Click()
wmp1.Controls.pause
End Sub

Private Sub cmdplay_Click()
wmp1.Controls.play
End Sub

Private Sub cmdstop_Click()
wmp1.Controls.stop
End Sub

