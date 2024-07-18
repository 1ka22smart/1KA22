VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "ULANG"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BERSIHKAN"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   5520
      Top             =   7200
   End
   Begin VB.Timer Timer4 
      Interval        =   300
      Left            =   5040
      Top             =   7200
   End
   Begin VB.Timer Timer3 
      Left            =   4560
      Top             =   7200
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4080
      Top             =   7200
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3600
      Top             =   7200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000007&
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   3600
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000007&
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   3360
      TabIndex        =   11
      Top             =   4560
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000008&
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000008&
      ForeColor       =   &H00FFFFFF&
      Height          =   1650
      Left            =   600
      TabIndex        =   9
      Top             =   6840
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000007&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   600
      TabIndex        =   8
      Top             =   4560
      Width           =   2415
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   4080
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   661
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      AutoEnable      =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   555
      Left            =   600
      TabIndex        =   6
      Top             =   3360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   979
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "00 : 00 : 00"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dari"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Judul Lagu"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MP3 PLAYER"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      BorderWidth     =   4
      FillColor       =   &H00FFFF80&
      FillStyle       =   6  'Cross
      Height          =   8415
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim File As String
Dim Kode As Boolean
Dim Endtrack As Long
Dim Jam, menit, detik, mldetik As Integer

Sub Play()
'Pengaturan Waktu
mldetik = 0
detik = 0
menit = 0
Jam = 0

File = List2

If Mid(File, 3, 1) = "\" And Mid(File, 4, 1) = "\" Then
    File = List1
Else
    File = List2
End If

MMControl1.FileName = File
MMControl1.Command = "Open"
Endtrack = MMControl1.TrackLength

If Endtrack = 0 Then
    MsgBox "Tidak dapat memainkan mp3", vbOKOnly + vbCritical, "Player Error"
End If

End Sub

Private Sub Command1_Click()
List1.Clear
List2.Clear
End Sub

Private Sub Command2_Click()
If Command2.Caption = "ULANG" Then
Command2.Caption = "OFF"
Else
Command2.Caption = "ULANG"
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.FileName = "*.mp3;*.mid1"
End Sub

Private Sub Drive1_Change()
On Error GoTo perangkap
Dir1.Path = Drive1.Drive
perangkap:
    Select Case Err
        Case 68
            MsgBox "tidak dapat mengakses drive", vbOKOnly + vbCritical, "Scope Error"
            Drive1.Refresh
        Case 0
        Exit Sub
    End Select
End Sub

Private Sub File1_DblClick()
If File1.FileName = "" Then
Exit Sub
Else
List1.AddItem File1.FileName
List2.AddItem File1.Path & "\" & File1.FileName
Label3.Caption = List1.ListIndex + 1
Label5.Caption = List1.ListCount
End If
End Sub

Private Sub Form_Load()
Me.Left = 5000
Me.Top = Screen.Height
Timer3.Interval = 1
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
Label3.Caption = List1.ListIndex + 1
Label5.Caption = List1.ListCount
Label2.Caption = List1
MMControl1.Command = "Close"
MMControl1.Refresh
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
If Kode = True Then Exit Sub
If MMControl1.TrackLength = MMControl1.Position Then
If Label3.Caption = Label5.Caption Then

If Command2.Caption = "ULANG" Then
If Label5.Caption = "1" Then
MMControl1.Command = "Close"
Timer2.Enabled = False

Else
If Label3.Caption = Label5.Caption Then
List1.ListIndex = 0
MMControl1.Command = "Play"
End If
End If

Else
If Label3.Caption = Label5.Caption Then
MMControl1.Command = "Close"
End If
End If

Else
With List1
.ListIndex = .ListIndex + 1
End With
MMControl1.Command = "Play"
End If
End If
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
Play
ProgressOke
Label2.Caption = List1
Timer2.Enabled = True
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
MMControl1.Refresh
MMControl1.Command = "Close"
Kode = True
Timer2.Enabled = False

End Sub

Sub ProgressOke()
Slider1.Min = 0
Slider1.Max = Val(MMControl1.TrackLength + 1)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Slider1.Value = MMControl1.Position
End Sub

Private Sub Timer2_Timer()
If mldetik = 10 Then
detik = detik + 1
mldetik = 0
End If
If detik = 60 Then
menit = menit + 1
detik = 0
End If
If menit = 60 Then
Jam = Jam + 1
menit = 0
End If
Label6.Caption = Jam & ":" & menit & ":" & detik
mldetik = mldetik + 1
End Sub

Private Sub Timer3_Timer()
If Me.Top <= 1000 Then
Timer3.Interval = 0
Else
Me.Top = Me.Top - 100
End If
End Sub

Private Sub Timer4_Timer()
BackColor = RGB(Rnd() * 225, Rnd() * 225, Rnd() * 225)
End Sub

Private Sub Timer5_Timer()
Label2.Left = Label2.Left - 40
If Label2.Left <= -Label2.Left Then
Label2.Left = Label2.Width
End If
End Sub
