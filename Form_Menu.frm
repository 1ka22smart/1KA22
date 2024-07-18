VERSION 5.00
Begin VB.Form Form_Menu 
   Caption         =   "Form Menu"
   ClientHeight    =   8160
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12855
   LinkTopic       =   "Form3"
   ScaleHeight     =   8160
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   170
      Left            =   14880
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   7800
      TabIndex        =   0
      Top             =   2520
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "TUTUP"
         Height          =   615
         Left            =   3120
         TabIndex        =   4
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Pembuat : Rasyid,Fikri,Rafli,Robby,Yeftal"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   600
         TabIndex        =   3
         Top             =   3000
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Versi Apikasi: 1.5"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   2
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Nama Aplikasi: JORDUNK"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.Image Image1 
      Height          =   4200
      Left            =   8040
      Picture         =   "Form_Menu.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   5025
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Tempatnya Beli Sepatu Kw Kualitas Premium "
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "JORDUNK STORE "
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      TabIndex        =   5
      Top             =   360
      Width           =   6015
   End
   Begin VB.Image Image2 
      Height          =   10650
      Left            =   0
      Picture         =   "Form_Menu.frx":178D22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20520
   End
   Begin VB.Menu minput 
      Caption         =   "MENU INPUT"
      Begin VB.Menu msepatu 
         Caption         =   "INPUT SEPATU"
      End
      Begin VB.Menu mtransaksi 
         Caption         =   "INPUT TRANSAKSI"
      End
   End
   Begin VB.Menu mtentang 
      Caption         =   "TENTANG APLIKASI"
   End
   Begin VB.Menu mkeluar 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = False
End Sub

Private Sub Form_Load()
Frame1.Visible = False
End Sub

Private Sub mkeluar_Click()
If MsgBox("Keluar???", vbYesNo + vbQuestion, "Peringatan") = vbYes Then
Form_Sepatu.Visible = False
End
End If
End Sub

Private Sub msepatu_Click()
Form_Sepatu.Show
End Sub

Private Sub mtentang_Click()
Frame1.Visible = True
End Sub

Private Sub mtransaksi_Click()
Form_Transaksi.Show
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Right(Label5.Caption, Len(Label5.Caption) - 1) & Left(Label5.Caption, 1)



End Sub
