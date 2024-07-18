VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Sepatu 
   Caption         =   "HALAMAN DATA SEPATU"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   3015
      Left            =   3960
      TabIndex        =   19
      Top             =   5520
      Width           =   13815
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form_Sepatu.frx":0000
         Height          =   1575
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   2778
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "id_sepatu"
            Caption         =   "id_sepatu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "merk_sepatu"
            Caption         =   "merk_sepatu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "harga_sepatu"
            Caption         =   "harga_sepatu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "terbilang"
            Caption         =   "terbilang"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   14345
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3480
      TabIndex        =   12
      Top             =   4560
      Width           =   14775
      Begin VB.CommandButton cmdkeluar 
         Caption         =   "KELUAR"
         Height          =   495
         Left            =   11280
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdcetak 
         Caption         =   "CETAK"
         Height          =   495
         Left            =   9120
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "BATAL"
         Height          =   495
         Left            =   6720
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "EDIT"
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "INPUT DATA SEPATU"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Width           =   9975
      Begin VB.CommandButton cmdtambah 
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox tterbilang 
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox thargasepatu 
         Height          =   405
         Left            =   4080
         TabIndex        =   9
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox tmerksepatu 
         Height          =   405
         Left            =   4080
         TabIndex        =   8
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox tidsepatu 
         Height          =   405
         Left            =   4080
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "TERBILANG"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "HARGA SEPATU"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "MERK SEPATU"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "ID SEPATU"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3960
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   9480
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\PENJUALAN SEPATU\db_penjualan.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\PENJUALAN SEPATU\db_penjualan.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from tbl_sepatu"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "SELAMAT DATANG DI TOKO SEPATU KAMI"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   11895
      End
   End
   Begin VB.Image Image3 
      Height          =   2880
      Left            =   840
      Picture         =   "Form_Sepatu.frx":0015
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   11010
      Left            =   -120
      Picture         =   "Form_Sepatu.frx":18CF9C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20280
   End
   Begin VB.Image Image2 
      Height          =   46920
      Left            =   2040
      Picture         =   "Form_Sepatu.frx":359E30
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   40065
   End
End
Attribute VB_Name = "Form_Sepatu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbatal_Click()
tidsepatu.Text = ""
tmerksepatu.Text = ""
thargasepatu.Text = ""
tterbilang.Text = ""
End Sub

Private Sub cmdcetak_Click()
Set DataReportSepatu.DataSource = Adodc1
Adodc1.RecordSource = "select * from tbl_sepatu where id_sepatu = ' " & tidsepatu & " '"
DataReportSepatu.Refresh
DataReportSepatu.Show
End Sub

Private Sub cmdedit_Click()
Adodc1.Recordset.UpdateBatch
Adodc1.Recordset.Fields(0) = tidsepatu.Text
Adodc1.Recordset.Fields(1) = tmerksepatu.Text
Adodc1.Recordset.Fields(2) = thargasepatu.Text
Adodc1.Recordset.Fields(3) = tterbilang.Text
Adodc1.Recordset.Update
tmerksepatu.SetFocus
tidsepatu.Enabled = False

tidsepatu.Text = ""
tmerksepatu.Text = ""
thargasepatu.Text = ""
tterbilang.Text = ""

Form_Load
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Hapus Data", vbYesNo + vbQuestion, "Peringatan") = vbYes Then
Adodc1.Recordset.Delete
End If
tmerksepatu.SetFocus
tidsepatu.Enabled = False

tidsepatu.Text = ""
tmerksepatu.Text = ""
thargasepatu.Text = ""
tterbilang.Text = ""

Form_Load
End Sub

Private Sub cmdkeluar_Click()
If MsgBox("Keluar???", vbYesNo + vbQuestion, "Peringatan") = vbYes Then
Form_Sepatu.Visible = False
End If
End Sub

Private Sub cmdsimpan_Click()

    If tidsepatu.Text = "" Then
        MsgBox ("Id Sepatu tidak boleh kosong, Harap Tekan tombol TAMBAH"), vbCritical, "Informasi"
    Else
        If tmerksepatu.Text = "" Then
            MsgBox ("Merk Sepatu tidak boleh kosong"), vbCritical, "Informasi"
            tmerksepatu.SetFocus
        Else
            If thargasepatu.Text = "" Then
                MsgBox ("Harga Sepatu tidak boleh Kosong"), vbCritical, "Informasi"
                thargasepatu.SetFocus
            Else
                Adodc1.Recordset.AddNew
                Adodc1.Recordset.Fields(0) = tidsepatu.Text
                Adodc1.Recordset.Fields(1) = tmerksepatu.Text
                Adodc1.Recordset.Fields(2) = thargasepatu.Text
                Adodc1.Recordset.Fields(3) = tterbilang.Text
                Adodc1.Recordset.Update
                tmerksepatu.SetFocus
                tidsepatu.Text = ""
                tmerksepatu.Text = ""
                thargasepatu.Text = ""
                tterbilang.Text = ""
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
tidsepatu.Enabled = False

cmdedit.Enabled = False
cmdhapus.Enabled = False
End Sub

Private Sub thargasepatu_Change()
    Dim angka As Double
    Dim teks As String
    angka = Val(thargasepatu.Text)
    teks = Terbilangan(angka)
    tterbilang.Text = teks + " rupiah"
End Sub
Public Function Terbilangan(X As Double) As String ' (diketik manual)
Dim tampung As Double
Dim teks As String
Dim bagian As String
Dim i As Integer
Dim tanda As Boolean
Dim letak(5)
letak(1) = "ribu "
letak(2) = "juta "
letak(3) = "milyar "
letak(4) = "trilyun "
If (X = 0) Then
Terbilangan = "nol"
Exit Function
End If
If (X < 2000) Then
tanda = True
End If
teks = ""
If (X >= 1E+15) Then
Terbilangan = "Nilai terlalu besar"
Exit Function
End If
For i = 4 To 1 Step -1
tampung = Int(X / (10 ^ (3 * i)))
If (tampung > 0) Then
bagian = ratusan(tampung, tanda)
teks = teks & bagian & letak(i)
End If
X = X - tampung * (10 ^ (3 * i))
Next
teks = teks & ratusan(X, False)
Terbilangan = teks
End Function
Function ratusan(ByVal Y As Double, ByVal flag As Boolean) As String ' (diketik manual)
Dim tmp As Double
Dim bilang As String
Dim bag As String
Dim j As Integer
Dim angka(9)
angka(1) = "se"
angka(2) = "dua "
angka(3) = "tiga "
angka(4) = "empat "
angka(5) = "lima "
angka(6) = "enam "
angka(7) = "tujuh "
angka(8) = "delapan "
angka(9) = "sembilan "
Dim posisi(2)
posisi(1) = "puluh "
posisi(2) = "ratus "
bilang = ""
For j = 2 To 1 Step -1
tmp = Int(Y / (10 ^ j))
If (tmp > 0) Then
bag = angka(tmp)
If (j = 1 And tmp = 1) Then
Y = Y - tmp * 10 ^ j
If (Y >= 1) Then
posisi(j) = "belas "
Else
angka(Y) = "se"
End If
bilang = bilang & angka(Y) & posisi(j)
ratusan = bilang
Exit Function
Else
bilang = bilang & bag & posisi(j)
End If
End If
Y = Y - tmp * 10 ^ j
Next
If (flag = False) Then
angka(1) = "satu "
End If
bilang = bilang & angka(Y)
ratusan = bilang
End Function
Private Sub cmdtambah_Click()
autonumbersepatu
tmerksepatu.Text = ""
thargasepatu.Text = ""
tterbilang.Text = ""
Adodc1.RecordSource = "select * from tbl_sepatu"
Adodc1.Refresh
End Sub
Private Sub DataGrid1_Click()
    tidsepatu.Text = Adodc1.Recordset!id_sepatu
    tmerksepatu.Text = Adodc1.Recordset!merk_sepatu
    thargasepatu.Text = Adodc1.Recordset!harga_sepatu
    tterbilang.Text = Adodc1.Recordset!terbilang
    
    cmdedit.Enabled = True
    cmdhapus.Enabled = True
End Sub

Private Sub autonumbersepatu()
Call buka
noidsepatu.Open ("select * from tbl_sepatu where id_sepatu in(select max(id_sepatu) from tbl_sepatu) order by id_sepatu desc"), conn
noidsepatu.Requery

    Dim urut As String * 6
    Dim hitung As Long
    With noidsepatu
        If .EOF Then
            urut = "SP-" + "001"
            tidsepatu = urut
        Else
            hitung = Right(!id_sepatu, 3) + 1
            urut = Right("SP-00" & hitung, 6)
        End If
        tidsepatu = urut
    End With

End Sub


