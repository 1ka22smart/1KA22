VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Transaksi 
   Caption         =   "HALAMAN DATA TRANSAKSI"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form_Transaksi.frx":0000
      Height          =   3255
      Left            =   14640
      TabIndex        =   31
      Top             =   2760
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   14760
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11520
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select * from tbl_transaksi"
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
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   14520
      TabIndex        =   28
      Top             =   7200
      Width           =   5175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form_Transaksi.frx":0015
         Height          =   2415
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4260
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "id_pembayaran"
            Caption         =   "id_pembayaran"
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
            DataField       =   "tanggal"
            Caption         =   "tanggal"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "jumlah"
            Caption         =   "jumlah"
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
         BeginProperty Column06 
            DataField       =   "total"
            Caption         =   "total"
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "bayar"
            Caption         =   "bayar"
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
         BeginProperty Column09 
            DataField       =   "kembali"
            Caption         =   "kembali"
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
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   5280
      TabIndex        =   21
      Top             =   8880
      Width           =   8655
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "EDIT"
         Height          =   495
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   4680
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "BATAL"
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdcetak 
         Caption         =   "CETAK"
         Height          =   495
         Left            =   2640
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdkeluar 
         Caption         =   "KELUAR"
         Height          =   495
         Left            =   4680
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
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
      Height          =   6975
      Left            =   5160
      TabIndex        =   1
      Top             =   1560
      Width           =   8655
      Begin VB.TextBox tjumlahbeli 
         Height          =   405
         Left            =   4080
         TabIndex        =   33
         Top             =   3480
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dttanggal 
         Height          =   495
         Left            =   4080
         TabIndex        =   32
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         Format          =   118161409
         CurrentDate     =   45485
      End
      Begin VB.CommandButton cmdidsepatu 
         Caption         =   "..."
         Height          =   375
         Left            =   6960
         TabIndex        =   30
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox tkembali 
         Height          =   405
         Left            =   4080
         TabIndex        =   20
         Top             =   5880
         Width           =   3495
      End
      Begin VB.TextBox tbayar 
         Height          =   375
         Left            =   4080
         TabIndex        =   19
         Top             =   5280
         Width           =   3495
      End
      Begin VB.TextBox tterbilang 
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox ttotal 
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox thargasepatu 
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox tidpembayaran 
         Height          =   405
         Left            =   4080
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox tidsepatu 
         Height          =   405
         Left            =   4080
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox tmerksepatu 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   2280
         Width           =   3495
      End
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
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "KEMBALI"
         BeginProperty Font 
            Name            =   "Berlin Sans FB Demi"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "BAYAR"
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
         Left            =   480
         TabIndex        =   14
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "TOTAL TERBILANG"
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
         Left            =   480
         TabIndex        =   13
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "TOTAL"
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
         Left            =   480
         TabIndex        =   12
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "JUMLAH BELI"
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
         Left            =   480
         TabIndex        =   11
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Left            =   480
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "ID PEMBAYARAN"
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
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "TANGGAL"
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
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
   End
   Begin VB.Image Image2 
      Height          =   2760
      Left            =   960
      Picture         =   "Form_Transaksi.frx":002A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2625
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
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   12495
   End
   Begin VB.Image Image1 
      Height          =   10890
      Left            =   0
      Picture         =   "Form_Transaksi.frx":18CFB1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20280
   End
End
Attribute VB_Name = "Form_Transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdbatal_Click()
tidpembayaran.Text = ""
tidsepatu.Text = ""
tmerksepatu.Text = ""
thargasepatu.Text = ""
tjumlahbeli.Text = ""
ttotal.Text = ""
tterbilang.Text = ""
tbayar.Text = ""
tkembali.Text = ""
End Sub

Private Sub cmdcetak_Click()
Set DataReportTransaksi.DataSource = Adodc1
Adodc1.RecordSource = "select * from tbl_transaksi where id_pembayaran = ' " & tidpembayaran & " '"
DataReportTransaksi.Refresh
DataReportTransaksi.Show
End Sub

Private Sub cmdedit_Click()
Adodc1.Recordset.UpdateBatch
Adodc1.Recordset.Fields(0) = tidpembayaran.Text
Adodc1.Recordset.Fields(1) = dttanggal.Value
Adodc1.Recordset.Fields(2) = tidsepatu.Text
Adodc1.Recordset.Fields(3) = tmerksepatu.Text
Adodc1.Recordset.Fields(4) = thargasepatu.Text
Adodc1.Recordset.Fields(5) = tjumlahbeli.Text
Adodc1.Recordset.Fields(6) = ttotal.Text
Adodc1.Recordset.Fields(7) = tterbilang.Text
Adodc1.Recordset.Fields(8) = tbayar.Text
Adodc1.Recordset.Fields(9) = tkembali.Text
Adodc1.Recordset.Update

tidpembayaran.Text = ""
tidsepatu.Text = ""
tmerksepatu.Text = ""
thargasepatu.Text = ""
tjumlahbeli.Text = ""
ttotal.Text = ""
tterbilang.Text = ""
tbayar.Text = ""
tkembali.Text = ""

Form_Load
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Hapus Data", vbYesNo + vbQuestion, "Peringatan") = vbYes Then
Adodc1.Recordset.Delete
End If
tmerksepatu.SetFocus
tidsepatu.Enabled = False

tidpembayaran.Text = ""
tidsepatu.Text = ""
tmerksepatu.Text = ""
thargasepatu.Text = ""
tjumlahbeli.Text = ""
ttotal.Text = ""
tterbilang.Text = ""
tbayar.Text = ""
tkembali.Text = ""

Form_Load
End Sub

Private Sub cmdidsepatu_Click()
DataGrid2.Visible = True
Adodc2.RecordSource = "select * from tbl_sepatu"
Adodc2.Refresh
End Sub

Private Sub cmdkeluar_Click()
If MsgBox("Keluar???", vbYesNo + vbQuestion, "Peringatan") = vbYes Then
Form_Transaksi.Visible = False
End If
End Sub

Private Sub cmdsimpan_Click()
    If tidpembayaran.Text = "" Then
        MsgBox ("Id Pembayaran tidak boleh kosong, Harap Tekan tombol TAMBAH"), vbCritical, "Informasi"
    Else

    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields(0) = tidpembayaran.Text
    Adodc1.Recordset.Fields(1) = dttanggal.Value
    Adodc1.Recordset.Fields(2) = tidsepatu.Text
    Adodc1.Recordset.Fields(3) = tmerksepatu.Text
    Adodc1.Recordset.Fields(4) = thargasepatu.Text
    Adodc1.Recordset.Fields(5) = tjumlahbeli.Text
    Adodc1.Recordset.Fields(6) = ttotal.Text
    Adodc1.Recordset.Fields(7) = tterbilang.Text
    Adodc1.Recordset.Fields(8) = tbayar.Text
    Adodc1.Recordset.Fields(9) = tkembali.Text
    Adodc1.Recordset.Update

    tidpembayaran.Text = ""
    tidsepatu.Text = ""
    tmerksepatu.Text = ""
    thargasepatu.Text = ""
    tjumlahbeli.Text = ""
    ttotal.Text = ""
    tterbilang.Text = ""
    tbayar.Text = ""
    tkembali.Text = ""
    End If
End Sub


Private Sub cmdtambah_Click()
autonumberpembayaran
dttanggal = Format(Now, "DD/MM/YYYY")
tidsepatu.Text = ""
thargasepatu.Text = ""
tterbilang.Text = ""
tjumlahbeli.Text = ""
ttotal.Text = ""
tbayar.Text = ""
tkembali.Text = ""
Adodc1.RecordSource = "select * from tbl_transaksi"
Adodc1.Refresh
End Sub
Private Sub autonumberpembayaran()
Call buka
noidpembayaran.Open ("select * from tbl_transaksi where id_pembayaran in(select max(id_pembayaran) from tbl_transaksi) order by id_pembayaran desc"), conn
noidpembayaran.Requery

    Dim urut As String * 7
    Dim hitung As Long
    With noidpembayaran
        If .EOF Then
            urut = "TRS-" + "001"
            tidpembayaran = urut
        Else
            hitung = Right(!id_pembayaran, 3) + 1
            urut = Right("TRS-00" & hitung, 7)
        End If
        tidpembayaran = urut
    End With

End Sub

Private Sub DataGrid1_Click()
tidpembayaran.Text = Adodc1.Recordset!id_pembayaran
tidsepatu.Text = Adodc1.Recordset!id_sepatu
tmerksepatu.Text = Adodc1.Recordset!merk_sepatu
thargasepatu.Text = Adodc1.Recordset!harga_sepatu
tjumlahbeli.Text = Adodc1.Recordset!jumlah
ttotal.Text = Adodc1.Recordset!total
tterbilang.Text = Adodc1.Recordset!terbilang
tbayar.Text = Adodc1.Recordset!bayar
tkembali.Text = Adodc1.Recordset!kembali

cmdedit.Enabled = True
cmdhapus.Enabled = True
End Sub

Private Sub DataGrid2_Click()
tidsepatu.Text = Adodc2.Recordset!id_sepatu
tmerksepatu.Text = Adodc2.Recordset!merk_sepatu
thargasepatu.Text = Adodc2.Recordset!harga_sepatu
End Sub

Private Sub DataGrid2_LostFocus()
DataGrid2.Visible = False
End Sub

Private Sub Form_Load()
tidpembayaran.Enabled = False
DataGrid2.Visible = False

cmdedit.Enabled = False
cmdhapus.Enabled = False
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
Private Sub tbayar_Change()
tkembali.Text = Val(tbayar.Text) - Val(ttotal.Text)
End Sub


Private Sub tjumlahbeli_Change()
Dim angka As Double
Dim teks As String
ttotal.Text = Val(thargasepatu.Text) * Val(tjumlahbeli.Text)
angka = Val(ttotal.Text)
teks = Terbilangan(angka)
tterbilang.Text = teks + " rupiah"
End Sub
