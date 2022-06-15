VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Transaksi 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8640
      TabIndex        =   30
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118358017
      CurrentDate     =   44727
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan dan Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10680
      TabIndex        =   29
      Top             =   8760
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   11880
      TabIndex        =   16
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   11880
      TabIndex        =   15
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   11880
      TabIndex        =   14
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txtunit 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   9600
      TabIndex        =   13
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   3840
      TabIndex        =   11
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtnostok 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   1680
      TabIndex        =   10
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identitas Customer"
      Height          =   2655
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   5295
      Begin VB.TextBox Text4 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtnocust 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   360
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton selesai 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      TabIndex        =   1
      Top             =   8760
      Width           =   2535
   End
   Begin VB.Data DataJual 
      Caption         =   "Jualan"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 17\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TJUAL"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Data DataCust 
      Caption         =   "Customer"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 17\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TCUSTOMER"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Data DataStock 
      Caption         =   "Stock"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 17\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TSTOCK"
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton savenoexit 
      Caption         =   "Simpan dan Isi lagi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   0
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Label Label14 
      Caption         =   "Nilai Penjualan Bersih"
      Height          =   255
      Left            =   9480
      TabIndex        =   28
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "Besaran Potongan"
      Height          =   255
      Left            =   9720
      TabIndex        =   27
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Nilai Jual"
      Height          =   255
      Left            =   12480
      TabIndex        =   26
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Unit Jual"
      Height          =   255
      Left            =   10080
      TabIndex        =   25
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Harga Jual"
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Nama Stok"
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Nomor Stok"
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1560
      X2              =   14880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lbltgl 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   21
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Faktur"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "FAKTUR PENJUALAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   19
      Top             =   840
      Width           =   2535
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1560
      X2              =   14880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "CV BITFINEX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label15 
      Caption         =   "Tanggal Penjualan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   17
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    'Add record baru pada Table Penjualan
    DataJual.Recordset.AddNew
    DataJual.Recordset!NOFAKTUR = Text1.Text
    DataJual.Recordset!TGLTRANS = DTPicker1.Value
    DataJual.Recordset!HARGAJUAL = Text8.Text
    DataJual.Recordset!NOCUST = txtnocust.Text
    DataJual.Recordset!NOSTOK = txtnostok.Text
    DataJual.Recordset!UNITJUAL = txtunit.Text
    DataJual.Recordset!POTONGAN = Text11.Text
    DataJual.Recordset.Update
    
    'Edit Customer Saldo Hutang
    DataCust.Recordset.Edit
    DataCust.Recordset!SALDOHUTANG = DataCust.Recordset!SALDOHUTANG + Val(Text12.Text)
    DataCust.Recordset.Update
    
    'Edit Stock Unit stock
    DataStock.Recordset.Edit
    DataStock.Recordset!UNITSTOCK = DataStock.Recordset!UNITSTOCK - Val(txtunit.Text)
    DataStock.Recordset.Update
    
    End
End Sub

Private Sub savenoexit_Click()
    'Add record baru pada Table Penjualan
    DataJual.Recordset.AddNew
    DataJual.Recordset!NOFAKT = Text1.Text
    DataJual.Recordset!TGLTRANS = DTPicker1.Value
    DataJual.Recordset!HARGAJUAL = Text8.Text
    DataJual.Recordset!NOCUST = txtnocust.Text
    DataJual.Recordset!NOSTOK = txtnostok.Text
    DataJual.Recordset!UNITJUAL = txtunit.Text
    DataJual.Recordset!POTONGAN = Text11.Text
    DataJual.Recordset.Update
    
    'Edit Customer Saldo Hutang
    DataCust.Recordset.Edit
    DataCust.Recordset!SALDOHUTANG = DataCust.Recordset!SALDOHUTANG + Val(Text12.Text)
    DataCust.Recordset.Update
    
    'Edit Stock Unit stock
    DataStock.Recordset.Edit
    DataStock.Recordset!UNITSTOCK = DataStock.Recordset!UNITSTOCK - Val(txtunit.Text)
    DataStock.Recordset.Update
    
    'Make the textbox blank
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    txtunit.Text = ""
    txtnocust.Text = ""
    txtnostok.Text = ""
    Text3.Enabled = True
    Text4.Enabled = True
    Text7.Enabled = True
    
    Text1.SetFocus
End Sub

Private Sub selesai_Click()
    End
End Sub

Private Sub DTPicker1_Click()
    lbltgl.Caption = DTPicker1.Value
End Sub

Private Sub Form_Activate()
    'Focus nomor faktur
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Transaksi.WindowState = 2
End Sub

Private Sub Text11_LostFocus()
    Text12.Text = Val(Text10.Text) - Val(Text11.Text)
    selesai.SetFocus
End Sub

Private Sub Text8_LostFocus()
    If txtunit.Text = "" Then
        txtunit.SetFocus
    Else
        Text10.Text = Val(Text8.Text) * Val(txtunit.Text)
    End If
End Sub

Private Sub txtnocust_LostFocus()
    Cari = "NOCUST = '" + txtnocust.Text + "'"
    DataCust.Recordset.FindFirst Cari
    If DataCust.Recordset.NoMatch Then
        If Respon = vbYes Then
            txtnocust.Text = ""
            txtnocust.SetFocus
        Else
            selesai.SetFocus
        End If
    Else
        'Fill the textbox
        Text3.Text = DataCust.Recordset!NAMACUST
        Text4.Text = DataCust.Recordset!ALAMATCUST
        'Disabled textbox
        Text3.Enabled = False
        Text4.Enabled = False
        'Focus
        txtnostok.SetFocus
    End If
End Sub

Private Sub txtnostok_LostFocus()
    Cari = "NOSTOCK = '" + txtnostok.Text + "'"
    DataStock.Recordset.FindFirst Cari
    If DataStock.Recordset.NoMatch Then
        Respon = MsgBox("Data Tidak ditemukan! Cari lainnya?", vbYesNo, "Cari Data")
        If Respon = vbYes Then
            txtnostok.Text = ""
            txtnostok.SetFocus
        Else
            savenoexit.SetFocus
        End If
    Else
        Text7.Text = DataStock.Recordset!NAMASTOCK
        Text7.Enabled = False
        
        Text8.SetFocus
    End If
End Sub

Private Sub txtunit_LostFocus()
    If Text8.Text = "" Then
        Text8.SetFocus
    Else
        Text10.Text = Val(Text8.Text) * Val(txtunit.Text)
        Text11.SetFocus
    End If
End Sub
