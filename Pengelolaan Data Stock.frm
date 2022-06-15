VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PengelolaanDataStock 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Form Inputan"
      TabPicture(0)   =   "Pengelolaan Data Stock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Data1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Tampilan Data"
      TabPicture(1)   =   "Pengelolaan Data Stock.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).ControlCount=   2
      Begin VB.TextBox Text3 
         DataField       =   "UNITSTOCK"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   15
         Top             =   3600
         Width           =   7215
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Pengelolaan Data Stock.frx":0038
         Height          =   4095
         Left            =   -73680
         OleObjectBlob   =   "Pengelolaan Data Stock.frx":004C
         TabIndex        =   1
         Top             =   2160
         Width           =   8655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Selesai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7920
         TabIndex        =   14
         Top             =   6120
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5400
         TabIndex        =   13
         Top             =   6000
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3000
         TabIndex        =   12
         Top             =   6000
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   11
         Top             =   6000
         Width           =   2055
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 17\Database.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TSTOCK"
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         DataField       =   "HARGASTOCK"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   10
         Top             =   4320
         Width           =   7215
      End
      Begin VB.TextBox Text2 
         DataField       =   "NAMASTOCK"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   9
         Top             =   2880
         Width           =   7215
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOSTOCK"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   8
         Top             =   2160
         Width           =   7215
      End
      Begin VB.Label Label6 
         Caption         =   "Harga Stock"
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Unit Stock"
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Nama Stock"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Nomor Stock"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PENGELOLAAN DATA STOCK"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   10695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PENGELOLAAN DATA CUSTOMER"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   2
         Top             =   960
         Width           =   10695
      End
   End
End
Attribute VB_Name = "PengelolaanDataStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Add data
    If Text1 <> "" Then
        Data1.Recordset.AddNew
    End If
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    'Search Data
    Respon = vbYes
    While Respon = vbYes
        Respon = InputBox("Masukkan Nomor Customer", "Pencarian Data")
        Cari = "NOSTOCK='" + Respon + "'"
        Data1.Recordset.FindFirst Cari
        If Data1.Recordset.NoMatch Then
            Respon = MsgBox("Data yang dicari tidak ditemukan, cari lainnya?", vbYesNo, "Cari Data")
        Else
            Respon = vbNo
        End If
    Wend
End Sub

Private Sub Command3_Click()
    'Delete data
    Respon = MsgBox("Menghapus data?", vbYesNo, "Hapus Data")
    If Respon = vbYes Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
    Text1.SetFocus
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Form_Activate()
    Data1.Recordset.MoveLast
    If Data1.Recordset!NOSTOCK <> "" Then
        Data1.Recordset.AddNew
    End If
End Sub

Private Sub Form_Load()
    'Fullscreen
    PengelolaanDataStock.WindowState = 2
End Sub
