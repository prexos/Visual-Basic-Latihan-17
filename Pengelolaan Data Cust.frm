VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PengelolaanDataCust 
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
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
      TabPicture(0)   =   "Pengelolaan Data Cust.frx":0000
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
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Data1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Tampilan Data"
      TabPicture(1)   =   "Pengelolaan Data Cust.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Pengelolaan Data Cust.frx":0038
         Height          =   3855
         Left            =   -74040
         OleObjectBlob   =   "Pengelolaan Data Cust.frx":004C
         TabIndex        =   19
         Top             =   2160
         Width           =   9375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "End"
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
         TabIndex        =   1
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         RecordSource    =   "TCUSTOMER"
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         DataField       =   "TELPONCUST"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   15
         Top             =   4440
         Width           =   7215
      End
      Begin VB.TextBox Text5 
         DataField       =   "KONTAKCUST"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   14
         Top             =   3960
         Width           =   7215
      End
      Begin VB.TextBox Text4 
         DataField       =   "KOTACUST"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   13
         Top             =   3480
         Width           =   7215
      End
      Begin VB.TextBox Text3 
         DataField       =   "ALAMATCUST"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   12
         Top             =   3000
         Width           =   7215
      End
      Begin VB.TextBox Text2 
         DataField       =   "NAMACUST"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   11
         Top             =   2520
         Width           =   7215
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOCUST"
         DataSource      =   "Data1"
         Height          =   390
         Left            =   2400
         TabIndex        =   10
         Top             =   2040
         Width           =   7215
      End
      Begin VB.Label Label8 
         Caption         =   "Telepon"
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Kontak"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Kota"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Alamat"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Nama"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nomor"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PENGELOLAAN DATA CUSTOMER"
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
Attribute VB_Name = "PengelolaanDataCust"
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
        Cari = "NOCUST='" + Respon + "'"
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
    MenuUtama.Show
End Sub

Private Sub Form_Activate()
    Data1.Recordset.MoveLast
    If Data1.Recordset!NOCUST <> "" Then
        Data1.Recordset.AddNew
    End If
End Sub

Private Sub Form_Load()
    'Fullscreen
    PengelolaanDataCust.WindowState = 2
End Sub
