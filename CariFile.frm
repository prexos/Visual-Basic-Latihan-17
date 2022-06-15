VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CariFIle 
   Caption         =   "Form1"
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   3840
      TabIndex        =   0
      Top             =   1560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Existing"
      TabPicture(0)   =   "CariFile.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Dir1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Drive1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "File1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cls"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Recent (FILE FRM)"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "File2"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton Command2 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6480
         TabIndex        =   10
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cls 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6480
         TabIndex        =   9
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   6720
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   6120
         Width           =   3375
      End
      Begin VB.FileListBox File2 
         Height          =   4680
         Left            =   -73680
         Pattern         =   "*.frm"
         TabIndex        =   4
         Top             =   1560
         Width           =   6015
      End
      Begin VB.FileListBox File1 
         Height          =   2640
         Left            =   1080
         TabIndex        =   3
         Top             =   3120
         Width           =   5055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.DirListBox Dir1 
         Height          =   1230
         Left            =   1080
         TabIndex        =   1
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Tipe File"
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nama File [Tanpa Extensi]"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   6000
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "HANYA MENAMPILKAN NAMA-NAMA FILE, TIDAK ADA PROSES LANJUTAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   8895
   End
End
Attribute VB_Name = "CariFIle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cls_Click()
    End
End Sub

Private Sub Combo1_Click()
    File1.Pattern = "*." + Left(Combo1.Text, 3)
End Sub

Private Sub Command2_Click()
    Pesan = MsgBox("PROSES SELANJUTNYA ATAS FILE YANG DIPILIH, SEMENTARA INI TIDAK ADA.", vbInformation, "Perhatian")
End Sub

Private Sub Dir1_Change()
    'Koneksi Directory dengan File
        File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    'Koneksi Drive dengan Directory
        Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Pesan = MsgBox("PROSES SELANJUTNYA ATAS FILE YANG DIPILIH, SEMENTARA INI TIDAK ADA.", vbInformation, "Perhatian")
End Sub

Private Sub Form_Activate()
    'Fullscreen
        CariFIle.WindowState = 2
    'Combo Box 1
        Combo1.AddItem "JPG - File Gambar"
        Combo1.AddItem "GIF - File Gambar Bergerak"
        Combo1.AddItem "Doc - Ms. Word File"
        Combo1.AddItem "XLS - Ms. Excel File"
        Combo1.AddItem "PDF - Portable Data File"
        Combo1.AddItem "* - Seluruh Jenis File"
End Sub

Private Sub Text1_LostFocus()
    File1.Pattern = Text1.Text + "*.*"
End Sub
