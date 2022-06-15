VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu PemDa 
      Caption         =   "Pemeliharaan Data"
      Begin VB.Menu PemDaCust 
         Caption         =   "Pemeliharaan Data Customer"
      End
      Begin VB.Menu PemDaStock 
         Caption         =   "Pemeliharaan Data Stock"
      End
   End
   Begin VB.Menu trx 
      Caption         =   "Transaksi"
   End
   Begin VB.Menu Lap 
      Caption         =   "Laporan"
      Begin VB.Menu LapDataCustomer 
         Caption         =   "Laporan Data Customer"
      End
      Begin VB.Menu LapDataPenju 
         Caption         =   "Laporan Data Penjualan"
      End
      Begin VB.Menu LapDataStock 
         Caption         =   "Laporan Data Stock"
      End
   End
   Begin VB.Menu Bantuan 
      Caption         =   "Bantuan"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bantuan_Click()
    CariFIle.Show
End Sub

Private Sub LapDataCustomer_Click()
    LapDataCust.Show
End Sub

Private Sub LapDataPenju_Click()
    LapDataPenjualan.Show
End Sub

Private Sub LapDataStock_Click()
    LapDataStok.Show
End Sub

Private Sub MDIForm_Load()
    MDIForm1.WindowState = 2
End Sub

Private Sub PemDaCust_Click()
    PengelolaanDataCust.Show
End Sub

Private Sub PemDaStock_Click()
    PengelolaanDataStock.Show
End Sub

Private Sub trx_Click()
    Transaksi.Show
End Sub
