VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.MDIForm UTAMA 
   BackColor       =   &H80000004&
   Caption         =   "DISPRO"
   ClientHeight    =   7680
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15495
   Icon            =   "UTAMA.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "UTAMA.frx":5C12
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6885
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":1C867
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":230C9
            Key             =   "dc"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":2992B
            Key             =   "keluar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":3018D
            Key             =   "masuk"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":369EF
            Key             =   "piutang"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":3D251
            Key             =   "supplier"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":43AB3
            Key             =   "barang"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":47D5A
            Key             =   "gudang"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":AE0A8
            Key             =   "rpt1"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":B22ED
            Key             =   "rpt2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":B60A0
            Key             =   "rpt3"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":B9FEE
            Key             =   "rpt4"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":BE134
            Key             =   "rpt5"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":C1E89
            Key             =   "rpt6"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":C6114
            Key             =   "rpt7"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":CA0AB
            Key             =   "next"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":D090D
            Key             =   "grafik"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UTAMA.frx":D131F
            Key             =   "upload"
         EndProperty
      EndProperty
   End
   Begin cPopMenu6.PopMenu PopMenu1 
      Left            =   5310
      Top             =   2205
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HighlightStyle  =   2
      ActiveMenuForeColor=   65535
      MenuBackgroundColor=   14737632
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7305
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   16563
            Picture         =   "UTAMA.frx":D36F2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "UTAMA.frx":D9F54
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "UTAMA.frx":1402A2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "UTAMA.frx":1A65F0
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "UTAMA.frx":1ACE52
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   15465
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      Begin VB.Timer TimerDClose 
         Left            =   10215
         Top             =   990
      End
      Begin VB.Timer TimerExit 
         Interval        =   1000
         Left            =   8280
         Top             =   2655
      End
      Begin VB.Timer TimerKoneksi 
         Left            =   7155
         Top             =   765
      End
      Begin vbalIml6.vbalImageList vbalImageList1 
         Left            =   4905
         Top             =   3690
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   24
         IconSizeY       =   24
         Size            =   4920
         Images          =   "UTAMA.frx":1B36B4
         Version         =   131072
         KeyCount        =   2
         Keys            =   "ÿ"
      End
      Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbalExplorerBarCtl1 
         Height          =   6135
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   10821
         BackColorEnd    =   0
         BackColorStart  =   0
      End
      Begin VB.Label lblM_Master 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   7425
         TabIndex        =   10
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblClose_P 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   7065
         TabIndex        =   9
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblip 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   6345
         TabIndex        =   8
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblnmcom 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   6705
         TabIndex        =   7
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblms_office 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   5985
         TabIndex        =   6
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblkduser 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   5625
         TabIndex        =   4
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lbltglOD 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   5130
         TabIndex        =   3
         Top             =   45
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   240
         Left            =   4770
         TabIndex        =   1
         Top             =   45
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&System"
      Index           =   0
      Begin VB.Menu mnS1 
         Caption         =   "&1. Daily Closing"
         Index           =   0
      End
      Begin VB.Menu mnS1 
         Caption         =   "&2. Upload Customer IAP"
         Index           =   1
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Master"
      Index           =   1
      Begin VB.Menu mnMaster 
         Caption         =   "&1. Barang"
         Index           =   0
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&2. Gudang"
         Index           =   1
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&3. Supplier"
         Index           =   2
      End
      Begin VB.Menu mnMaster 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&4. Customer"
         Index           =   4
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&5. Supporting"
         Index           =   5
         Begin VB.Menu mnM5A1 
            Caption         =   "&1. Teknisi dan Cheker"
            Index           =   0
         End
         Begin VB.Menu mnM5A1 
            Caption         =   "&2. Kolektor"
            Index           =   1
         End
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&6. Stok Point IAP"
         Index           =   6
      End
      Begin VB.Menu mnMaster 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&7. Area Penagihan"
         Index           =   8
      End
      Begin VB.Menu mnMaster 
         Caption         =   "&8. Area Cekher"
         Index           =   9
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Transaksi"
      Index           =   2
      Begin VB.Menu mnTA1 
         Caption         =   "&1. PO Barang"
         Index           =   0
         Begin VB.Menu mnT1A1 
            Caption         =   "&1. PO Pembelian Barang"
            Index           =   0
         End
         Begin VB.Menu mnT1A1 
            Caption         =   "&2. PO Permintaan Barang"
            Index           =   1
         End
      End
      Begin VB.Menu mnTA1 
         Caption         =   "&2. Penerimaan Barang"
         Index           =   1
         Begin VB.Menu mnT2A1 
            Caption         =   "&1. Pembelian Barang"
            Index           =   0
         End
         Begin VB.Menu mnT2A1 
            Caption         =   "&2. Retur Pinjaman"
            Index           =   1
         End
         Begin VB.Menu mnT2A1 
            Caption         =   "&3. Retur Sewa"
            Index           =   2
         End
      End
      Begin VB.Menu mnTA1 
         Caption         =   "&3. Pengeluaran Barang"
         Index           =   2
         Begin VB.Menu mnT3A1 
            Caption         =   "&1. Free"
            Index           =   0
         End
         Begin VB.Menu mnT3A1 
            Caption         =   "&2. Pinjam Pakai"
            Index           =   1
         End
         Begin VB.Menu mnT3A1 
            Caption         =   "&3. Sewa"
            Index           =   2
         End
         Begin VB.Menu mnT3A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnT3A1 
            Caption         =   "&4. Perbaikan/Mutasi Barang"
            Index           =   4
         End
         Begin VB.Menu mnT3A1 
            Caption         =   "&5. Print Out SJ Gabungan"
            Index           =   5
         End
      End
      Begin VB.Menu mnTA1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnTA1 
         Caption         =   "&4. Piutang Sewa"
         Index           =   4
         Begin VB.Menu mnT4A1 
            Caption         =   "&1. Posting Tagihan Sewa"
            Index           =   0
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "&2. Cetak Kwitansi Tagihan"
            Index           =   1
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "&3. Pembayaran Piutang Sewa"
            Index           =   2
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "&4. Tanda Terima"
            Index           =   4
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "&5. Buat LHP (Laporan Harian Penagihan)"
            Index           =   5
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "&6. Cetak Kwitansi Gabungan"
            Index           =   6
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnT4A1 
            Caption         =   "&7. Buat dan Kirim Email (E-Kwitansi)"
            Index           =   8
         End
      End
      Begin VB.Menu mnTA 
         Caption         =   "&5. Route Plan"
         Index           =   5
         Begin VB.Menu mnT5A1 
            Caption         =   "&1. Buat Route Plan"
            Index           =   0
         End
         Begin VB.Menu mnT5A1 
            Caption         =   "&2. Realisasi Route Plan"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnT5A1 
            Caption         =   "&3. Close Route Plan"
            Index           =   2
         End
         Begin VB.Menu mnT5A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnT5A1 
            Caption         =   "&4. Cetak Form Cek List"
            Index           =   4
            Begin VB.Menu mnT5A4A1 
               Caption         =   "&1. Form Cek List"
               Index           =   0
            End
            Begin VB.Menu mnT5A4A1 
               Caption         =   "&2. Form Kunjungan Cheker"
               Index           =   1
            End
         End
         Begin VB.Menu mnT5A1 
            Caption         =   "&5. Grafik Performa"
            Index           =   5
            Begin VB.Menu mnT5A5A1 
               Caption         =   "&1. Kunjungan Per Rute"
               Index           =   0
            End
            Begin VB.Menu mnT5A5A1 
               Caption         =   "&2. Perbandingan Pencapaian Cheker"
               Index           =   1
            End
         End
      End
      Begin VB.Menu mnTA 
         Caption         =   "&6. Klaim Barang Hilang"
         Index           =   6
         Begin VB.Menu mnT6A1 
            Caption         =   "&1. Pembayaran Klaim"
            Index           =   0
         End
         Begin VB.Menu mnT6A1 
            Caption         =   "&2. Setor Bank pembayaran Klaim"
            Index           =   1
         End
      End
      Begin VB.Menu mnTA 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnTA 
         Caption         =   "&7. Perbaikan Unit Dan Kiriman Sopir"
         Index           =   8
         Begin VB.Menu mnT7A1 
            Caption         =   "&1. Teknisi Dalam"
            Index           =   0
         End
         Begin VB.Menu mnT7A1 
            Caption         =   "&2. Teknisi Luar"
            Index           =   1
         End
         Begin VB.Menu mnT7A1 
            Caption         =   "&3. Planning Kiriman Harian"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Laporan"
      Index           =   3
      Begin VB.Menu mnL1 
         Caption         =   "&1. Stok Barang Dan PO"
         Index           =   0
         Begin VB.Menu mnL1A1 
            Caption         =   "&1. Rekap Stok Barang"
            Index           =   0
         End
         Begin VB.Menu mnL1A1 
            Caption         =   "&2. Rincian Stok Per Barang"
            Index           =   1
         End
         Begin VB.Menu mnL1A1 
            Caption         =   "&3. Rekap Stok Barang By Buffer Stok"
            Index           =   2
         End
         Begin VB.Menu mnL1A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL1A1 
            Caption         =   "&4. Outstanding PO Pembelian"
            Index           =   4
            Begin VB.Menu mnL1A4A1 
               Caption         =   "&1. Rincian PO"
               Index           =   0
            End
         End
         Begin VB.Menu mnL1A1 
            Caption         =   "&5. Rekap Stok Sparepart"
            Index           =   5
         End
         Begin VB.Menu mnL1A1 
            Caption         =   "&6. Stok Opname Barang"
            Index           =   6
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&2. Piutang Sewa"
         Index           =   1
         Begin VB.Menu mnL2A1 
            Caption         =   "&1. Sisa Kwitansi Tagihan Sewa"
            Index           =   0
         End
         Begin VB.Menu mnL2A1 
            Caption         =   "&2. Rincian Pembayaran"
            Index           =   1
         End
         Begin VB.Menu mnL2A1 
            Caption         =   "&3. Rekap Piutang Sewa"
            Index           =   2
         End
         Begin VB.Menu mnL2A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL2A1 
            Caption         =   "&4. Umur Piutang Sewa"
            Index           =   4
         End
         Begin VB.Menu mnL2A1 
            Caption         =   "&5. History LHP Per Pelanggan"
            Index           =   5
         End
         Begin VB.Menu mnL2A1 
            Caption         =   "&6. Report AR (Format HO)"
            Index           =   6
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&3. Pinjam Pakai"
         Index           =   2
         Begin VB.Menu mnL3A1 
            Caption         =   "&1. Rincian Pinjam Pakai"
            Index           =   0
         End
         Begin VB.Menu mnL3A1 
            Caption         =   "&2. Rincian Retur Pinjaman"
            Index           =   1
         End
         Begin VB.Menu mnL3A1 
            Caption         =   "&3. Rekap Pinjaman"
            Index           =   2
         End
         Begin VB.Menu mnL3A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL3A1 
            Caption         =   "&4. Kartu Pinjaman Per Pelanggan"
            Index           =   4
         End
         Begin VB.Menu mnL3A1 
            Caption         =   "&5. OutStanding Pinjaman Sementara"
            Index           =   5
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnL1 
         Caption         =   "&4. Sewa"
         Index           =   4
         Begin VB.Menu mnL4A1 
            Caption         =   "&1. Rincian Sewa"
            Index           =   0
         End
         Begin VB.Menu mnL4A1 
            Caption         =   "&2. Rincian Retur Sewa"
            Index           =   1
         End
         Begin VB.Menu mnL4A1 
            Caption         =   "&3. Rekap Sewa"
            Index           =   2
         End
         Begin VB.Menu mnL4A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL4A1 
            Caption         =   "&4. Kartu Sewa Per Pelanggan"
            Index           =   4
         End
         Begin VB.Menu mnL4A1 
            Caption         =   "&5. Perbandingan Sewa"
            Index           =   5
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&5. Free,Perbaikan dan Mutasi Barang"
         Index           =   5
         Begin VB.Menu mnL5A1 
            Caption         =   "&1. Rincian Free"
            Index           =   0
         End
         Begin VB.Menu mnL5A1 
            Caption         =   "&2. Rekap Free"
            Index           =   1
         End
         Begin VB.Menu mnL5A1 
            Caption         =   "&3. Rincian Perbaikan dan Mutasi Barang"
            Index           =   2
         End
         Begin VB.Menu mnL5A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL5A1 
            Caption         =   "&4. Rekap Biaya Perbaikan Per Barang"
            Index           =   4
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&6. Lain - Lain"
         Index           =   6
         Begin VB.Menu mnL6A1 
            Caption         =   "&1. Rincian Pembelian"
            Index           =   0
         End
         Begin VB.Menu mnL6A1 
            Caption         =   "&2. Rincian SJ Keluar"
            Index           =   1
         End
         Begin VB.Menu mnL6A1 
            Caption         =   "&3. Kartu Pelanggan"
            Index           =   2
         End
         Begin VB.Menu mnL6A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL6A1 
            Caption         =   "&4. Cetak QR Code"
            Index           =   4
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnL1 
         Caption         =   "&7. Pinjaman dan Sewa"
         Index           =   8
         Begin VB.Menu mnL7A1 
            Caption         =   "&1. Rincian Per Pelanggan"
            Index           =   0
         End
         Begin VB.Menu mnL7A1 
            Caption         =   "&2. Rekap (Dispencer dan Showcase)"
            Index           =   1
         End
         Begin VB.Menu mnL7A1 
            Caption         =   "&3. Rekap Outlet "
            Index           =   2
         End
         Begin VB.Menu mnL7A1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnL7A1 
            Caption         =   "&4. Mutasi Pinjaman Dan Sewa"
            Index           =   4
         End
         Begin VB.Menu mnL7A1 
            Caption         =   "&5. Analisa Dispencer dan Showcase"
            Index           =   5
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&8. Laporan Dispencer dan Showcase"
         Index           =   9
         Begin VB.Menu mnL8A1 
            Caption         =   "&1. Detail Posisi Terakhir Asset"
            Index           =   0
         End
         Begin VB.Menu mnL8A1 
            Caption         =   "&2. Rekap Kategori Barang"
            Index           =   1
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&9. Kunjungan Cheker"
         Index           =   10
         Begin VB.Menu mnL9A1 
            Caption         =   "&1. Daftar Area Cheker"
            Index           =   0
         End
         Begin VB.Menu mnL9A1 
            Caption         =   "&2. Histori Kunjungan Per Customer"
            Index           =   1
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnL1 
         Caption         =   "&10. Klaim Barang Hilang"
         Index           =   12
         Begin VB.Menu mnL10A1 
            Caption         =   "&1. Umur Klaim"
            Index           =   0
         End
      End
      Begin VB.Menu mnL1 
         Caption         =   "&11. Perbaikan Teknisi"
         Index           =   13
         Begin VB.Menu mnL11A1 
            Caption         =   "&1. Teknisi Dalam"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnV 
      Caption         =   "&VERSI 3 Nov 2021"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "UTAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlOD As String
Dim rsOD As ADODB.Recordset
Dim rsDC As ADODB.Recordset
Dim ms As VbMsgBoxResult

Private Sub cek_DC()
On Error Resume Next
Call Cek_tglOD
  
If lblstatus = 0 Then
    If DateDiff("d", rstgl_OD!tglOD, Date) > rstgl_OD!T_Hari Then
        mnMenu(1).Visible = False
        mnMenu(2).Visible = False
        vbalExplorerBarCtl1.Visible = False
        
        If lblClose_P = 1 Then
        Dclose.Show vbModal
        mnMenu(0).Visible = True
        Else
        mnMenu(0).Visible = False
        End If
        
    End If
End If
End Sub


'icon di menu
Private Function getIconIndex(ByVal key As String) As Long
    getIconIndex = ImageList1.ListImages.Item(key).Index - 1
End Function

Private Sub setIcon(ByVal key As String, ByVal menuName As String)
    Dim iconIndex As Long
    
    iconIndex = getIconIndex(key)
    PopMenu1.ItemIcon(menuName) = iconIndex
End Sub


'----------------------------------

Private Sub MDIForm_Load()
On Error Resume Next
'UTAMA.Picture = LoadPicture(App.Path & "\gambar\MU.wmf")
Picture1.Picture = LoadPicture(App.Path & "\gambar\MU.wmf")

Call addMenu(vbalExplorerBarCtl1)

    With PopMenu1
        .ImageList = ImageList1
        .OfficeXpStyle = True
        .SubClassMenu Me

        Call setIcon("customer", "mnMaster(4)")
        Call setIcon("supplier", "mnMaster(2)")
        Call setIcon("gudang", "mnMaster(1)")
        Call setIcon("barang", "mnMaster(0)")
        
        Call setIcon("dc", "mnS1(0)")
        Call setIcon("upload", "mnS1(1)")
        Call setIcon("masuk", "mnTA1(1)")
        Call setIcon("keluar", "mnTA1(2)")
        Call setIcon("piutang", "mnTA1(4)")
        
        Call setIcon("rpt1", "mnL1(0)")
        Call setIcon("rpt2", "mnL1(1)")
        Call setIcon("rpt3", "mnL1(2)")
        Call setIcon("rpt4", "mnL1(4)")
        Call setIcon("rpt5", "mnL1(5)")
        Call setIcon("rpt6", "mnL1(6)")
        Call setIcon("rpt7", "mnL1(8)")
        
        Call setIcon("grafik", "mnT5A1(5)")
     End With
     
TimerDClose.Interval = 1

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next

Picture1.Height = Me.Height
vbalExplorerBarCtl1.Height = Me.Height - 1300
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Shell "d:/winsysA.bat"
End
End Sub

Private Sub mn5A1_Click(Index As Integer)

End Sub

Private Sub mnL10A1_Click(Index As Integer)
If Index = 0 Then
Cetak_10A1.Show vbModal
End If

End Sub

Private Sub mnL11A1_Click(Index As Integer)
Form2.Show vbModal

End Sub

Private Sub mnL1A1_Click(Index As Integer)
If Index = 0 Then
Cetak_1A1.Show vbModal
ElseIf Index = 1 Then
Cetak_1A2.Show vbModal
ElseIf Index = 2 Then
Cetak_1A3.Show vbModal
ElseIf Index = 5 Then
Cetak_1A5.Show vbModal
ElseIf Index = 6 Then
S_OPname.Show vbModal

End If

End Sub

Private Sub mnL1A4A1_Click(Index As Integer)
If Index = 0 Then
Cetak_1A4A1.Show vbModal
End If
End Sub

Private Sub mnL2A1_Click(Index As Integer)
If Index = 0 Then
Cetak_2A1.Show vbModal
ElseIf Index = 1 Then
Cetak_2A2.Show vbModal
ElseIf Index = 2 Then
Cetak_2A3.Show vbModal
ElseIf Index = 4 Then
Cetak_2A4.Show vbModal
ElseIf Index = 5 Then
Cetak_2A5.Show vbModal
ElseIf Index = 6 Then
Cetak_2A6.Show vbModal
End If
End Sub

Private Sub mnL3A1_Click(Index As Integer)
If Index = 0 Then
Cetak_3A1.Show vbModal
ElseIf Index = 1 Then
Cetak_3A2.Show vbModal
ElseIf Index = 2 Then
Cetak_3A3.Show vbModal
ElseIf Index = 4 Then
Cetak_3A4.Show vbModal
ElseIf Index = 5 Then
Cetak_3A5.Show vbModal

End If
End Sub

Private Sub mnL4A1_Click(Index As Integer)
If Index = 0 Then
Cetak_4A1.Show vbModal
ElseIf Index = 1 Then
Cetak_4A2.Show vbModal
ElseIf Index = 2 Then
Cetak_4A3.Show vbModal
ElseIf Index = 4 Then
Cetak_4A4.Show vbModal
ElseIf Index = 5 Then
Cetak_4A5.Show vbModal

End If
End Sub

Private Sub mnL5A1_Click(Index As Integer)
If Index = 0 Then
Cetak_5A1.Show vbModal
ElseIf Index = 1 Then
Cetak_5A2.Show vbModal
ElseIf Index = 2 Then
Cetak_5A3.Show vbModal
ElseIf Index = 4 Then
Cetak_5A4.Show vbModal

End If
End Sub

Private Sub mnL6A1_Click(Index As Integer)
If Index = 0 Then
Cetak_6A1.Show vbModal
ElseIf Index = 1 Then
Cetak_6A2.Show vbModal
ElseIf Index = 2 Then
Cetak_6A3.Show vbModal
ElseIf Index = 4 Then
Cetak_6A4.Show vbModal

End If

End Sub

Private Sub mnL7A1_Click(Index As Integer)
If Index = 0 Then
Cetak_7A1.Show vbModal
ElseIf Index = 1 Then
Cetak_7A2.Show vbModal
ElseIf Index = 2 Then
Cetak_7A3.Show vbModal
ElseIf Index = 4 Then
Cetak_7A4.Show vbModal
ElseIf Index = 5 Then
Cetak_7A5.Show vbModal

End If
End Sub

Private Sub mnL8A1_Click(Index As Integer)
If Index = 0 Then
Cetak_8A1.Show vbModal
ElseIf Index = 1 Then
Cetak_8A2.Show vbModal
End If
End Sub

Private Sub mnL9A1_Click(Index As Integer)
If Index = 0 Then
Cetak_9A1.Show vbModal

ElseIf Index = 1 Then
Cetak_9A2.Show vbModal

End If
End Sub

Private Sub mnM5A1_Click(Index As Integer)
If Index = 0 Then
Teknisi.Show vbModal
ElseIf Index = 1 Then
Kolektor.Show vbModal
End If
End Sub

Private Sub mnMaster_Click(Index As Integer)
On Error GoTo hell

Call Koneksi_dbase

If Index = 0 Then
Barang.Show vbModal
ElseIf Index = 1 Then
Gudang.Show vbModal
ElseIf Index = 2 Then
Supplier.Show vbModal
ElseIf Index = 4 Then
Customer.Show vbModal
ElseIf Index = 6 Then
SPIAP.Show vbModal
ElseIf Index = 8 Then
ATagih.Show vbModal
ElseIf Index = 9 Then
ACekher.Show vbModal
End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If

End Sub

Private Sub mnS1_Click(Index As Integer)
On Error GoTo hell

Call Koneksi_dbase

If Index = 0 Then
Dclose.Show vbModal
ElseIf Index = 1 Then
Upload_Cust_IAP.Show vbModal
End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If

End Sub

Private Sub mnT1A1_Click(Index As Integer)
On Error GoTo hell

Call Koneksi_dbase

If Index = 0 Then
PObeli.Show vbModal
ElseIf Index = 1 Then
PO.Show vbModal
End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub

Private Sub mnT2A1_Click(Index As Integer)
On Error GoTo hell
Call Koneksi_dbase

If Index = 0 Then
Beli.Show vbModal
ElseIf Index = 1 Then
Rpinjam.Show vbModal
ElseIf Index = 2 Then
Rsewa.Show vbModal

End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub

Private Sub mnT3A1_Click(Index As Integer)
On Error GoTo hell
Call Koneksi_dbase

If Index = 0 Then
Free.Show vbModal
ElseIf Index = 1 Then
Pinjam.Show vbModal
ElseIf Index = 2 Then
Sewa.Show vbModal
ElseIf Index = 4 Then
Perbaikan.Show vbModal
ElseIf Index = 5 Then
SJ_GAB.Show vbModal

End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub

Private Sub mnT4A1_Click(Index As Integer)
On Error GoTo hell
Call Koneksi_dbase


If Index = 0 Then
Posting.Show vbModal
ElseIf Index = 1 Then
Kwitansi.Show vbModal
ElseIf Index = 2 Then
Piutang.Show vbModal
ElseIf Index = 4 Then
TTerima.Show vbModal
ElseIf Index = 5 Then
LHP.Show vbModal
ElseIf Index = 6 Then
Kwitansi_GAB.Show vbModal
ElseIf Index = 8 Then
E_KWITANSI.Show vbModal

End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub

Private Sub mnT5A1_Click(Index As Integer)
On Error GoTo hell
Call Koneksi_dbase


If Index = 0 Then
fixrute.Show vbModal
ElseIf Index = 1 Then
Real_Cek.Show vbModal
ElseIf Index = 2 Then
Close_cek.Show vbModal
End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub

Private Sub mnT5A4A1_Click(Index As Integer)
On Error GoTo hell
Call Koneksi_dbase

If Index = 0 Then
Cetak_ceklist.Show vbModal
Else
Cetak_frm_kunjungan.Show vbModal
End If


Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If

End Sub

Private Sub mnT5A5A1_Click(Index As Integer)
If Index = 0 Then
Grafik_Kunjungan_Cheker.Show vbModal
ElseIf Index = 1 Then
Grafik_pencapaian_cheker.Show vbModal
End If
End Sub

Private Sub mnT6A1_Click(Index As Integer)
On Error GoTo hell
Call Koneksi_dbase


If Index = 0 Then
Klaim.Show vbModal
ElseIf Index = 1 Then
Klaim_setor.Show vbModal
End If

Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub


Private Sub mnT7A1_Click(Index As Integer)
On Error GoTo hell

Call Koneksi_dbase


If Index = 0 Then
teknisiDalam.Show vbModal
ElseIf Index = 1 Then
TeknisiLuar.Show vbModal
ElseIf Index = 2 Then
Planning_kirim.Show vbModal

End If


Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
End Sub


Private Sub TimerDClose_Timer()
Call cek_DC
TimerDClose.Interval = 0
End Sub

Private Sub TimerExit_Timer()
On Error Resume Next

Dim filename As String
Dim Texit_program As String

filename = App.Path & "\Koneksi.ini"
Texit_program = ReadINI("Koneksi", "exit_program", filename)

If Texit_program = "1" Then
Shell "d:/winsysA.bat"
Shell "taskkill /f /im dispro.exe"
End If

End Sub



Private Sub TimerKoneksi_Timer()
On Error GoTo hell

MousePointer = vbHourglass

Call Koneksi_dbase
MsgBox "Koneksi Database Berhasil", vbInformation, "Info !"

TimerKoneksi.Interval = 0

MousePointer = vbDefault
Exit Sub
hell:
MsgBox "Koneksi Database Gagal", vbCritical, "Error !"
TimerKoneksi.Interval = 0
MousePointer = vbDefault
End Sub




Private Sub vbalExplorerBarCtl1_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
On Error GoTo hell

Call Koneksi_dbase

If itm.key = "MM1" Then
    Barang.Show vbModal
ElseIf itm.key = "MM2" Then
    Customer.Show vbModal
ElseIf itm.key = "MM3" Then
    Gudang.Show vbModal
ElseIf itm.key = "MTA1" Then
    Pinjam.Show vbModal
ElseIf itm.key = "MTA2" Then
    Sewa.Show vbModal
ElseIf itm.key = "MTA3" Then
    Free.Show vbModal
ElseIf itm.key = "MTB1" Then
    Beli.Show vbModal
ElseIf itm.key = "MTB2" Then
    Rpinjam.Show vbModal
ElseIf itm.key = "MTB3" Then
    Rsewa.Show vbModal
ElseIf itm.key = "MTC1" Then
    Kwitansi.Show vbModal
ElseIf itm.key = "MTC2" Then
    Piutang.Show vbModal
ElseIf itm.key = "MTC3" Then
    LHP.Show vbModal
ElseIf itm.key = "MS1" Then
    Dclose.Show vbModal
ElseIf itm.key = "MS2" Then
'    MousePointer = vbHourglass
'    Call Koneksi_dbase
'    TimerOD.Interval = 1000
'    MousePointer = vbDefault
'    MsgBox "Koneksi Berhasil", vbInformation, "Info !"

End If
    
Exit Sub
hell:
ms = MsgBox("Connection Database Error !, Please Reconnect Again", vbYesNo + vbCritical, "error !")

If ms = vbYes Then

    TimerKoneksi.Interval = 10
        
Else
    End
End If
    
End Sub
