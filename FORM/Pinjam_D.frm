VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Pinjam_D 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   15690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblnosj 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3465
      TabIndex        =   7
      Top             =   2430
      Width           =   1590
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00000000&
      Caption         =   "SJ MANUAL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1485
      TabIndex        =   6
      Top             =   2430
      Width           =   1275
   End
   Begin VB.TextBox txtnoPP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8055
      TabIndex        =   5
      Top             =   2070
      Width           =   1590
   End
   Begin VB.TextBox txttglkembali 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10980
      TabIndex        =   3
      Top             =   1350
      Width           =   1590
   End
   Begin VB.ComboBox CMBStatus 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7785
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1350
      Width           =   1905
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      TabIndex        =   4
      Top             =   2070
      Width           =   4335
   End
   Begin VB.TextBox txttglpinjam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4275
      TabIndex        =   0
      Top             =   990
      Width           =   1590
   End
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
   End
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin VB.Timer TimerNO 
      Left            =   7290
      Top             =   720
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   17
      Top             =   720
      Width           =   14415
      _Version        =   524288
      _ExtentX        =   25426
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   12105
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   945
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   741
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   735
      Left            =   14760
      TabIndex        =   8
      ToolTipText     =   "Simpan"
      Top             =   990
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":2832
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   18
      Top             =   2790
      Width           =   14595
      _Version        =   524288
      _ExtentX        =   25744
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   0
      Left            =   14760
      TabIndex        =   9
      ToolTipText     =   "Tambah"
      Top             =   2925
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16744576
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":529F
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   1
      Left            =   14760
      TabIndex        =   11
      ToolTipText     =   "Ubah"
      Top             =   3690
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":7F13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   2
      Left            =   14760
      TabIndex        =   12
      ToolTipText     =   "Hapus"
      Top             =   4455
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":B110
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   3
      Left            =   14760
      TabIndex        =   13
      ToolTipText     =   "Refresh"
      Top             =   5220
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":E1A9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   4
      Left            =   14760
      TabIndex        =   14
      ToolTipText     =   "Cetak"
      Top             =   5985
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":11325
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   16
      Top             =   8280
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   661
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":14D82
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBatal 
      Height          =   735
      Left            =   14760
      TabIndex        =   10
      ToolTipText     =   "Batal"
      Top             =   1755
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Pinjam_D.frx":1B5E4
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   5010
      Left            =   225
      TabIndex        =   15
      Top             =   2880
      Width           =   14415
      _cx             =   25426
      _cy             =   8837
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777152
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Pinjam_D.frx":1E883
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   4
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "NO EASAP :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5715
      TabIndex        =   42
      Top             =   2475
      Width           =   825
   End
   Begin VB.Label lblnoEASAP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6570
      TabIndex        =   41
      Top             =   2430
      Width           =   2040
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "NO SJ :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2880
      TabIndex        =   40
      Top             =   2475
      Width           =   645
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   5940
      TabIndex        =   39
      Top             =   8865
      Width           =   915
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "NO. FORM PINJAM PAKAI :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5850
      TabIndex        =   38
      Top             =   2115
      Width           =   2265
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL KEMBALI :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9810
      TabIndex        =   37
      Top             =   1395
      Width           =   1905
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS PJMAN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6435
      TabIndex        =   36
      Top             =   1395
      Width           =   1320
   End
   Begin VB.Label lblkdcustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1080
      TabIndex        =   35
      Top             =   1710
      Width           =   1140
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   34
      Top             =   1755
      Width           =   1050
   End
   Begin VB.Label lblnmcustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2250
      TabIndex        =   33
      Top             =   1710
      Width           =   4065
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3420
      TabIndex        =   32
      Top             =   1035
      Width           =   870
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMER :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   31
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label lblKDPinjam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1080
      TabIndex        =   30
      Top             =   990
      Width           =   2175
   End
   Begin VB.Label lbltglPO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   10530
      TabIndex        =   29
      Top             =   990
      Width           =   1590
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   315
      TabIndex        =   28
      Top             =   2115
      Width           =   1320
   End
   Begin VB.Label lblnmgudang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2250
      TabIndex        =   27
      Top             =   1350
      Width           =   4065
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GUDANG :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   26
      Top             =   1395
      Width           =   825
   End
   Begin VB.Label lblkdgudang 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1080
      TabIndex        =   25
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pinjam Pakai"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Index           =   1
      Left            =   990
      TabIndex        =   24
      Top             =   45
      Width           =   8025
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL PO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9720
      TabIndex        =   23
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label txtkdPO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7335
      TabIndex        =   22
      Top             =   990
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6615
      TabIndex        =   21
      Top             =   1035
      Width           =   645
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3690
      TabIndex        =   20
      Top             =   8775
      Width           =   1545
   End
   Begin VB.Label lblalamat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6345
      TabIndex        =   19
      Top             =   1710
      Width           =   6225
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "Pinjam_D.frx":1E9A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15585
   End
End
Attribute VB_Name = "Pinjam_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rsN As ADODB.Recordset
Dim sqlN As String
Dim rs2 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK, rsT As ADODB.Recordset
Dim a As Integer
Dim b As Currency
Dim rsACC As ADODB.Recordset
Dim rscp As ADODB.Recordset
Dim rsR As ADODB.Recordset
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim rsCS As ADODB.Recordset
Dim rsOD As ADODB.Recordset
Dim color As Long, flag As Byte





Private Sub cek_retur()
sqlR = "select * from RPinjam_d where kdpinjam='" & lblkdPinjam & "'  "
Set rsR = con.Execute(sqlR)
End Sub

Private Sub Chk_Click()
If Chk.Value = 1 Then
    lblnosj.Enabled = True
    
    
        
ElseIf Chk.Value = 0 Then
    lblnosj.Enabled = False
End If
End Sub

Private Sub Chk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub CMBStatus_Click()
If CMBStatus.ListIndex = 0 Then
txttglkembali.Enabled = True
Else
txttglkembali.Enabled = False
txttglkembali = txttglpinjam
End If
End Sub

Private Sub CMBStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdBatal_Click()
On Error GoTo hell




Call Cek_tglOD
If CDate(txttglpinjam) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    Call cek_retur

    If rsR.RecordCount <> 0 Then
    MsgBox "Pinjaman Tidak Dapat dibatalkan, Karena sudah pernah diretur !", vbCritical, "Error !"
    Exit Sub
    Else
    
         ms = MsgBox("Apakah anda ingin Membatalkan Pengeluaran Barang ini ?", vbYesNo + vbQuestion, "Info")
         If ms = vbYes Then
            sql = "update PO set kdkeluar='' where kdPO='" & txtkdPO & "'"
            con.Execute (sql)
            
            sql = "delete from pinjam_d where kdpinjam='" & lblkdPinjam & "'"
            con.Execute (sql)
            
            sql = "delete from pinjam where kdpinjam='" & lblkdPinjam & "'"
            con.Execute (sql)
            
            txtkdPO = ""
            txttglPO = ""
            cmdBR.Enabled = True
            txttglpinjam = Date
            txttglpinjam.Enabled = True
            CMBStatus.Enabled = True
            CMBStatus.ListIndex = 0
            txttglkembali.Enabled = True
            txtketerangan.Enabled = True
            
            
            lblkdgudang = ""
            lblnmgudang = ""
            lblkdcustomer = ""
            lblnmcustomer = ""
            
            txtketerangan = ""
            lblkode = 1
            
            
            TimerAll.Interval = 10
            Pinjam.TimerAll.Interval = 10
        Else
            Exit Sub
        End If
    
    End If
    
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub cmdBatal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Cetak()
MousePointer = vbHourglass

sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
         "from RKP_stok where kdgudang='" & lblkdgudang & "' and tgl <= '" & Format(txttglpinjam, "yyyy/MM/dd") & "' and kdbarang in (select kdbarang from pinjam_d where kdpinjam='" & lblkdPinjam & "') group by kdbarang"

sqlCS = "select * from (" & sqlCS1 & ") a where unit < 0 order by kdbarang"

Set rsCS = con.Execute(sqlCS)

If rsCS.RecordCount <> 0 Then
    ms = MsgBox("Stok Barang Kurang , Tampilkan List Barang ?", vbCritical + vbYesNo, "Error !")
    If ms = vbYes Then
    List_Stok_selisih.lblkode = "pinjam"
    List_Stok_selisih.Show vbModal
    End If
    
Else
    
    
    Unload AR_SJ
    
    sqlX = "select a.kdbarang,b.kd1,b.nmbarang,b.kdkategori,a.unit,b.satuan,a.keterangan from pinjam_d a left join barang b " & vbCrLf & _
           "on a.kdbarang=b.kdbarang where a.kdpinjam='" & lblkdPinjam & "' order by a.kdbarang"
    
    Set rsX = con.Execute(sqlX)
    
    With AR_SJ.DC1
    .ConnectionString = koneksi
    .Source = sqlX
    End With
    
    With AR_SJ
    .fldunit.DataField = "unit"
    .fldnmbarang.DataField = "nmbarang"
    .fldsatuan.DataField = "satuan"
    .fldketerangan.DataField = "keterangan"
    .fldkdbarang.DataField = "kdbarang"
    .fldkdkategori.DataField = "kdkategori"
    
    .lblj_EASAP = "NO EASAP :"
    .lblno_EASAP = lblnoEASAP
    
    
    .lblnosj = lblkdPinjam
    .lblnosj1 = lblnosj
    .lblnmcustomer = lblnmcustomer
    .lbltglSJ = Format(txttglpinjam, "dd/MM/yyyy")
    .lblalamat = lblalamat
    
    If txtketerangan = "" Then
    .lblNB = ""
    Else
    .lblNB = "NB : " & txtketerangan
    End If
    
    sqlACC = "select * from Signature where kdFrm='" & lblkdgudang & "'"
    Set rsACC = con.Execute(sqlACC)
    
    .lblAcc1 = rsACC!Acc1
    .lblAcc2 = rsACC!Acc2
    .lblAcc3 = rsACC!Acc3
    .lblAcc4 = rsACC!Acc4
    
    sqlCP = "select * from customer where kdcustomer='" & lblkdcustomer & "'"
    Set rscp = con.Execute(sqlCP)
    
    .lblCP = rscp!CP
    .lbltelp = rscp!telp
    
    
    AR_SJ.Show vbModal
    
    End With

End If

MousePointer = vbDefault
End Sub


Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub

Private Sub tbl()
If rs.RecordCount = 0 Then
    cmdT(1).Enabled = False
    cmdT(2).Enabled = False
    datagrid1.Enabled = False

Else
    cmdT(1).Enabled = True
    cmdT(2).Enabled = True
    datagrid1.Enabled = True
End If
End Sub


Private Sub LG()
On Error GoTo hell

Call tbl

Exit Sub
hell:
End Sub


Private Sub all()
MousePointer = vbHourglass

sql = "select a.kdbarang,b.kd1,b.nmbarang,a.unit,b.satuan,a.harga,a.rupiah,a.keterangan,a.kdpinjam_d from pinjam_d a left join barang b " & vbCrLf & _
      "on a.kdbarang=b.kdbarang where a.kdpinjam='" & lblkdPinjam & "' order by a.kdbarang "
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


Call LG
MousePointer = vbDefault
End Sub



Private Sub tbh()
Call Cek_tglOD
If CDate(txttglpinjam) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else
    If txttglpinjam.Enabled = False Then
    Pinjam_DTU.lblkode = 1
    
    Pinjam_DTU.Show vbModal
    
    Else
    MsgBox "Kepala data belum disimpan !", vbCritical, "INfo !!"
    End If
End If


End Sub


Private Sub ubh()
Call Cek_tglOD
If CDate(txttglpinjam) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    Pinjam_DTU.lblkode = 2
    
    
    lblpos = rs.AbsolutePosition
    kode = 2
    
    
    Pinjam_DTU.lblkdbarang = rs!kdbarang
    Pinjam_DTU.lblnmbarang = rs!nmbarang
    Pinjam_DTU.lblsatuan = rs!satuan
    Pinjam_DTU.txtunit = FormatNumber(rs!unit, 0)
    Pinjam_DTU.txtharga = FormatNumber(rs!harga, 0)
    Pinjam_DTU.lblrupiah = FormatNumber(rs!rupiah, 0)
    Pinjam_DTU.txtketerangan = rs!keterangan
    Pinjam_DTU.lblkdpinjam_d = rs!kdpinjam_d
    Pinjam_DTU.lblunit_awal = rs!unit
    Pinjam_DTU.cmdBR.Enabled = False
    
    'Pinjam_DTU.txtunit.Enabled = False
      
    Pinjam_DTU.Show vbModal
End If
End Sub


Private Sub hps()
On Error GoTo hell
Call Cek_tglOD
If CDate(txttglpinjam) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else


    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from pinjam_d where kdpinjam_d ='" & rs!kdpinjam_d & "'"
        con.Execute (sql)
        TimerAll.Interval = 10
    End If

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub









Private Sub nomer()
On Error GoTo hell

If lblkode = 1 Then
    sql = "select isnull(max(right(kdpinjam,4)),0) as xx from pinjam where Month(tglpinjam)='" & Month(txttglpinjam) & "'  and year(tglpinjam)='" & Year(txttglpinjam) & "' and kdgudang= '" & lblkdgudang & "'"
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    lblkdPinjam = lblkdgudang & "/E/" & Format(txttglpinjam, "MMyy") & "/" & "000" & a
                Case 2
                    lblkdPinjam = lblkdgudang & "/E/" & Format(txttglpinjam, "MMyy") & "/" & "00" & a
                Case 3
                    lblkdPinjam = lblkdgudang & "/E/" & Format(txttglpinjam, "MMyy") & "/" & "0" & a
                Case 4
                    lblkdPinjam = lblkdgudang & "/E/" & Format(txttglpinjam, "MMyy") & "/" & a
        End Select
    
    Else
        lblkdPinjam = lblkdgudang & "/E/" & Format(txttglpinjam, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
lblkdPinjam = lblkdgudang & "/E/" & Format(txttglpinjam, "MMyy") & "/" & "0001"
End Sub


Private Sub nomer1()
On Error GoTo hell

If Chk.Value = 0 Then

    If lblkode = 1 Then
        sqlN = "select isnull(max(right(nosj,6)),0) as xx from urutSJ where kdgudang= '" & lblkdgudang & "'"
        Set rsN = con.Execute(sqlN)
        
        
        b = CCur(rsN!xx) + 1
            
        If b > 0 Then
        
            Select Case Len(CStr(b))
                    Case 1
                        lblnosj = lblkdgudang & "/" & "00000" & b
                    Case 2
                        lblnosj = lblkdgudang & "/" & "0000" & b
                    Case 3
                        lblnosj = lblkdgudang & "/" & "000" & b
                    Case 4
                        lblnosj = lblkdgudang & "/" & "00" & b
                    Case 5
                        lblnosj = lblkdgudang & "/" & "0" & b
                    Case 6
                        lblnosj = lblkdgudang & "/" & b
            End Select
        
        Else
            lblnosj = lblkdgudang & "/" & "000001"
        
        End If
    
    End If
End If

Exit Sub
hell:
lblnosj = lblkdgudang & "/" & "000001"
End Sub


Private Sub cmdBR_Click()
PO_BR.lblkode = "PINJAM_D"
PO_BR.lblkdkategori = "02"
PO_BR.Show vbModal

End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
Call ubh
ElseIf Index = 2 Then
Call hps
ElseIf Index = 3 Then
Call all
ElseIf Index = 4 Then
Call Cetak
End If

End Sub

Private Sub cmdT_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
 TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
End If
End Sub


Private Sub cmdsimpan_Click()
Call Cek_tglOD
If CDate(txttglpinjam) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    If txtkdPO = "" Or lblkdgudang = "" Or lblkdcustomer = "" Or txtnoPP = "" Then
    MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
    Exit Sub
    Else
    
        If lblkode = 1 Then
            Call nomer
            Call nomer1
            
            sqlA = "select a.kdbarang,sum(a.unit) as UKeluar,b.kdpo from pinjam_d a left join pinjam b  on a.kdpinjam =b.kdpinjam where b.kdpo ='" & txtkdPO & "' group by a.kdbarang,b.kdpo"
            
            sqlA1 = "select a.kdbarang,b.nmbarang,a.unit,isnull(c.Ukeluar,0) as Ukeluar,b.satuan,a.keterangan,a.kdpo_d from po_d a left join barang b " & vbCrLf & _
                    "on a.kdbarang=b.kdbarang left join (" & sqlA & ") c on a.kdPO=c.kdPO and a.kdbarang=c.kdbarang where a.kdpo='" & txtkdPO & "' "
            
            sqlA2 = "select kdbarang,nmbarang,unit,ukeluar,unit - ukeluar as sisa,satuan,keterangan,kdPO_D from (" & sqlA1 & ") a  "
            
            sql = "insert into Pinjam values ('" & lblkdPinjam & "','" & Format(txttglpinjam, "yyyy-MM-dd") & "','" & lblkdgudang & "','" & lblkdcustomer & "','" & UCase(txtketerangan) & "','" & txtkdPO & "','" & CMBStatus.Text & "','" & Format(txttglkembali, "yyyy/MM/dd") & "','" & UCase(txtnoPP) & "','" & lblnosj & "'," & Chk.Value & ",'" & lblnoEASAP & "')"
            con.Execute (sql)
            
            sql = "insert into pinjam_d select kdbarang  + '" & "_" & lblkdPinjam & "','" & lblkdPinjam & "',kdbarang,sisa,0,0,keterangan from (" & sqlA2 & ") a where sisa<>0 "
            con.Execute (sql)
            
            sql = "update PO set kdkeluar='" & lblkdPinjam & "' where kdpo ='" & txtkdPO & "'"
            con.Execute (sql)
            
            txttglpinjam.Enabled = False
            cmdBR.Enabled = False
            txtketerangan.Enabled = False
            CMBStatus.Enabled = False
            txttglkembali.Enabled = False
            cmdsimpan.Enabled = False
            cmdBatal.Enabled = True
            txtnoPP.Enabled = False
            Chk.Enabled = False
            lblnosj.Enabled = False
        
        
        ElseIf lblkode = 2 Then
            sql = "Update Pinjam set keterangan='" & UCase(txtketerangan) & "',status='" & CMBStatus.Text & "',tglpengembalian='" & Format(txttglkembali, "yyyy/MM/dd") & "',nopp='" & UCase(txtnoPP) & "',sj_manual=" & Chk.Value & ",nosj='" & UCase(lblnosj) & "'  where kdpinjam='" & lblkdPinjam & "'"
            con.Execute (sql)
        
            txtketerangan.Enabled = False
            CMBStatus.Enabled = False
            txttglkembali.Enabled = False
            cmdsimpan.Enabled = False
            txtnoPP.Enabled = False
            Chk.Enabled = False
            lblnosj.Enabled = False
            
        
            MsgBox "Header berhasil di Ubah ", vbInformation, "Info !"
        End If
     
    End If
     
    Pinjam.TimerAll.Interval = 10
    TimerAll.Interval = 10

End If

End Sub




Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub datagrid1_DblClick()
Call ubh
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyLeft Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyRight Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
 
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttglpinjam = Date
txttglpinjam.Enabled = True

txttglkembali = Date
lblnosj.Enabled = False

CMBStatus.AddItem "SEMENTARA"
CMBStatus.AddItem "TETAP"
CMBStatus.ListIndex = 0

TimerAll.Interval = 10
TimerNO.Interval = 10


Call nul(lblkdgudang)
Call nul(lblnmgudang)
Call nul(txtkdPO)
Call nul(lbltglPO)
Call nul(lblkdcustomer)
Call nul(lblnmcustomer)
Call nul(lblalamat)
Call nul(txtnoPP)

End Sub



Private Sub Form_Unload(Cancel As Integer)
sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
         "from RKP_stok where kdgudang='" & lblkdgudang & "' and tgl <= '" & Format(txttglpinjam, "yyyy/MM/dd") & "' and kdbarang in (select kdbarang from pinjam_d where kdpinjam='" & lblkdPinjam & "') group by kdbarang"

sqlCS = "select * from (" & sqlCS1 & ") a where unit < 0 order by kdbarang"

Set rsCS = con.Execute(sqlCS)

If rsCS.RecordCount <> 0 Then
Cancel = 1
    ms = MsgBox("Stok Barang Kurang, Tampilkan List Barang ?", vbCritical + vbYesNo, "Error !")
    If ms = vbYes Then
    List_Stok_selisih.lblkode = "PINJAM"
    List_Stok_selisih.Show vbModal
    End If
End If

End Sub

Private Sub lblalamat_Change()
Call nul(lblalamat)
End Sub

Private Sub lblkdpinjam_Change()
Call nul(lblkdPinjam)
End Sub

Private Sub lblkdgudang_Change()
Call nul(lblkdgudang)
Call nomer
Call nomer1
End Sub

Private Sub lblkdcustomer_Change()
Call nul(lblkdcustomer)
End Sub

Private Sub lblnmcustomer_Change()
Call nul(lblnmcustomer)
End Sub

Private Sub lblnmgudang_Change()
Call nul(lblnmgudang)
End Sub



Private Sub lblnosj_Change()
Call nul(lblnosj)
End Sub

Private Sub lblnoSJ_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub lblnoSJ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub lblnoSJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub lblnoSJ_LostFocus()
lblnosj = UCase(lblnosj)
End Sub

Private Sub lbltglPO_Change()
Call nul(lbltglPO)
End Sub


Private Sub Text1_Change()

End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If rs.RecordCount = 0 Then
cmdT(1).SetFocus
Else
datagrid1.SetFocus
End If

If kode = 2 Then
rs.AbsolutePosition = lblpos
End If

 

TimerAll.Interval = 0

End Sub

Private Sub TimerNO_Timer()
If lblkode = 1 Then
Call nomer
Call nomer1
End If


TimerNO.Interval = 0
End Sub



Private Sub txtkdPO_Change()
Call nul(txtkdPO)

sql1 = "select * from PO where kdPO='" & txtkdPO & "'"
Set rs1 = con.Execute(sql1)

If rs1.RecordCount <> 0 Then
lbltglPO = rs1!tglPO
End If
End Sub

Private Sub txtketerangan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtketerangan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtketerangan_LostFocus()
txtketerangan = UCase(txtketerangan)
End Sub


Private Sub txtnoPP_Change()
Call nul(txtnoPP)
End Sub

Private Sub txtnoPP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoPP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtnoPP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtnoPP_LostFocus()
txtnoPP = UCase(txtnoPP)
End Sub

Private Sub txttglpinjam_Change()
Call nul(txttglpinjam)
Call nomer
Call nomer1
End Sub

Private Sub txttglpinjam_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglpinjam_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglpinjam_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttglpinjam_LostFocus()
On Error GoTo hell

txttglpinjam = FormatDateTime(txttglpinjam, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglpinjam.SetFocus

End Sub

Private Sub txttglkembali_Change()
Call nul(txttglkembali)

End Sub

Private Sub txttglkembali_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglkembali_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglkembali_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttglkembali_LostFocus()
On Error GoTo hell

txttglkembali = FormatDateTime(txttglkembali, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglkembali.SetFocus

End Sub










