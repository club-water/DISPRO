VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form PO_D 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   16560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CMBket 
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
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   1905
   End
   Begin VB.TextBox txtnoEASAP 
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
      Left            =   10800
      TabIndex        =   8
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Timer TimerCMB 
      Left            =   7920
      Top             =   900
   End
   Begin VB.ComboBox CMBKATEGORI 
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
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1755
      Width           =   2670
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
      Left            =   3375
      TabIndex        =   7
      Top             =   2160
      Width           =   6405
   End
   Begin VB.TextBox txttglPO 
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
      Left            =   4365
      TabIndex        =   0
      Top             =   1035
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
      Left            =   1980
      Top             =   990
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   16
      Top             =   720
      Width           =   14325
      _Version        =   524288
      _ExtentX        =   25268
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   12240
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   990
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
      Picture         =   "PO_D.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   780
      Left            =   14625
      TabIndex        =   9
      ToolTipText     =   "Simpan"
      Top             =   1845
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "PO_D.frx":2832
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   -135
      TabIndex        =   17
      Top             =   2745
      Width           =   14730
      _Version        =   524288
      _ExtentX        =   25982
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   0
      Left            =   14625
      TabIndex        =   10
      ToolTipText     =   "Tambah"
      Top             =   2880
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "PO_D.frx":529F
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   1
      Left            =   14625
      TabIndex        =   11
      ToolTipText     =   "Ubah"
      Top             =   3690
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "PO_D.frx":7F13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   2
      Left            =   14625
      TabIndex        =   12
      ToolTipText     =   "Hapus"
      Top             =   4500
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "PO_D.frx":B110
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   3
      Left            =   14625
      TabIndex        =   13
      ToolTipText     =   "Refresh"
      Top             =   5310
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "PO_D.frx":E1A9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   4
      Left            =   14625
      TabIndex        =   14
      ToolTipText     =   "Cetak"
      Top             =   6120
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "PO_D.frx":11325
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   15
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
      Picture         =   "PO_D.frx":14D82
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   11745
      TabIndex        =   2
      ToolTipText     =   "Simpan"
      Top             =   1350
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
      Picture         =   "PO_D.frx":1B5E4
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdInfo 
      Height          =   420
      Left            =   12195
      TabIndex        =   3
      ToolTipText     =   "Informasi Stok yg Tersedia"
      Top             =   1350
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
      Picture         =   "PO_D.frx":1DE16
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand CmdBR2 
      Height          =   420
      Left            =   12240
      TabIndex        =   5
      Top             =   1710
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
      Picture         =   "PO_D.frx":204B4
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   5010
      Left            =   135
      TabIndex        =   38
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
      FormatString    =   $"PO_D.frx":22CE6
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NO. EASAP"
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
      Left            =   9900
      TabIndex        =   37
      Top             =   2205
      Width           =   1320
   End
   Begin VB.Label lblnmbarang 
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
      Left            =   8235
      TabIndex        =   36
      Top             =   1755
      Width           =   4020
   End
   Begin VB.Label lblkdbarang 
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
      Left            =   6705
      TabIndex        =   35
      Top             =   1755
      Width           =   1500
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "DISPENCER / SHOWCASE :"
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
      Left            =   4635
      TabIndex        =   34
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   285
      Left            =   6210
      TabIndex        =   33
      Top             =   8910
      Width           =   1050
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
      Left            =   6075
      TabIndex        =   32
      Top             =   1395
      Width           =   5685
   End
   Begin VB.Label lblkdkategori 
      Height          =   330
      Left            =   4140
      TabIndex        =   31
      Top             =   1755
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "KATEGORI :"
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
      Left            =   360
      TabIndex        =   30
      Top             =   1800
      Width           =   1320
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
      Left            =   2565
      TabIndex        =   29
      Top             =   1395
      Width           =   3480
   End
   Begin VB.Label Label6 
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
      Left            =   360
      TabIndex        =   28
      Top             =   1440
      Width           =   1005
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
      Left            =   1395
      TabIndex        =   27
      Top             =   1395
      Width           =   1140
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
      Left            =   225
      TabIndex        =   26
      Top             =   2205
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
      Left            =   8190
      TabIndex        =   25
      Top             =   1035
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
      Left            =   6165
      TabIndex        =   24
      Top             =   1080
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
      Left            =   7020
      TabIndex        =   23
      Top             =   1035
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Permintaan Barang"
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
      TabIndex        =   22
      Top             =   45
      Width           =   6000
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
      Left            =   3645
      TabIndex        =   21
      Top             =   1080
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
      Left            =   1395
      TabIndex        =   20
      Top             =   1035
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
      Left            =   675
      TabIndex        =   19
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3690
      TabIndex        =   18
      Top             =   8775
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "PO_D.frx":22E09
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15450
   End
End
Attribute VB_Name = "PO_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK, rsT As ADODB.Recordset
Dim a As Integer
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim color As Long, flag As Byte
Dim rsST As ADODB.Recordset
Dim rscek As ADODB.Recordset
Dim rsB As ADODB.Recordset



Private Sub cek_dalem()
sqlcek = "select * from PO_d where kdPO='" & txtkdPO & "'"
Set rscek = con.Execute(sqlcek)
End Sub


Private Sub set_cmbkeluar()
On Error GoTo hell

sql = "Select * from kategori order by kdkategori"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
cmbkategori.AddItem rs!nmkategori
rs.MoveNext
Loop

If LBLKODE = "1" Then
cmbkategori.ListIndex = 0
End If


        
 
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub


Private Sub cmbkategori_Click()
On Error Resume Next
sql1 = "select * from KATEGORI where nmKATEGORI='" & cmbkategori.Text & "'"
Set rs1 = con.Execute(sql1)

lblkdkategori = rs1!kdkategori

If lblkdkategori = "04" Then
cmdBR2.Visible = True
lblkdbarang.Visible = True
lblnmbarang.Visible = True
lbl1.Visible = True
cmdBR2.Visible = True


Call nul(lblkdbarang)
Call nul(lblnmbarang)
Else
lblkdbarang = ""
lblnmbarang = ""
cmdBR2.Visible = False
lblkdbarang.Visible = False
lblnmbarang.Visible = False
lbl1.Visible = False
cmdBR2.Visible = False

lblkdbarang.BackColor = vbWhite
lblnmbarang.BackColor = vbWhite
End If


If lblkdkategori = "04" Or lblkdkategori = "05" Then
txtnoEASAP.Enabled = False
Else
txtnoEASAP.Enabled = True
End If



End Sub

Private Sub CMBKATEGORI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub CMBket_Click()
If CMBket.ListIndex = 2 Then
txtketerangan = ""
txtketerangan.Enabled = True
Else
txtketerangan = CMBket.Text
txtketerangan.Enabled = False
End If
End Sub

Private Sub cmdBR1_Click()
Customer_br.LBLKODE = "PO_D"
Customer_br.Show vbModal
End Sub

Private Sub cmdBR1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdBR2_Click()
Barang_BR.LBLKODE = UCase("PO_D")
Barang_BR.Show vbModal

End Sub

Private Sub cmdBR2_KeyPress(KeyAscii As Integer)
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

Private Sub cmdInfo_Click()
On Error GoTo hell
sqlA1 = "select kdcustomer,sum(pjm) as pjm,sum(sewa) as sewa from (" & vbCrLf & _
       "select a.kdcustomer,sum(b.unit) as pjm,0 as sewa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttglPO, "yyyy/MM/dd") & "'  and a.kdcustomer='" & lblkdcustomer & "' group by a.kdcustomer" & vbCrLf & _
       "Union" & vbCrLf & _
       "select a.kdcustomer,-sum(b.unit) as pjm,0 as sewa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttglPO, "yyyy/MM/dd") & "' and a.kdcustomer='" & lblkdcustomer & "' group by a.kdcustomer" & vbCrLf & _
       "union" & vbCrLf & _
       "select a.kdcustomer,0 as pjm,sum(b.unit) as sewa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttglPO, "yyyy/MM/dd") & "'  and a.kdcustomer='" & lblkdcustomer & "' group by a.kdcustomer" & vbCrLf & _
       "Union" & vbCrLf & _
       "select a.kdcustomer,0as pjm,-sum(b.unit) as sewa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttglPO, "yyyy/MM/dd") & "' and a.kdcustomer='" & lblkdcustomer & "' group by a.kdcustomer" & vbCrLf & _
       ") a group by kdcustomer"


Set rsST = con.Execute(sqlA1)

    If rsST.RecordCount <> 0 Then
    MsgBox "Jumlah Pinjaman = " & rsST!pjm & " ,Sewa = " & rsST!Sewa & "", vbInformation, "Info !!"
    Else
    MsgBox "Tidak Ada Pinjaman atau Sewa", vbInformation, "Info !!"
    End If



Exit Sub
hell:

MsgBox err.Description, vbCritical, "Error !!"
End Sub

Private Sub cmdInfo_KeyPress(KeyAscii As Integer)


If KeyAscii = 27 Then Unload Me
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

Unload AR_PObeli

sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from po_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdpo='" & txtkdPO & "' order by a.kdbarang"

Set rsX = con.Execute(sqlX)

With AR_PObeli.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_PObeli
.fldunit.DataField = "unit"
.fldnmbarang.DataField = "nmbarang"
.fldsatuan.DataField = "satuan"
.fldketerangan.DataField = "keterangan"

.lblnoPO = txtkdPO
.lblnmgudang = lblnmcustomer
.lbltglPO = Format(txttglPO, "dd/MM/yyyy")

.lbljudul = "PO PERMINTAAN BARANG"
.lbljudul1 = "CUSTOMER : "
.lblkategori = cmbkategori.Text

If txtketerangan = "" Then
.lblNB = ""
Else
.lblNB = "NB : " & txtketerangan
End If


AR_PObeli.Show vbModal

End With

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

If cmbkategori.Text = "FREE" Then
sqlA = "select a.kdbarang,sum(a.unit) as UKeluar,b.kdpo from FREE_d a left join free b  on a.kdFree =b.kdfree where b.kdpo ='" & txtkdPO & "' group by a.kdbarang,b.kdpo"
ElseIf cmbkategori.Text = "PINJAM PAKAI" Then
sqlA = "select a.kdbarang,sum(a.unit) as UKeluar,b.kdpo from pinjam_d a left join pinjam b  on a.kdpinjam =b.kdpinjam where b.kdpo ='" & txtkdPO & "' group by a.kdbarang,b.kdpo"
ElseIf cmbkategori.Text = "SEWA" Then
sqlA = "select a.kdbarang,sum(a.unit) as UKeluar,b.kdpo from sewa_d a left join sewa b  on a.kdsewa =b.kdsewa where b.kdpo ='" & txtkdPO & "' group by a.kdbarang,b.kdpo"
Else
sqlA = "select a.kdbarang,sum(a.unit) as UKeluar,b.kdpo from perbaikan_d a left join perbaikan b  on a.kdperbaikan =b.kdperbaikan where b.kdpo ='" & txtkdPO & "' group by a.kdbarang,b.kdpo"
End If



sql1 = "select a.kdbarang,b.kd1,b.nmbarang,a.unit,isnull(c.Ukeluar,0) as Ukeluar,b.satuan,a.keterangan,a.kdpo_d from po_d a left join barang b " & vbCrLf & _
      "on a.kdbarang=b.kdbarang left join (" & sqlA & ") c on a.kdPO=c.kdPO and a.kdbarang=c.kdbarang where a.kdpo='" & txtkdPO & "' "
      
sql = "select kdbarang,kd1,nmbarang,unit,ukeluar,unit - ukeluar as sisa,satuan,keterangan,kdPO_D from (" & sql1 & ") a  order by kdbarang"
      
      
      
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs

sqlB = "select kdbarang ,ukeluar from (" & sqlA & ") a "
Set rsB = con.Execute(sqlB)

If rsB.RecordCount <> 0 Then
cmbkategori.Enabled = False
cmdBR1.Enabled = False
Else
cmbkategori.Enabled = True
cmdBR1.Enabled = True
End If


Call LG

MousePointer = vbDefault
End Sub



Private Sub tbh()
Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat diUpdate, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else
    If txttglPO.Enabled = False Then
    PO_DTU.LBLKODE = 1
    PO_DTU.Show vbModal
    Else
    MsgBox "Kepala data belum disimpan !", vbCritical, "INfo !!"
    End If
End If
End Sub


Private Sub ubh()
Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat diUpdate, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else
    

    PO_DTU.LBLKODE = 2
    
    
    lblpos = rs.AbsolutePosition
    kode = 2
    
    PO_DTU.lblkdPO_d = rs!kdPO_d
    
    PO_DTU.lblkdbarang = rs!kdbarang
    PO_DTU.lblnmbarang = rs!nmbarang
    PO_DTU.lblsatuan = rs!satuan
    PO_DTU.txtunit = FormatNumber(rs!unit, 0)
    PO_DTU.lblunit_awal = FormatNumber(rs!unit, 0)
    PO_DTU.txtketerangan = rs!keterangan
    PO_DTU.cmdBR.Enabled = False
    
      
    PO_DTU.Show vbModal
End If
End Sub


Private Sub hps()
On Error GoTo hell

Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat diUpdate, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from PO_d where kdpo_d ='" & rs!kdPO_d & "'"
        con.Execute (sql)
        TimerALL.Interval = 10
    End If
    
End If
         

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub









Private Sub nomer()
On Error GoTo hell

If LBLKODE = 1 Then
    sql = "select isnull(max(right(kdpo,4)),0) as xx from PO where Month(tglPO)='" & Month(txttglPO) & "'  and year(tglPO)='" & Year(txttglPO) & "' and kdgudang= '" & lblkdgudang & "'"
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    txtkdPO = lblkdgudang & "/C/" & Format(txttglPO, "MMyy") & "/" & "000" & a
                Case 2
                    txtkdPO = lblkdgudang & "/C/" & Format(txttglPO, "MMyy") & "/" & "00" & a
                Case 3
                    txtkdPO = lblkdgudang & "/C/" & Format(txttglPO, "MMyy") & "/" & "0" & a
                Case 4
                    txtkdPO = lblkdgudang & "/C/" & Format(txttglPO, "MMyy") & "/" & a
        End Select
    
    Else
        txtnoPO = lblkdgudang & "/C/" & Format(txttglPO, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
txtnoPO = lblkdgudang & "/C/" & Format(txttglPO, "MMyy") & "/" & "0001"
End Sub




Private Sub cmdBR_Click()
Gudang_BR.LBLKODE = "PO_D"
Gudang_BR.Show vbModal

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
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Hanya Meng-Update No EASAP saja ya Gaes, Data Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Info !"
    
    sql = "Update PO set noEASAP='" & UCase(txtnoEASAP) & "' where kdpo='" & txtkdPO & "'"
    con.Execute (sql)
    
    sql = "Update pinjam set noEASAP='" & UCase(txtnoEASAP) & "' where kdPO='" & txtkdPO & "'"
    con.Execute (sql)
    
    sql = "Update sewa set noEASAP='" & UCase(txtnoEASAP) & "' where kdPO='" & txtkdPO & "'"
    con.Execute (sql)
    
    sql = "Update free set noEASAP='" & UCase(txtnoEASAP) & "' where kdPO='" & txtkdPO & "'"
    con.Execute (sql)
        
'    sql = "Update perbaikan set noEASAP='" & UCase(txtnoEASAP) & "' where kdPO='" & txtkdPO & "'"
'    con.Execute (sql)
'
    
    txttglPO.Enabled = False
    cmdBR2.Enabled = False
    cmdBR1.Enabled = False
    cmdBR.Enabled = False
    cmbkategori.Enabled = False
    txtketerangan.Enabled = False
    cmdsimpan.Enabled = False
    txtnoEASAP.Enabled = False
    
    PO.TimerALL.Interval = 10
    
    Exit Sub
Else

    If txtkdPO = "" Or lblkdgudang = "" Or lblkdcustomer = "" Or lblkdbarang.BackColor <> vbWhite Then
    MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
    Exit Sub
    Else
    
    
    
        If LBLKODE = 1 Then
        Call nomer
        
        sql = "insert into PO values ('" & txtkdPO & "','" & Format(txttglPO, "yyyy-MM-dd") & "','" & lblkdgudang & "','" & lblkdcustomer & "','" & lblkdkategori & "','" & UCase(txtketerangan) & "','','" & lblkdbarang & "','" & UCase(txtnoEASAP) & "')"
        con.Execute (sql)
        
        txttglPO.Enabled = False
        cmdBR2.Enabled = False
        cmdBR1.Enabled = False
        cmdBR.Enabled = False
        cmbkategori.Enabled = False
        txtketerangan.Enabled = False
        cmdsimpan.Enabled = False
        txtnoEASAP.Enabled = False
        cmdT(0).SetFocus
        
        
        ElseIf LBLKODE = 2 Then
        sql = "Update PO set keterangan='" & UCase(txtketerangan) & "',kdkategori='" & lblkdkategori & "',kdcustomer='" & lblkdcustomer & "',kdbarang='" & lblkdbarang & "',noeasap='" & UCase(txtnoEASAP) & "' where kdpo='" & txtkdPO & "'"
        con.Execute (sql)
        
        txttglPO.Enabled = False
        cmdBR.Enabled = False
        txtketerangan.Enabled = False
        cmdsimpan.Enabled = False
        cmbkategori.Enabled = False
        cmdBR1.Enabled = False
        cmdBR2.Enabled = False
        txtnoEASAP.Enabled = False
        cmdT(0).SetFocus
        
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Header PO berhasil di Ubah ", vbInformation, "Info !"
        End If
     
    End If
     
    PO.TimerALL.Interval = 10
    PO_D.TimerALL.Interval = 10
    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttglPO = Date
txttglPO.Enabled = True


sql = "Select * from kategori order by kdkategori"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
cmbkategori.AddItem rs!nmkategori
rs.MoveNext
Loop

CMBket.AddItem "BARU"
CMBket.AddItem "PENGGANTIAN"
CMBket.AddItem "LAIN - LAIN"
CMBket.ListIndex = 0



TimerCMB.Interval = 10
TimerALL.Interval = 10
TimerNO.Interval = 10


Call nul(lblkdgudang)
Call nul(lblnmgudang)
Call nul(lblkdcustomer)
Call nul(lblnmcustomer)
Call nul(lblalamat)


End Sub



Private Sub Form_Unload(Cancel As Integer)
Call cek_dalem

If txttglPO.Enabled = False And rscek.RecordCount = 0 Then
 ms = MsgBox("Tidak Ada Detail PO, apa anda ingin membatalkan Header PO ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = " delete from PO where kdPO='" & txtkdPO & "' "
        con.Execute (sql)
           
        PO.TimerALL.Interval = 10
           
        Unload Me
        
    Else
        Cancel = 1
    End If
End If

End Sub

Private Sub lblalamat_Change()
Call nul(lblalamat)
End Sub

Private Sub lblkdbarang_Change()
If lblkdkategori = "04" Then
Call nul(lblkdbarang)
Else
lblkdbarang.BackColor = vbWhite
End If
End Sub

Private Sub lblkdcustomer_Change()
Call nul(lblkdcustomer)
End Sub

Private Sub lblkdgudang_Change()
Call nul(lblkdgudang)
Call nomer
End Sub



Private Sub lblnmbarang_Change()
If lblkdkategori = "04" Then
Call nul(lblnmbarang)
Else
lblnmbarang.BackColor = vbWhite
End If
End Sub

Private Sub lblnmcustomer_Change()
Call nul(lblnmcustomer)
End Sub

Private Sub lblnmgudang_Change()
Call nul(lblnmgudang)
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

If rs.RecordCount <> 0 Then
datagrid1.SetFocus
End If

 

TimerALL.Interval = 0
MousePointer = vbDefault

End Sub

Private Sub TimerCMB_Timer()
If LBLKODE = "1" Then
cmbkategori.ListIndex = 0
End If


TimerCMB.Interval = 0
End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If


TimerNO.Interval = 0
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

Private Sub txtnoeasap_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoeasap_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnoeasap_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtnoeasap_LostFocus()
txtnoEASAP = UCase(txtnoEASAP)
End Sub

Private Sub txttglPO_Change()
Call nul(txttglPO)
Call nomer

End Sub

Private Sub txttglPO_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglPO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txttglPO_KeyPress(KeyAscii As Integer)
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

Private Sub txttglPO_LostFocus()
On Error GoTo hell

txttglPO = FormatDateTime(txttglPO, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglPO.SetFocus

End Sub




