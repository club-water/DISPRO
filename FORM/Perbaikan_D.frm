VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Perbaikan_D 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerNO 
      Left            =   7290
      Top             =   720
   End
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
   End
   Begin VB.TextBox txttglperbaikan 
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
      Top             =   1170
      Width           =   1590
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
      Left            =   8190
      TabIndex        =   4
      Top             =   1890
      Width           =   4335
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   13
      Top             =   720
      Width           =   13020
      _Version        =   524288
      _ExtentX        =   22966
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   12015
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   1125
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
      Picture         =   "Perbaikan_D.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   735
      Left            =   13320
      TabIndex        =   5
      ToolTipText     =   "Simpan"
      Top             =   1305
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
      Picture         =   "Perbaikan_D.frx":2832
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   180
      TabIndex        =   14
      Top             =   2790
      Width           =   13020
      _Version        =   524288
      _ExtentX        =   22966
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   0
      Left            =   9855
      TabIndex        =   15
      ToolTipText     =   "Tambah"
      Top             =   7650
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "Perbaikan_D.frx":529F
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   1
      Left            =   13320
      TabIndex        =   8
      ToolTipText     =   "Ubah"
      Top             =   2970
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
      Picture         =   "Perbaikan_D.frx":7F13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   2
      Left            =   13320
      TabIndex        =   9
      ToolTipText     =   "Hapus"
      Top             =   3735
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
      Picture         =   "Perbaikan_D.frx":B110
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   3
      Left            =   13320
      TabIndex        =   10
      ToolTipText     =   "Refresh"
      Top             =   4500
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
      Picture         =   "Perbaikan_D.frx":E1A9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   4
      Left            =   13320
      TabIndex        =   11
      ToolTipText     =   "Cetak"
      Top             =   5265
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
      Picture         =   "Perbaikan_D.frx":11325
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   12
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
      Picture         =   "Perbaikan_D.frx":14D82
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBatal 
      Height          =   735
      Left            =   13320
      TabIndex        =   6
      ToolTipText     =   "Batal"
      Top             =   2070
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
      Picture         =   "Perbaikan_D.frx":1B5E4
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   11160
      TabIndex        =   2
      ToolTipText     =   "Simpan"
      Top             =   1485
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
      Picture         =   "Perbaikan_D.frx":1E883
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR2 
      Height          =   420
      Left            =   6300
      TabIndex        =   3
      ToolTipText     =   "Simpan"
      Top             =   1845
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
      Picture         =   "Perbaikan_D.frx":210B5
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   5010
      Left            =   180
      TabIndex        =   7
      Top             =   2925
      Width           =   13065
      _cx             =   23045
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Perbaikan_D.frx":238E7
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
   Begin VB.Label lblkdkategori 
      Caption         =   "Label1"
      Height          =   330
      Left            =   10260
      TabIndex        =   39
      Top             =   315
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "UNTUK DISPENCER / SHOWCASE :"
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
      TabIndex        =   38
      Top             =   2295
      Width           =   2760
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
      Left            =   3060
      TabIndex        =   37
      Top             =   2250
      Width           =   1500
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
      Left            =   4590
      TabIndex        =   36
      Top             =   2250
      Width           =   4020
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   420
      Left            =   5805
      TabIndex        =   35
      Top             =   8955
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TEKNISI :"
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
      TabIndex        =   34
      Top             =   1575
      Width           =   735
   End
   Begin VB.Label lblkdteknisi 
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
      TabIndex        =   33
      Top             =   1530
      Width           =   870
   End
   Begin VB.Label lblnmteknisi 
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
      TabIndex        =   32
      Top             =   1530
      Width           =   2940
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3690
      TabIndex        =   31
      Top             =   8775
      Visible         =   0   'False
      Width           =   1545
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
      TabIndex        =   30
      Top             =   1215
      Width           =   645
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
      TabIndex        =   29
      Top             =   1170
      Width           =   2175
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
      TabIndex        =   28
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label lbljudul 
      BackStyle       =   0  'Transparent
      Caption         =   "Perbaikan"
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
      Left            =   990
      TabIndex        =   27
      Top             =   45
      Width           =   8025
   End
   Begin VB.Label lblkdgudang1 
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
      TabIndex        =   26
      Top             =   1530
      Width           =   1140
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
      Left            =   225
      TabIndex        =   25
      Top             =   1575
      Width           =   825
   End
   Begin VB.Label lblnmgudang1 
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
      TabIndex        =   24
      Top             =   1530
      Width           =   4065
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
      Left            =   7020
      TabIndex        =   23
      Top             =   1935
      Width           =   1320
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
      Left            =   10440
      TabIndex        =   22
      Top             =   1170
      Width           =   1590
   End
   Begin VB.Label lblKDPerbaikan 
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
      TabIndex        =   21
      Top             =   1170
      Width           =   2175
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
      TabIndex        =   20
      Top             =   1215
      Width           =   735
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
      TabIndex        =   19
      Top             =   1215
      Width           =   870
   End
   Begin VB.Label lblnmgudang2 
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
      TabIndex        =   18
      Top             =   1890
      Width           =   4065
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "MSK GDG :"
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
      Left            =   180
      TabIndex        =   17
      Top             =   1935
      Width           =   1050
   End
   Begin VB.Label lblkdgudang2 
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
      TabIndex        =   16
      Top             =   1890
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "Perbaikan_D.frx":23A21
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14100
   End
End
Attribute VB_Name = "Perbaikan_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK, rsT As ADODB.Recordset
Dim a As Integer
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim rsACC As ADODB.Recordset
Dim color As Long, flag As Byte
Dim rsTD As ADODB.Recordset
Dim rscek As ADODB.Recordset




Private Sub cek_dalem()
sqlcek = "select * from Perbaikan_d where kdPerbaikan='" & lblKDPerbaikan & "'"
Set rscek = con.Execute(sqlcek)
End Sub


Private Sub cek_teknisiDalam()
sqlTD = "select a.kdbarang,b.kd1,b.kdsap,c.nmkategori,a.unit from PO_d a left join barang b on a.kdbarang=b.kdbarang left join kategoribrg c on b.kdkategori=c.kdkategori " & vbCrLf & _
        "where a.kdpo='" & txtkdPO & "' and a.kdbarang not in (select kdbarang from teknisidalam where tglTD='" & Format(txttglperbaikan, "yyyy/MM/dd") & "')"
Set rsTD = con.Execute(sqlTD)

End Sub



Private Sub cmdBatal_Click()
On Error GoTo hell

Call Cek_tglOD
If CDate(txttglperbaikan) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

     ms = MsgBox("Apakah anda ingin Membatalkan Pengeluaran Barang ini ?", vbYesNo + vbQuestion, "Info")
     If ms = vbYes Then
        sql = "update PO set kdkeluar='' where kdPO='" & txtkdPO & "'"
        con.Execute (sql)
        
        sql = "delete from perbaikan_d where kdperbaikan='" & lblKDPerbaikan & "'"
        con.Execute (sql)
        
        sql = "delete from perbaikan where kdperbaikan='" & lblKDPerbaikan & "'"
        con.Execute (sql)
        
        txtkdPO = ""
        txttglPO = ""
        cmdBR.Enabled = True
        cmdBR1.Enabled = True
        txttglperbaikan = Date
        txttglperbaikan.Enabled = True
        txtketerangan.Enabled = True
        
        
        lblkdteknisi = ""
        lblnmteknisi = ""
        lblkdgudang1 = ""
        lblnmgudang1 = ""
        lblkdgudang2 = ""
        lblnmgudang2 = ""
        
        txtketerangan = ""
        lblkode = 1
        
        
        TimerAll.Interval = 10
        Perbaikan.TimerAll.Interval = 10
    Else
        Exit Sub
    End If

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub cmdBatal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub cmdBR1_Click()
Teknisi_BR.lblkode = "PERBAIKAN_D"
Teknisi_BR.Show vbModal

End Sub

Private Sub cmdBR1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR2_Click()
Gudang_BR.lblkode = "PERBAIKAN_D"
Gudang_BR.Show vbModal

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

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Cetak()
'On Error GoTo hell
sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
                 "from RKP_stok where kdgudang='" & lblkdgudang1 & "' and tgl <= '" & Format(txttglperbaikan, "yyyy/MM/dd") & "' and kdbarang in (select kdbarang from perbaikan_d where kdperbaikan='" & lblKDPerbaikan & "') group by kdbarang"

sqlCS = "select * from (" & sqlCS1 & ") a where unit < 0 order by kdbarang"

Set rsCS = con.Execute(sqlCS)

If rsCS.RecordCount <> 0 Then
    Cancel = 1
    ms = MsgBox("Stok Barang Kurang, Tampilkan List Barang ?", vbCritical + vbYesNo, "Error !")
    If ms = vbYes Then
    List_Stok_selisih.lblkode = "PERBAIKAN"
    List_Stok_selisih.Show vbModal
    End If

Else


    Unload AR_SJ
    
    sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from perbaikan_d a left join barang b " & vbCrLf & _
           "on a.kdbarang=b.kdbarang where a.kdperbaikan='" & lblKDPerbaikan & "' order by a.kdbarang"
    
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
    
    .lblnosj = lblKDPerbaikan
    .lblnmcustomer = lblnmcustomer
    .lbltglSJ = Format(txttglperbaikan, "dd/MM/yyyy")
    .lblalamat = lblalamat
    
    If txtketerangan = "" Then
    .lblNB = ""
    Else
    .lblNB = "NB : " & txtketerangan
    End If
    
    sqlACC = "select * from Signature where kdFrm='SJ'"
    Set rsACC = con.Execute(sqlACC)
    
    .lblAcc1 = rsACC!Acc1
    .lblAcc2 = rsACC!Acc2
    .lblAcc3 = rsACC!Acc3
    .lblAcc4 = rsACC!Acc4
    
    
    
    AR_SJ.Show vbModal
    
    End With
    
End If

'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !"
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

sql = "select a.kdbarang,b.kd1,b.nmbarang,a.unit,b.satuan,a.harga,a.rupiah,a.keterangan,a.kdperbaikan_d,a.klik_hrg from perbaikan_d a left join barang b " & vbCrLf & _
      "on a.kdbarang=b.kdbarang where a.kdperbaikan='" & lblKDPerbaikan & "' order by a.kdbarang "
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


Call LG

MousePointer = vbDefault
End Sub



Private Sub tbh()


End Sub


Private Sub ubh()


Call Cek_tglOD
If CDate(txttglperbaikan) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    Perbaikan_DTU.lblkode = 2
    
    
    lblpos = rs.AbsolutePosition
    kode = 2
    
    
    Perbaikan_DTU.lblkdbarang = rs!kdbarang
    Perbaikan_DTU.lblnmbarang = rs!nmbarang
    Perbaikan_DTU.lblsatuan = rs!satuan
    Perbaikan_DTU.txtunit = FormatNumber(rs!unit, 0)
    Perbaikan_DTU.txtharga = FormatNumber(rs!harga, 0)
    Perbaikan_DTU.lblrupiah = FormatNumber(rs!rupiah, 0)
    Perbaikan_DTU.txtketerangan = rs!keterangan
    Perbaikan_DTU.lblkdperbaikan_d = rs!kdperbaikan_d
    Perbaikan_DTU.lblklik = rs!klik_hrg
    Perbaikan_DTU.lblunit_awal = rs!unit
    
    'Perbaikan_DTU.txtunit.Enabled = False
    
    If UTAMA.lblstatus = 1 Then
    Perbaikan_DTU.txtharga.Enabled = True
    Else
    Perbaikan_DTU.txtharga.Enabled = False
    End If
    
      
    Perbaikan_DTU.Show vbModal
     
End If
End Sub


Private Sub hps()
On Error GoTo hell
Call Cek_tglOD
If CDate(txttglperbaikan) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else


    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from perbaikan_d where kdperbaikan_d ='" & rs!kdperbaikan_d & "'"
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
    sql = "select isnull(max(right(kdperbaikan,4)),0) as xx from perbaikan where Month(tglperbaikan)='" & Month(txttglperbaikan) & "'  and year(tglperbaikan)='" & Year(txttglperbaikan) & "' and kdgudang1= '" & lblkdgudang1 & "'"
    Set rs = con.Execute(sql)
    
    
    a = CCur(rs!xx) + 1
    
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    lblKDPerbaikan = lblkdgudang1 & "/G/" & Format(txttglperbaikan, "MMyy") & "/" & "000" & a
                Case 2
                    lblKDPerbaikan = lblkdgudang1 & "/G/" & Format(txttglperbaikan, "MMyy") & "/" & "00" & a
                Case 3
                    lblKDPerbaikan = lblkdgudang1 & "/G/" & Format(txttglperbaikan, "MMyy") & "/" & "0" & a
                Case 4
                    lblKDPerbaikan = lblkdgudang1 & "/G/" & Format(txttglperbaikan, "MMyy") & "/" & a
        End Select
    
    Else
        lblKDPerbaikan = lblkdgudang1 & "/G/" & Format(txttglperbaikan, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
lblKDPerbaikan = lblkdgudang1 & "/G/" & Format(txttglperbaikan, "MMyy") & "/" & "0001"
End Sub




Private Sub cmdBR_Click()
PO_BR.lblkode = "PERBAIKAN_D"
PO_BR.lblkdkategori = "04"
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
If CDate(txttglperbaikan) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    If txtkdPO = "" Or lblkdgudang1 = "" Or lblkdteknisi = "" Or lblkdgudang2 = "" Then
    MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
    Exit Sub
    Else
    
        If lblkode = 1 Then
            Call nomer
            
            sql = "insert into perbaikan values ('" & lblKDPerbaikan & "','" & Format(txttglperbaikan, "yyyy-MM-dd") & "','" & lblkdgudang1 & "','" & lblkdgudang2 & "','" & lblkdteknisi & "','" & UCase(txtketerangan) & "','" & txtkdPO & "','" & lblkdkategori & "','" & lblkdbarang & "')"
            con.Execute (sql)
            
            sql = "insert into perbaikan_d select kdbarang  + '" & "_" & lblKDPerbaikan & "','" & lblKDPerbaikan & "',kdbarang,unit,0,0,keterangan,0 from PO_d where kdPO='" & txtkdPO & "'"
            con.Execute (sql)
            
            sql = "update PO set kdkeluar='" & lblKDPerbaikan & "' where kdpo ='" & txtkdPO & "'"
            con.Execute (sql)
            
            txttglperbaikan.Enabled = False
            cmdBR.Enabled = False
            cmdBR1.Enabled = False
            cmdBR2.Enabled = False
            txtketerangan.Enabled = False
            cmdsimpan.Enabled = False
            cmdBatal.Enabled = True
            
            
            Call cek_teknisiDalam
            If rsTD.RecordCount <> 0 And lblkdkategori = "05" And lblkdgudang1 = "GD2" And lblkdgudang2 = "GD1" Then
            List_TeknisiDalam.Show vbModal
            
            
            con.Execute ("delete from perbaikan_d where kdperbaikan= '" & lblKDPerbaikan & "' and kdbarang not in (select kdbarang from teknisidalam where tglTD='" & Format(txttglperbaikan, "yyyy/MM/dd") & "')")
            End If
            
            
        
        ElseIf lblkode = 2 Then
            sql = "Update perbaikan set keterangan='" & UCase(txtketerangan) & "',kdteknisi='" & lblkdteknisi & "',kdgudang2='" & lblkdgudang2 & "' where kdperbaikan='" & lblKDPerbaikan & "'"
            con.Execute (sql)
        
            txtketerangan.Enabled = False
            cmdsimpan.Enabled = False
            cmdBR1.Enabled = False
            cmdBR2.Enabled = False
            
        
            MsgBox "Header berhasil di Ubah ", vbInformation, "Info !"
        End If
     
    End If
     
    Perbaikan.TimerAll.Interval = 10
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

txttglperbaikan = Date
txttglperbaikan.Enabled = True

TimerAll.Interval = 10
TimerNO.Interval = 10


Call nul(lblkdgudang1)
Call nul(lblnmgudang1)
Call nul(txtkdPO)
Call nul(lbltglPO)
Call nul(lblkdgudang2)
Call nul(lblnmgudang2)
Call nul(lblkdteknisi)
Call nul(lblnmteknisi)


End Sub




Private Sub Form_Unload(Cancel As Integer)
If cmdBR.Enabled = False And UTAMA.lblstatus = 0 And lblkdbarang <> "" Then
sql2 = "select * from perbaikan_d where kdperbaikan='" & lblKDPerbaikan & "' and klik_hrg=0 "
Set rs2 = con.Execute(sql2)

    If rs2.RecordCount <> 0 Then
        MsgBox "tidak dapat keluar karena ada yg blom di klik tombol harganya !!", vbCritical, "Error !"
        Cancel = 1
        Exit Sub
    End If

End If


sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
                 "from RKP_stok where kdgudang='" & lblkdgudang1 & "' and tgl <= '" & Format(txttglperbaikan, "yyyy/MM/dd") & "' and kdbarang in (select kdbarang from perbaikan_d where kdperbaikan='" & lblKDPerbaikan & "') group by kdbarang"

sqlCS = "select * from (" & sqlCS1 & ") a where unit < 0 order by kdbarang"

Set rsCS = con.Execute(sqlCS)

If rsCS.RecordCount <> 0 Then
Cancel = 1
    ms = MsgBox("Stok Barang Kurang, Tampilkan List Barang ?", vbCritical + vbYesNo, "Error !")
    If ms = vbYes Then
    List_Stok_selisih.lblkode = "PERBAIKAN"
    List_Stok_selisih.Show vbModal
    End If
End If

Call cek_dalem

If txttglperbaikan.Enabled = False And rscek.RecordCount = 0 Then
 ms = MsgBox("Tidak Ada Detail Perbaikan, apa anda ingin membatalkan Header ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = " delete from perbaikan where kdperbaikan='" & lblKDPerbaikan & "' "
        con.Execute (sql)
                   
        sql = "update PO set kdkeluar='' where kdPO='" & txtkdPO & "'"
        con.Execute (sql)
                   
        Perbaikan.TimerAll.Interval = 10
           
        Unload Me
        
    Else
        Cancel = 1
    End If
End If



End Sub

Private Sub lblkdperbaikan_Change()
Call nul(lblKDPerbaikan)
End Sub

Private Sub lblkdgudang1_Change()
Call nul(lblkdgudang1)
Call nomer
End Sub

Private Sub lblkdgudang2_Change()
Call nul(lblkdgudang2)
End Sub

Private Sub lblkdteknisi_Change()
Call nul(lblkdteknisi)
End Sub

Private Sub lblnmgudang2_Change()
Call nul(lblnmgudang2)
End Sub

Private Sub lblnmgudang1_Change()
Call nul(lblnmgudang1)
End Sub



Private Sub lblnmteknisi_Change()
Call nul(lblnmteknisi)
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
MousePointer = vbDefault
End Sub

Private Sub TimerNO_Timer()
If lblkode = 1 Then
Call nomer
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



Private Sub txttglperbaikan_Change()
Call nul(txttglperbaikan)
Call nomer

End Sub

Private Sub txttglperbaikan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglperbaikan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglperbaikan_KeyPress(KeyAscii As Integer)
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

Private Sub txttglperbaikan_LostFocus()
On Error GoTo hell

txttglperbaikan = FormatDateTime(txttglperbaikan, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglperbaikan.SetFocus

End Sub













