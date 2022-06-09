VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form TeknisiDalam_D 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerNO 
      Left            =   6480
      Top             =   135
   End
   Begin VB.TextBox txttglTD 
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
      Left            =   3555
      TabIndex        =   0
      Top             =   990
      Width           =   1590
   End
   Begin VB.TextBox txtkerusakan 
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
      Left            =   1260
      TabIndex        =   4
      Top             =   1710
      Width           =   11490
   End
   Begin VB.TextBox txtnoForm 
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
      Left            =   6300
      TabIndex        =   1
      Top             =   990
      Width           =   1680
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
   Begin Threed.SSCommand cmdsimpan 
      Height          =   780
      Left            =   13320
      TabIndex        =   5
      ToolTipText     =   "Simpan"
      Top             =   1305
      Width           =   735
      _ExtentX        =   1296
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
      Picture         =   "TeknisiDalam_D.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   0
      TabIndex        =   14
      Top             =   2295
      Width           =   13110
      _Version        =   524288
      _ExtentX        =   23125
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
      Left            =   13320
      TabIndex        =   7
      ToolTipText     =   "Tambah"
      Top             =   2610
      Width           =   735
      _ExtentX        =   1296
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
      Picture         =   "TeknisiDalam_D.frx":2A6D
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   1
      Left            =   13320
      TabIndex        =   8
      ToolTipText     =   "Ubah"
      Top             =   3420
      Width           =   735
      _ExtentX        =   1296
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
      Picture         =   "TeknisiDalam_D.frx":56E1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   2
      Left            =   13320
      TabIndex        =   9
      ToolTipText     =   "Hapus"
      Top             =   4230
      Width           =   735
      _ExtentX        =   1296
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
      Picture         =   "TeknisiDalam_D.frx":88DE
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   3
      Left            =   13320
      TabIndex        =   10
      ToolTipText     =   "Refresh"
      Top             =   5040
      Width           =   735
      _ExtentX        =   1296
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
      Picture         =   "TeknisiDalam_D.frx":B977
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   4
      Left            =   13320
      TabIndex        =   11
      ToolTipText     =   "Cetak"
      Top             =   5850
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
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
      Picture         =   "TeknisiDalam_D.frx":EAF3
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   630
      TabIndex        =   12
      Top             =   7065
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
      Picture         =   "TeknisiDalam_D.frx":12550
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand CmdBR 
      Height          =   420
      Left            =   12735
      TabIndex        =   3
      Top             =   1305
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
      Picture         =   "TeknisiDalam_D.frx":18DB2
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   3975
      Left            =   135
      TabIndex        =   6
      Top             =   2745
      Width           =   13020
      _cx             =   22966
      _cy             =   7011
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"TeknisiDalam_D.frx":1B5E4
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      Begin VB.Timer TimerAll 
         Left            =   4770
         Top             =   2025
      End
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   12735
      TabIndex        =   2
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
      Picture         =   "TeknisiDalam_D.frx":1B6B2
      ButtonStyle     =   4
   End
   Begin VB.Label Label6 
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
      Left            =   8235
      TabIndex        =   32
      Top             =   1035
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
      Left            =   8955
      TabIndex        =   31
      Top             =   990
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
      Left            =   9855
      TabIndex        =   30
      Top             =   990
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPART YG DIGUNAKAN :"
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
      Left            =   135
      TabIndex        =   29
      Top             =   2430
      Width           =   2445
   End
   Begin VB.Label lblmerk 
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
      Left            =   9855
      TabIndex        =   28
      Top             =   1350
      Width           =   2895
   End
   Begin VB.Label lblkdsap 
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
      Left            =   4725
      TabIndex        =   27
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Label lblkd1 
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
      Left            =   2520
      TabIndex        =   26
      Top             =   1350
      Width           =   2175
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3690
      TabIndex        =   25
      Top             =   8100
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
      Left            =   315
      TabIndex        =   24
      Top             =   1035
      Width           =   645
   End
   Begin VB.Label txtkdTD 
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
      Left            =   945
      TabIndex        =   23
      Top             =   990
      Width           =   1545
   End
   Begin VB.Label Label2 
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
      Left            =   2655
      TabIndex        =   22
      Top             =   1035
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Perbaikan"
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
      TabIndex        =   21
      Top             =   45
      Width           =   6000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KERUSAKAN :"
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
      TabIndex        =   20
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   285
      Left            =   6390
      TabIndex        =   19
      Top             =   7740
      Width           =   1050
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "BARANG :"
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
      TabIndex        =   18
      Top             =   1440
      Width           =   870
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
      Left            =   945
      TabIndex        =   17
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Label lblnmkategori 
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
      Left            =   6300
      TabIndex        =   16
      Top             =   1350
      Width           =   3525
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NO. FORM :"
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
      Left            =   5310
      TabIndex        =   15
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   7620
      Left            =   0
      Picture         =   "TeknisiDalam_D.frx":1DEE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14100
   End
End
Attribute VB_Name = "TeknisiDalam_D"
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
Dim i, j As Integer



Private Sub cek_dalem()
sqlcek = "select * from teknisidalam_d where kdTD='" & txtkdTD & "'"
Set rscek = con.Execute(sqlcek)
End Sub





Private Sub cmdBR_Click()
Barang_BR.LBLKODE = UCase("TEKNISIDALAM_D")
Barang_BR.Show vbModal
End Sub


Private Sub cmdBR1_Click()
Teknisi_BR.LBLKODE = "TEKNISIDALAM_D"
Teknisi_BR.Show vbModal

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

'Unload AR_PObeli
'
'sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from po_d a left join barang b " & vbCrLf & _
'       "on a.kdbarang=b.kdbarang where a.kdpo='" & txtkdPO & "' order by a.kdbarang"
'
'Set rsX = con.Execute(sqlX)
'
'With AR_PObeli.DC1
'.ConnectionString = koneksi
'.Source = sqlX
'End With
'
'With AR_PObeli
'.fldunit.DataField = "unit"
'.fldnmbarang.DataField = "nmbarang"
'.fldsatuan.DataField = "satuan"
'.fldketerangan.DataField = "keterangan"
'
'.lblnoPO = txtkdPO
'.lblnmgudang = lblnmcustomer
'.lbltglTD = Format(txttglTD, "dd/MM/yyyy")
'
'.lbljudul = "PO PERMINTAAN BARANG"
'.lbljudul1 = "CUSTOMER : "
'.lblkategori = cmbkategori.Text
'
'If txtkerusakan = "" Then
'.lblNB = ""
'Else
'.lblNB = "NB : " & txtkerusakan
'End If
'
'
'AR_PObeli.Show vbModal
'
'End With

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

sql = "select a.kdtd_D,a.kdsparepart,b.nmbarang,a.qty,b.satuan,a.keterangan from teknisiDalam_d a left join barang b on a.kdsparepart =b.kdbarang where a.kdTD ='" & txtkdTD & "' order by a.tglinput"
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs
Call LG

For i = 1 To (datagrid1.Rows - 1)

If rs.RecordCount <> 0 Then
datagrid1.TextMatrix(i, 0) = i
End If

Next

MousePointer = vbDefault
End Sub



Private Sub tbh()
Call Cek_tglOD
If CDate(txttglTD) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else
    If txttglTD.Enabled = False Then
    TeknisiDalam_DTU.LBLKODE = 1
    TeknisiDalam_DTU.lblform = "TEKNISIDALAM_D"
    
    
    TeknisiDalam_DTU.Show vbModal
    
    Else
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Kepala data belum disimpan !", vbCritical, "Info !!"
    End If
End If
End Sub


Private Sub ubh()
Call Cek_tglOD
If CDate(txttglTD) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

Else

    TeknisiDalam_DTU.LBLKODE = 2
    TeknisiDalam_DTU.lblform = "TEKNISIDALAM_D"
    lblpos = rs.AbsolutePosition
    kode = 2
    
    TeknisiDalam_DTU.lblkdTD_d = rs!kdTD_D
    
    TeknisiDalam_DTU.lblkdbarang = rs!kdsparepart
    TeknisiDalam_DTU.lblnmbarang = rs!nmbarang
    TeknisiDalam_DTU.lblsatuan = rs!satuan
    TeknisiDalam_DTU.txtunit = FormatNumber(rs!qty, 0)
    TeknisiDalam_DTU.txtketerangan = rs!keterangan
    TeknisiDalam_DTU.cmdBR.Enabled = False
          
    TeknisiDalam_DTU.Show vbModal
End If
End Sub


Private Sub hps()
On Error GoTo hell
Call Cek_tglOD
If CDate(txttglTD) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

Else

    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from teknisidalam_D where kdTD_d ='" & rs!kdTD_D & "'"
        con.Execute (sql)
        TimerALL.Interval = 10
    End If
    
End If
         

Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"
End Sub









Private Sub nomer()
On Error GoTo hell

If LBLKODE = 1 Then
    sql = "select isnull(max(right(kdTD,4)),0) as xx from teknisiDalam where Month(tglTD)='" & Month(txttglTD) & "'  and year(tglTD)='" & Year(txttglTD) & "' "
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    txtkdTD = "D/" & Format(txttglTD, "MMyy") & "/" & "000" & a
                Case 2
                    txtkdTD = "D/" & Format(txttglTD, "MMyy") & "/" & "00" & a
                Case 3
                    txtkdTD = "D/" & Format(txttglTD, "MMyy") & "/" & "0" & a
                Case 4
                    txtkdTD = "D/" & Format(txttglTD, "MMyy") & "/" & a
        End Select
    
    Else
        txtnoPO = "D/" & Format(txttglTD, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
txtkdTD = "D/" & Format(txttglTD, "MMyy") & "/" & "0001"
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
    If CDate(txttglTD) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
        MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
        Exit Sub

    ElseIf txtkdTD = "" Or lblkdbarang = "" Or txtkerusakan = "" Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
        Exit Sub
    Else
    
    
    
        If LBLKODE = 1 Then
        Call nomer
        
        sql = "insert into teknisiDALAM values ('" & txtkdTD & "','" & Format(txttglTD, "yyyy/MM/dd") & "','" & lblkdbarang & "','" & UCase(txtkerusakan) & "','" & UCase(txtnoForm) & "','" & lblkdteknisi & "')"
        con.Execute (sql)
        
        txttglTD.Enabled = False
        cmdBR.Enabled = False
        cmdBR1.Enabled = False
        txtkerusakan.Enabled = False
        txtnoForm.Enabled = False
        cmdsimpan.Enabled = False
        cmdT(0).SetFocus
        
        
        ElseIf LBLKODE = 2 Then
        sql = "Update teknisidalam set kerusakan='" & UCase(txtkerusakan) & "',noform='" & Trim(UCase(txtnoForm)) & "',kdbarang='" & lblkdbarang & "',kdteknisi='" & lblkdteknisi & "' where kdTD='" & txtkdTD & "'"
        con.Execute (sql)
        
        cmdBR.Enabled = False
        cmdBR1.Enabled = False
        txtkerusakan.Enabled = False
        cmdsimpan.Enabled = False
        txtnoForm.Enabled = False
        cmdT(0).SetFocus
        
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Header berhasil di Ubah ", vbInformation, "Info !"
        End If
     
    End If
     
    teknisiDalam.TimerALL.Interval = 10
    TeknisiDalam_D.TimerALL.Interval = 10
    

End Sub




Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub datagrid1_DblClick()
Call ubh
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)


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


txttglTD = Date
txttglTD.Enabled = True


TimerALL.Interval = 10
TimerNO.Interval = 10

Call nul(lblkdbarang)
Call nul(lblnmkategori)
Call nul(lblkdteknisi)
Call nul(lblnmteknisi)
Call nul(txtkerusakan)

End Sub



Private Sub Form_Unload(Cancel As Integer)
Call cek_dalem

If txttglTD.Enabled = False And rscek.RecordCount = 0 Then
 ms = MsgBox("Tidak Ada Detail Perbaikan, apa anda ingin membatalkan ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = " delete from teknisidalam where kdTD='" & txtkdTD & "' "
        con.Execute (sql)
           
        teknisiDalam.TimerALL.Interval = 10
           
        Unload Me
        
    Else
        Cancel = 1
    End If
End If

End Sub



Private Sub lblnmbarang_Click()

End Sub

Private Sub lblkdbarang_Change()
Call nul(lblkdbarang)
End Sub

Private Sub lblkdteknisi_Change()
Call nul(lblkdteknisi)
End Sub

Private Sub lblnmkategori_Change()
Call nul(lblnmkategori)
End Sub


Private Sub lblnmteknisi_Change()
Call nul(lblnmteknisi)
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

 
MousePointer = vbDefault

TimerALL.Interval = 0

End Sub



Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If


TimerNO.Interval = 0
End Sub



Private Sub txtkerusakan_Change()
Call nul(txtkerusakan)
End Sub

Private Sub txtkerusakan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtkerusakan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtkerusakan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtkerusakan_LostFocus()
txtkerusakan = UCase(txtkerusakan)
End Sub



Private Sub txtnoform_Change()
Call nul(txtnoForm)
End Sub

Private Sub txtnoform_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtnoform_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
KeyAscii = 0
Beep
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If

End Sub

Private Sub txtnoform_LostFocus()
txtnoForm = UCase(txtnoForm)
End Sub

Private Sub txtTGLTD_Change()
Call nul(txttglTD)
Call nomer

End Sub

Private Sub txtTGLTD_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtTGLTD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtTGLTD_KeyPress(KeyAscii As Integer)
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

Private Sub txtTGLTD_LostFocus()
On Error GoTo hell

txttglTD = FormatDateTime(txttglTD, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglTD.SetFocus

End Sub






