VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form fixrute_TU 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerNR 
      Left            =   17010
      Top             =   270
   End
   Begin MSComCtl2.DTPicker DTPCari 
      Height          =   330
      Left            =   6300
      TabIndex        =   4
      Top             =   1485
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16761024
      CheckBox        =   -1  'True
      CustomFormat    =   "dd / MM / yyyy"
      Format          =   90374145
      CurrentDate     =   43923
   End
   Begin VB.TextBox txttglspk1 
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
      Left            =   9000
      TabIndex        =   2
      Text            =   "01/01/1900"
      Top             =   945
      Width           =   1320
   End
   Begin VB.TextBox txtcari 
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
      Left            =   1350
      TabIndex        =   3
      Top             =   1485
      Width           =   2490
   End
   Begin VB.TextBox txtperiode 
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
      Left            =   990
      TabIndex        =   0
      Top             =   945
      Width           =   1500
   End
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
   End
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   540
      TabIndex        =   19
      Top             =   720
      Width           =   18690
      _Version        =   524288
      _ExtentX        =   32967
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   270
      TabIndex        =   20
      Top             =   1440
      Width           =   18960
      _Version        =   524288
      _ExtentX        =   33443
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
      Left            =   19395
      TabIndex        =   6
      ToolTipText     =   "Tambah"
      Top             =   1845
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
      Picture         =   "fixrute_TU.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   19395
      TabIndex        =   7
      ToolTipText     =   "Hapus Per Customer"
      Top             =   2790
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   19395
      TabIndex        =   9
      ToolTipText     =   "Refresh"
      Top             =   4680
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":5D0D
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1350
      TabIndex        =   18
      Top             =   10935
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
      Picture         =   "fixrute_TU.frx":8E89
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   7155
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   900
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
      Picture         =   "fixrute_TU.frx":F6EB
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   5
      Left            =   19395
      TabIndex        =   8
      ToolTipText     =   "Hapus Semua"
      Top             =   3735
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":11F1D
      ButtonStyle     =   4
   End
   Begin Threed.SSOption O1 
      Height          =   330
      Left            =   12375
      TabIndex        =   14
      Top             =   1485
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Belum Dikunjungi"
   End
   Begin Threed.SSOption O2 
      Height          =   330
      Left            =   14265
      TabIndex        =   15
      Top             =   1485
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sudah Dikunjungi"
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   8655
      Left            =   225
      TabIndex        =   5
      Top             =   1845
      Width           =   19005
      _cx             =   33523
      _cy             =   15266
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744703
      ForeColorSel    =   8388608
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648384
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
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"fixrute_TU.frx":16500
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
      Begin VB.ComboBox DGKeterangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11790
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1215
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Timer Timerflood 
         Left            =   3105
         Top             =   2115
      End
      Begin VB.TextBox DGTGLPlan 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5805
         TabIndex        =   30
         Text            =   "dgtglplan"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1230
      End
      Begin C1SizerLibCtl.C1Elastic flood 
         Height          =   465
         Left            =   8235
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4905
         Visible         =   0   'False
         Width           =   4155
         _cx             =   7329
         _cy             =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   255
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   1
         FloodPercent    =   0
         CaptionPos      =   4
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
   End
   Begin Threed.SSOption O3 
      Height          =   330
      Left            =   16290
      TabIndex        =   16
      Top             =   1485
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Non Route "
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   19395
      TabIndex        =   11
      ToolTipText     =   "Split Kunjungan"
      Top             =   6570
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":1676E
      ButtonStyle     =   4
   End
   Begin Threed.SSOption O4 
      Height          =   330
      Left            =   17730
      TabIndex        =   17
      Top             =   1485
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Disegel"
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   6
      Left            =   19395
      TabIndex        =   12
      ToolTipText     =   "Realisasi Route Plan"
      Top             =   7515
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":1AE91
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   600
      Index           =   7
      Left            =   18135
      TabIndex        =   41
      ToolTipText     =   "Ada Outlet Non Route yg Belum dikunjungi"
      Top             =   90
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16744576
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureAnimationDelay=   240
      PictureFrames   =   4
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
      Picture         =   "fixrute_TU.frx":1F1DF
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   19395
      TabIndex        =   10
      ToolTipText     =   "Cetak Bentuk List"
      Top             =   5625
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":222F8
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   8
      Left            =   19395
      TabIndex        =   13
      ToolTipText     =   "Cek Omset"
      Top             =   8460
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "fixrute_TU.frx":2567E
      ButtonStyle     =   4
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000008&
      Caption         =   "TOTAL ROUTE PLAN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12825
      TabIndex        =   39
      Top             =   10575
      Width           =   1770
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   690
      Left            =   12690
      Top             =   10665
      Width           =   6135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "RAK GLN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   16965
      TabIndex        =   38
      Top             =   10890
      Width           =   825
   End
   Begin VB.Label lblTRG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   17820
      TabIndex        =   37
      Top             =   10845
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SHOWCASE :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   14895
      TabIndex        =   36
      Top             =   10890
      Width           =   1005
   End
   Begin VB.Label lblTSHOW 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   15930
      TabIndex        =   35
      Top             =   10845
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DISPENCER :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   12870
      TabIndex        =   34
      Top             =   10890
      Width           =   1005
   End
   Begin VB.Label lblTDISP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   13905
      TabIndex        =   33
      Top             =   10845
      Width           =   825
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Tanggal :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   4770
      TabIndex        =   32
      Top             =   1485
      Width           =   1500
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PER TGL :"
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
      Left            =   7740
      TabIndex        =   29
      Top             =   990
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   225
      TabIndex        =   28
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3870
      Picture         =   "fixrute_TU.frx":29CF5
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   420
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
      Left            =   4230
      TabIndex        =   27
      Top             =   945
      Width           =   2940
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
      Left            =   3330
      TabIndex        =   26
      Top             =   945
      Width           =   870
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CHEKER :"
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
      Left            =   2610
      TabIndex        =   25
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Route Plan"
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
      Left            =   1260
      TabIndex        =   24
      Top             =   45
      Width           =   6000
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ROUTE :"
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
      TabIndex        =   23
      Top             =   990
      Width           =   735
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   7425
      TabIndex        =   22
      Top             =   10890
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   6210
      TabIndex        =   21
      Top             =   10890
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   11490
      Left            =   0
      Picture         =   "fixrute_TU.frx":36BA5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20445
   End
End
Attribute VB_Name = "fixrute_TU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rsD As ADODB.Recordset
Dim rsC As ADODB.Recordset
Dim rsC1 As ADODB.Recordset
Dim rsT As ADODB.Recordset
Dim a, i As Integer
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim sqlA, sqlB, sqlC, sqlA1, sqlA2 As String
Dim color As Long, flag As Byte
Dim rsA As ADODB.Recordset
Dim rsB As ADODB.Recordset
Dim sql1, sqlK1 As String
Dim sqldel As String
Dim rsdel As ADODB.Recordset
Dim rsK As ADODB.Recordset
Dim rsNR As ADODB.Recordset

Private Sub cek_NR()
If txtperiode.Enabled = False Then
    sqlNRX = "select kdcustomer from route_plan where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi ='" & fixrute_TU.lblkdteknisi & "' union all" & vbCrLf & _
             "select kdcustomer from real_cek where nmrute='" & fixrute_TU.txtperiode & "' and kdcustomer not in (select kdcustomer from Route_plan  where nmrute='" & txtperiode & "' and kdteknisi ='" & lblkdteknisi & "')"
    
    sqlNR1 = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
                "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
                "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
                "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                    "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= getdate() group by kdcustomer,kdkategori" & vbCrLf & _
                ") a group by kdcustomer " & vbCrLf & _
           ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2+RG <>0"
    
    
    sqlNR2 = "select d.nmareaC,e.nmteknisi,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,a.disp,a.showC,a.RG from (" & sqlNR1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
             "left join  area_cheker d on b.kdareaC=d.kdareaC left join teknisi e on b.kdteknisi= e.kdteknisi where b.kdteknisi='" & fixrute_TU.lblkdteknisi & "'"
          
    sqlNR3 = "select kdcustomer,tglplan from plan_non_route where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi='" & fixrute_TU.lblkdteknisi & "'"
          
    sqlNR = "select b.tglplan,c.tglsj,a.nmareaC,a.nmteknisi,a.kdcustomer,a.nmcustomer,a.alamat,a.cp,a.telp,a.disp,a.showC,a.RG,a.disp+a.showC+a.RG as total from (" & sqlNR2 & ") a left join (" & sqlNR3 & ") b on a.kdcustomer=b.kdcustomer left join V_tglsj c on a.kdcustomer=c.kdcustomer where a.kdcustomer not in (" & sqlNRX & ") order by a.nmareaC,a.nmteknisi,a.nmcustomer,a.alamat"
    
    Set rsNR = con.Execute(sqlNR)
    
    If rsNR.RecordCount <> 0 Then
    cmdT(7).Visible = True
    Else
    cmdT(7).Visible = False
    End If
    
Else
cmdT(7).Visible = False
End If

End Sub



Private Sub cmdBR1_Click()
Teknisi_BR.lblkode = "FIXRUTE_TU"
Teknisi_BR.Show vbModal

End Sub

Private Sub cmdBR1_KeyPress(KeyAscii As Integer)
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


Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 2 Then
Call hps

ElseIf Index = 5 Then
Call hps_ALL
ElseIf Index = 3 Then

TimerALL.Interval = 10
ElseIf Index = 4 Then
    
    If cmdT(4).Enabled = True And rs!kode = "1" Then
    fixrute_S.lblidrute = rs!idrute
    fixrute_S.lblkdcustomer = rs!kdcustomer
    fixrute_S.Show vbModal
    Else
    MsgBox "Tidak Bisa di Split, karena data ini merupakan hasil Split Route Plan", vbCritical, "Error "
    End If

ElseIf Index = 1 Then

fixrute_list.TimerCetak.Interval = 10
fixrute_list.Show vbModal

ElseIf Index = 6 Then

    If txtperiode <> "" And txtperiode.Enabled = False Then
        Real_cek_TU.lblkode = 2
        Real_cek_TU.lblfrm = "FIXRUTE_TU"
        Real_cek_TU.txtperiode = txtperiode
        Real_cek_TU.lblkdteknisi = lblkdteknisi
        Real_cek_TU.lblnmteknisi = lblnmteknisi
        Real_cek_TU.Show vbModal
    End If

ElseIf Index = 7 Then

    If txtperiode <> "" And txtperiode.Enabled = False Then
        List_non_route.Show vbModal
    End If
ElseIf Index = 8 Then

sqlC1 = "select a.kdcustomer,a.kdsp + '/' + a.kdcustomer_IAP as kdcust_IAP,isnull(b.nmcustomer_iap,'-') as nmcustomer_IAP,isnull(alamat_iap,'-') as alamat_iap,isnull(c.nmsp,'-') as nmsp from customer a left join customer_IAP b " & vbCrLf & _
       "on a.kdsp + '/' + a.kdcustomer_iap = b.pk_cust_IAP left join sp_iap c on a.kdsp=c.kdsp where a.kdcustomer='" & rs!kdcustomer & "'"
Set rsC1 = con.Execute(sqlC1)

LIST_Omset_IAP.lblkdcustomer_IAP = rsC1!kdcust_IAP
LIST_Omset_IAP.lblnmCustomer_IAP = rsC1!nmcustomer_IAP
LIST_Omset_IAP.lblalamat_IAP = rsC1!alamat_IAP
LIST_Omset_IAP.lblnmsp = rsC1!nmsp
LIST_Omset_IAP.lblkdcustomer = rs!kdcustomer
LIST_Omset_IAP.Show vbModal

End If
End Sub

Private Sub datagrid1_DblClick()
On Error Resume Next

If datagrid1.Col = 4 And rs!kode = "1" And UTAMA.lblstatus = 1 And O1.Value = True Then

    If rs!tglplan < (CDate(Date) - 7) And UTAMA.lblstatus = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
        MsgBox "Tdk Bisa diubah, TGL Route plan maximal 7 hari dari tgl Skrg !!", vbCritical, "Error !"
        Exit Sub
    Else
    
        kode = 2
        lblpos = rs.AbsolutePosition
        
        DGTglplan.Top = datagrid1.Top + datagrid1.CellTop - 150
        DGTglplan.Left = datagrid1.Left + datagrid1.CellLeft
        
        DGTglplan = rs!tglplan
        
        DGTglplan.Visible = True
        DGTglplan.Height = datagrid1.CellHeight
        DGTglplan.Width = datagrid1.CellWidth
        DGTglplan.BackColor = vbYellow
        DGTglplan.SetFocus
        SendKeys "{Home}+{End}"
    End If

ElseIf datagrid1.Col = 5 And rs!kode = "1" And UTAMA.lblstatus = 1 And O2.Value = True Then

        
        kode = 2
        lblpos = rs.AbsolutePosition
        
        DGTglplan.Top = datagrid1.Top + datagrid1.CellTop - 150
        DGTglplan.Left = datagrid1.Left + datagrid1.CellLeft
        
        DGTglplan = rs!tglcek
        
        DGTglplan.Visible = True
        DGTglplan.Height = datagrid1.CellHeight
        DGTglplan.Width = datagrid1.CellWidth
        DGTglplan.BackColor = vbYellow
        DGTglplan.SetFocus
        SendKeys "{Home}+{End}"
    


ElseIf datagrid1.Col = 20 And rs!kode = "1" Then
kode = 2
lblpos = rs.AbsolutePosition

DGKeterangan.Top = datagrid1.Top + datagrid1.CellTop - 150
DGKeterangan.Left = datagrid1.Left + datagrid1.CellLeft


DGKeterangan.Text = rs!keterangan
DGKeterangan.Visible = True
DGKeterangan.Height = datagrid1.CellHeight
DGKeterangan.Width = datagrid1.CellWidth
DGKeterangan.SetFocus

End If

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
' flood.Visible = True
' Timerflood.Interval = 10
 TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("s") Or KeyAscii = Asc("S") Then
 
 If cmdT(4).Enabled = True And rs!kode = "1" Then
 fixrute_S.lblidrute = rs!idrute
 fixrute_S.lblkdcustomer = rs!kdcustomer
 fixrute_S.Show vbModal
 Else
 MsgBox "Tidak Bisa di Split, karena data ini merupakan hasil Split Route Plan", vbCritical, "Error "
 End If


End If
End Sub


Private Sub DGketerangan_KeyPress(KeyAscii As Integer)
On Error GoTo hell

If KeyAscii = 13 Then
    If O1.Value = True Then
    
        con.Execute ("update route_plan set keterangan='" & UCase(DGKeterangan.Text) & "' where idrute='" & rs!idrute & "'")
        DGKeterangan.Visible = False
        
        
        ms = InputBox("Input Detail keterangan !", "Detail Keterangan", rs!det_keterangan)
        
        con.Execute ("update route_plan set det_keterangan='" & Trim(UCase(ms)) & "' where idrute='" & rs!idrute & "'")
        
        TimerALL.Interval = 10
        
    Else
        con.Execute ("update real_cek set keterangan='" & UCase(DGKeterangan.Text) & "' where nmrute='" & txtperiode & "' and kdcustomer='" & rs!kdcustomer & "'")
        DGKeterangan.Visible = False
        
        
        ms = InputBox("Input Detail keterangan !", "Detail Keterangan", rs!det_keterangan)
        
        con.Execute ("update real_cek set det_keterangan='" & Trim(UCase(ms)) & "' where nmrute='" & txtperiode & "' and kdcustomer='" & rs!kdcustomer & "'")
        
        TimerALL.Interval = 10
    End If

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub DGketerangan_LostFocus()
DGKeterangan.Visible = False
End Sub

Private Sub DGTGLPlan_Change()
Call nul(DGTglplan)
End Sub

Private Sub DGTGLPlan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then


    If O1.Value = True Then
    con.Execute ("update route_plan set tglplan='" & Format(DGTglplan, "yyyy/MM/dd") & "' where idrute='" & rs!idrute & "'")
    DGTglplan.Visible = False
    
    ElseIf O2.Value = True Then
    con.Execute ("update real_cek set tglcek='" & Format(DGTglplan, "yyyy/MM/dd") & "' where idrute='" & rs!idrute & "' ")
    DGTglplan.Visible = False
    
    End If

    TimerALL.Interval = 10



End If
End Sub

Private Sub DGTGLPlan_LostFocus()
DGTglplan.Visible = False
End Sub


Private Sub DTPCari_Change()
TimerALL.Interval = 10
End Sub


Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub




Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub

Private Sub tbl()
sqlC = "select * from route_plan where nmrute='" & txtperiode & "' and kdteknisi='" & lblkdteknisi & "'"
Set rsC = con.Execute(sqlC)

If rsC.RecordCount = 0 And O1.Value = True Then
    txtperiode.Enabled = True
    cmdBR1.Enabled = True
    txttglspk1.Enabled = True
    cmdT(0).Enabled = True
    datagrid1.Enabled = False
    cmdT(4).Enabled = False
    
    
ElseIf rsC.RecordCount <> 0 And O1.Value = True Then
    txtperiode.Enabled = False
    cmdBR1.Enabled = False
    txttglspk1.Enabled = False
    cmdT(0).Enabled = True
    datagrid1.Enabled = True
    cmdT(4).Enabled = True
    
Else
  datagrid1.Enabled = True
  cmdT(4).Enabled = False
End If


If rs.RecordCount = 0 And O1.Value = True Then
cmdT(2).Enabled = False
cmdT(5).Enabled = False
ElseIf rs.RecordCount <> 0 And O1.Value = True Then
cmdT(2).Enabled = True
cmdT(5).Enabled = True
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

'planing
sqlQ = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC ,RG from ( " & vbCrLf & _
            "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
            "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
            "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl <= '" & Format(txttglspk1, "yyyy/MM/dd") & "' group by kdcustomer,kdkategori" & vbCrLf & _
            ") a group by kdcustomer " & vbCrLf & _
       ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2 + RG <>0"


'realisasi
sqlR1 = "select idrute,kdteknisi,nmrute,kdcustomer,keterangan,det_keterangan,min(tglcek) as tglcek from Real_Cek group by idrute,kdteknisi,nmrute,kdcustomer,keterangan,det_keterangan"

sqlR = "select a.*,isnull(b.tglcek,'1900/01/01') as tglcek,isnull(b.keterangan,'') as keterangan,isnull(b.det_keterangan,'') as det_keterangan from V_real_cek a left join (" & sqlR1 & ") b on a.idrute=b.idrute  where a.kdteknisi='" & fixrute_TU.lblkdteknisi & "' and a.nmrute='" & fixrute_TU.txtperiode & "'"



If O1.Value = True Then

    
    
    If TXTCARI = "" Then
    sql1 = "select '1' as kode,a.idrute,a.tglplan,d.tglcek,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,a.keterangan,a.det_keterangan from ROUTE_PLAN a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer  where a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "' "
           
    sql2 = "select '2' as kode,a.idrute_S as idrute,a.tglrute_S as tglplan,'1900/01/01' as tglcek,a.tglrute_S as tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,0 as jmlunit,0 as disp,0 as showC,0 as RG,0 as disp1,0 as showC1,0 as RG1,'' as keterangan,'' as det_keterangan from ROUTE_PLAN_S a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and a.kdcustomer not in (select kdcustomer from real_cek where kdteknisi='" & lblkdteknisi & "' and  nmrute= '" & txtperiode & "') "
            
    Else
    sql1 = "select '1' as kode,a.idrute,a.tglplan,d.tglcek,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,a.keterangan,a.det_keterangan from ROUTE_PLAN a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and (a.kdcustomer like '%" & TXTCARI & "%' or b.nmcustomer like '%" & TXTCARI & "%' or b.alamat like '%" & TXTCARI & "%' or b.CP like '%" & TXTCARI & "%' or b.telp like '%" & TXTCARI & "%' or a.keterangan like '%" & TXTCARI & "%' )  "
           
    sql2 = "select '2' as kode,a.idrute_S as idrute,a.tglrute_S as tglplan,'1900/01/01' as tglcek,a.tglrute_S as tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,0 as jmlunit,0 as disp,0 as showC,0 as RG,0 as disp1,0 as showC1,0 as RG1,'' as keterangan,'' as det_keterangan from ROUTE_PLAN_S a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and a.kdcustomer not in (select kdcustomer from real_cek where kdteknisi='" & lblkdteknisi & "' and  nmrute= '" & txtperiode & "') and (a.kdcustomer like '%" & TXTCARI & "%' or b.nmcustomer like '%" & TXTCARI & "%' or b.alamat like '%" & TXTCARI & "%' or b.CP like '%" & TXTCARI & "%' or b.telp like '%" & TXTCARI & "%')"
    End If
    
    sql = " select hari=(CASE WHEN DATENAME(dw, tglplan)='Sunday' then 'MING' WHEN DATENAME(dw, tglplan)='Monday' THEN 'SEN' WHEN DATENAME(dw, tglplan)='Tuesday' THEN 'SEL' WHEN DATENAME(dw, tglplan)='Wednesday' THEN 'RAB' WHEN DATENAME(dw, tglplan)='Thursday' THEN 'KAM' WHEN DATENAME(dw, tglplan)='Friday' THEN 'JUM' ELSE 'SAB' END )"
    
    If IsNull(DTPCari.Value) Then
    sql = sql & ", * from (" & sql1 & " union all " & sql2 & ") a  where disp1 + showC1 + RG1 = 0 order by a.tglplan,a.tglinput,a.nmcustomer ,a.alamat"
    
    sqlT = "select kode, sum(disp) as disp,sum(showC) as ShowC,sum(RG) as RG from (" & sql1 & ") a where disp1 + showC1 + RG1 = 0 group by kode"
    
    sqldel = "select max(tglplan) as tglplan from (" & sql1 & ") x where disp1 + showC1 + RG1 = 0"
    Else
    sql = sql & ", * from (" & sql1 & " union all " & sql2 & ") a  where disp1 + showC1 + RG1 = 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' order by a.tglplan,a.tglinput,a.nmcustomer ,a.alamat"
    
    sqlT = "select kode, sum(disp) as disp,sum(showC) as ShowC,sum(RG) as RG from (" & sql1 & ") a where disp1 + showC1 + RG1 = 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' group by kode"
    
    sqldel = "select max(tglplan) as tglplan from (" & sql1 & ") x where disp1 + showC1 + RG1 = 0 and x.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "'"
    End If
    
    
ElseIf O2.Value = True Then

    If TXTCARI = "" Then
    sql1 = "select '1' as kode,a.idrute,a.tglplan,d.tglcek,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,d.keterangan,d.det_keterangan from ROUTE_PLAN a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer where a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "'"
              
    Else
    sql1 = "select '1' as kode,a.idrute,a.tglplan,d.tglcek,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,d.keterangan,d.det_keterangan from ROUTE_PLAN a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and (a.kdcustomer like '%" & TXTCARI & "%' or b.nmcustomer like '%" & TXTCARI & "%' or b.alamat like '%" & TXTCARI & "%' or b.CP like '%" & TXTCARI & "%' or b.telp like '%" & TXTCARI & "%') "
           
    End If
    
    sql = " select hari=(CASE WHEN DATENAME(dw, tglplan)='Sunday' then 'MING' WHEN DATENAME(dw, tglplan)='Monday' THEN 'SEN' WHEN DATENAME(dw, tglplan)='Tuesday' THEN 'SEL' WHEN DATENAME(dw, tglplan)='Wednesday' THEN 'RAB' WHEN DATENAME(dw, tglplan)='Thursday' THEN 'KAM' WHEN DATENAME(dw, tglplan)='Friday' THEN 'JUM' ELSE 'SAB' END )"
    
    If IsNull(DTPCari.Value) Then
    sql = sql & ", * from (" & sql1 & ") a  where disp1 + showC1 + RG1<> 0 order by a.tglplan desc,a.tglinput,a.nmcustomer ,a.alamat"
    sqlT = "select kode, sum(disp1) as disp,sum(showC1) as ShowC,sum(RG1) as RG from (" & sql1 & ") a where disp1 + showC1 + RG1 <> 0 group by kode"
    Else
    sql = sql & ", * from (" & sql1 & ") a  where disp1 + showC1 + RG1<> 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' order by a.tglplan Desc,a.tglinput,a.nmcustomer ,a.alamat"
    sqlT = "select kode, sum(disp1) as disp,sum(showC1) as ShowC,sum(RG1) as RG from (" & sql1 & ") a where disp1 + showC1 + RG1 <> 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' group by kode"
    End If
    
ElseIf O3.Value = True Then
    sqlA1 = "select idrute from route_plan where nmrute='" & txtperiode & "' and kdteknisi='" & lblkdteknisi & "'"
    
    If TXTCARI = "" Then
    sql1 = "select '1' as kode,a.idrute,a.tglcek as tglplan,a.tglcek,a.tglinput,C.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,sum(unit) as jmlunit,0 as disp,0 as showC,0 as RG from" & vbCrLf & _
           "Real_Cek a left join customer b on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareac=c.kdareaC  where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and a.idrute not in (" & sqlA1 & ") group by" & vbCrLf & _
           "a.idrute , a.tglcek,a.tglinput, C.nmareaC, a.kdcustomer, b.nmcustomer,b.alamat, b.cp, b.telp"
    Else
    sql1 = "select '1' as kode,a.idrute,a.tglcek as tglplan,a.tglcek,a.tglinput,C.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,sum(unit) as jmlunit,0 as disp,0 as showC,0 as RG from" & vbCrLf & _
           "Real_Cek a left join customer b on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareac=c.kdareaC  where a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "' and a.idrute not in (" & sqlA1 & ") and (a.kdcustomer like '%" & TXTCARI & "%' or b.nmcustomer like '%" & TXTCARI & "%' or b.alamat like '%" & TXTCARI & "%' or b.CP like '%" & TXTCARI & "%' or b.telp like '%" & TXTCARI & "%') group by" & vbCrLf & _
           "a.idrute , a.tglcek,a.tglinput, C.nmareaC, a.kdcustomer, b.nmcustomer,b.alamat, b.cp, b.telp"
    
    End If
    
    sql = " select hari=(CASE WHEN DATENAME(dw, tglplan)='Sunday' then 'MING' WHEN DATENAME(dw, tglplan)='Monday' THEN 'SEN' WHEN DATENAME(dw, tglplan)='Tuesday' THEN 'SEL' WHEN DATENAME(dw, tglplan)='Wednesday' THEN 'RAB' WHEN DATENAME(dw, tglplan)='Thursday' THEN 'KAM' WHEN DATENAME(dw, tglplan)='Friday' THEN 'JUM' ELSE 'SAB' END )"
    
    If IsNull(DTPCari.Value) Then
    sql = sql & ",a.*,isnull(b.disp1,0) as disp1,isnull(b.showC1,0) as showC1,isnull(b.RG,0) as RG1,b.keterangan,b.det_keterangan from (" & sql1 & ") a left join (" & sqlR & ") b on a.idrute=b.idrute and a.kdcustomer=b.kdcustomer order by a.tglplan,a.tglinput,a.nmcustomer ,a.alamat"
    sqlT = "select a.kode, sum(b.disp1) as disp,sum(b.showC1) as ShowC,sum(b.RG) as RG from (" & sql1 & ") a left join (" & sqlR & ") b on a.idrute=b.idrute and a.kdcustomer=b.kdcustomer group by kode"
    Else
    sql = sql & ",a.*,isnull(b.disp1,0) as disp1,isnull(b.showC1,0) as showC1,isnull(b.RG,0) as RG1,b.keterangan,b.det_keterangan from (" & sql1 & ") a left join (" & sqlR & ") b on a.idrute=b.idrute and a.kdcustomer=b.kdcustomer where a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' order by a.tglplan desc,a.tglinput,a.nmcustomer ,a.alamat"
    sqlT = "select a.kode, sum(b.disp1) as disp,sum(b.showC1) as ShowC,sum(b.RG) as RG from (" & sql1 & ") a left join (" & sqlR & ") b on a.idrute=b.idrute and a.kdcustomer=b.kdcustomer where a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' group by kode"
    End If

'ElseIf O4.Value = True Then
'
'
'
'    If txtcari = "" Then
'    sql1 = "select '1' as kode,a.idrute,a.tglplan,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,a.keterangan from ROUTE_PLAN a left join Customer b " & vbCrLf & _
'           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer  where a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "' and b.disegel=1"
'
'    sql2 = "select '2' as kode,a.idrute_S as idrute,a.tglrute_S as tglplan,a.tglrute_S as tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,0 as jmlunit,0 as disp,0 as showC,0 as RG,0 as disp1,0 as showC1,0 as RG1,'' as keterangan from ROUTE_PLAN_S a left join Customer b " & vbCrLf & _
'           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and a.kdcustomer not in (select kdcustomer from real_cek where kdteknisi='" & lblkdteknisi & "' and  nmrute= '" & txtperiode & "') and b.disegel=1"
'
'    Else
'    sql1 = "select '1' as kode,a.idrute,a.tglplan,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,a.keterangan from ROUTE_PLAN a left join Customer b " & vbCrLf & _
'           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and (a.kdcustomer like '%" & txtcari & "%' or b.nmcustomer like '%" & txtcari & "%' or b.alamat like '%" & txtcari & "%' or b.CP like '%" & txtcari & "%' or b.telp like '%" & txtcari & "%')  and b.disegel=1"
'
'    sql2 = "select '2' as kode,a.idrute_S as idrute,a.tglrute_S as tglinput,a.tglrute_S as tglplan,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,0 as jmlunit,0 as disp,0 as showC,0 as RG,0 as disp1,0 as showC1,0 as RG1,'' as keterangan from ROUTE_PLAN_S a left join Customer b " & vbCrLf & _
'           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & txtperiode & "') and a.kdcustomer not in (select kdcustomer from real_cek where kdteknisi='" & lblkdteknisi & "' and  nmrute= '" & txtperiode & "') and (a.kdcustomer like '%" & txtcari & "%' or b.nmcustomer like '%" & txtcari & "%' or b.alamat like '%" & txtcari & "%' or b.CP like '%" & txtcari & "%' or b.telp like '%" & txtcari & "%') and b.disegel=1"
'    End If
'
'    sql = " select hari=(CASE WHEN DATENAME(dw, tglplan)='Sunday' then 'MING' WHEN DATENAME(dw, tglplan)='Monday' THEN 'SEN' WHEN DATENAME(dw, tglplan)='Tuesday' THEN 'SEL' WHEN DATENAME(dw, tglplan)='Wednesday' THEN 'RAB' WHEN DATENAME(dw, tglplan)='Thursday' THEN 'KAM' WHEN DATENAME(dw, tglplan)='Friday' THEN 'JUM' ELSE 'SAB' END )"
'
'    If IsNull(DTPCari.Value) Then
'    sql = sql & ", * from (" & sql1 & " union all " & sql2 & ") a  where disp1 + showC1 + RG1 = 0 order by a.tglplan,a.tglinput,a.nmcustomer ,a.alamat"
'
'    sqlT = "select kode, sum(disp) as disp,sum(showC) as ShowC,sum(RG) as RG from (" & sql1 & ") a where disp1 + showC1 + RG1 = 0 group by kode"
'
'    Else
'    sql = sql & ", * from (" & sql1 & " union all " & sql2 & ") a  where disp1 + showC1 + RG1 = 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' order by a.tglplan,a.tglinput,a.nmcustomer ,a.alamat"
'
'    sqlT = "select kode, sum(disp) as disp,sum(showC) as ShowC,sum(RG) as RG from (" & sql1 & ") a where disp1 + showC1 + RG1 = 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' group by kode"
'    End If
    
End If


Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

For i = 1 To (datagrid1.Rows - 1)
For j = 1 To (datagrid1.Cols - 1)

If rs.RecordCount <> 0 Then
datagrid1.TextMatrix(i, 0) = i
End If

If datagrid1.TextMatrix(i, 20) <> "" Then
datagrid1.Cell(flexcpForeColor, i, j) = vbRed
End If

If datagrid1.TextMatrix(i, 4) < datagrid1.TextMatrix(i, 5) And O2.Value = True Then
datagrid1.Cell(flexcpBackColor, i, j) = &HC0C0FF
ElseIf datagrid1.TextMatrix(i, 4) > datagrid1.TextMatrix(i, 5) And O2.Value = True Then
datagrid1.Cell(flexcpBackColor, i, j) = &HFF8080
End If

Next
Next



    'total unit
    
    Set rsT = con.Execute(sqlT)
    
    If rsT.RecordCount <> 0 Then
    lblTDISP = FormatNumber(rsT!disp, 0)
    lblTSHOW = FormatNumber(rsT!showC, 0)
    lblTRG = FormatNumber(rsT!RG, 0)
    Else
    lblTDISP = "0"
    lblTSHOW = "0"
    lblTRG = "0"
    End If



fixrute.TimerALL.Interval = 10





Call LG

MousePointer = vbDefault
End Sub



Private Sub tbh()
If O1.Value = True Then

  If txtperiode = "" Or lblkdteknisi = "" Or txttglspk1 = "" Then
        MsgBox "Header Belum Lengkap !!", vbCritical, "Error !"
        Exit Sub
    Else
    
    
        
        If IsNull(DTPCari.Value) Then
        Rute_Cheker_BR.txttglplan.Enabled = True
        Rute_Cheker_BR.txttglplan = Date
        Else
        Rute_Cheker_BR.txttglplan.Enabled = False
        Rute_Cheker_BR.txttglplan = DTPCari.Value
               
        End If
        
        Rute_Cheker_BR.lblkdteknisi = lblkdteknisi
        Rute_Cheker_BR.lblnmteknisi = lblnmteknisi
        Rute_Cheker_BR.lbltgl1 = Format(txttglspk1, "dd/MM/yyyy")
        Rute_Cheker_BR.Show vbModal
    End If
   
End If

End Sub


Private Sub ubh()
End Sub


Private Sub hps()
On Error GoTo hell

If O1.Value = True Then
    
'    If rs!tglplan < (CDate(Date) - 7) And UTAMA.lblstatus = 0 Then
'        SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
'        MsgBox "Tdk Bisa dihapus, TGL Route plan maximal 7 hari dari tgl Skrg !!", vbCritical, "Error !"
'        Exit Sub
'    Else

    If UTAMA.lblstatus = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
        MsgBox "Tdk Bisa dihapus selain Administrator !!", vbCritical, "Error !"
        Exit Sub
    Else
        kode = 2
        Call max
        
        
        ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
        If ms = vbYes Then
'            flood.Visible = True
'            Timerflood.Interval = 10
    
            
            sql = "delete from route_plan_S where idrute  ='" & rs!idrute & "' "
            con.Execute (sql)
             
            sql = "delete from route_plan where idrute  ='" & rs!idrute & "' "
            con.Execute (sql)
            
            
            TimerALL.Interval = 10
            fixrute.TimerALL.Interval = 10
        End If

    End If
    
ElseIf O2.Value = True Or O3.Value = True Then
    
        kode = 2
        Call max
        
        
        ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
        If ms = vbYes Then
    
            
            sql = "delete from real_cek where idrute  ='" & rs!idrute & "' "
            con.Execute (sql)
             
            TimerALL.Interval = 10
            fixrute.TimerALL.Interval = 10
        End If



End If
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub


Private Sub hps_ALL()
On Error GoTo hell

If O1.Value = True Then
    kode = 2
    Call max
    
    Set rsdel = con.Execute(sqldel)
    
'    If rsdel!tglplan < (CDate(Date) - 7) And UTAMA.lblstatus = 0 Then
'        SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
'        MsgBox "Tdk Bisa dihapus, TGL Route plan maximal 7 hari dari tgl Skrg !!", vbCritical, "Error !"
'        Exit Sub
'    Else

    If UTAMA.lblstatus = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
        MsgBox "Tdk Bisa dihapus selain Administrator !!", vbCritical, "Error !"
        Exit Sub
    Else


        ms = MsgBox("Apakah anda ingin menghapus Semua data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
        If ms = vbYes Then
        
              
'            flood.Visible = True
'            Timerflood.Interval = 10
                   
    
            If IsNull(DTPCari.Value) Then
            sql3 = "select * from (" & sql1 & " ) a  where disp1 + showC1 + RG1 = 0 "
            Else
            sql3 = "select * from (" & sql1 & " ) a  where disp1 + showC1 + RG1 = 0 and a.tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' "
            End If
        
        
            sql = "delete from route_plan_S where idrute in (select idrute from (" & sql3 & ") a)"
            con.Execute (sql)
        
            sql = "delete from route_plan where idrute in (select idrute from (" & sql3 & ") a)"
            con.Execute (sql)
            
            
            
            TimerALL.Interval = 10
            fixrute.TimerALL.Interval = 10
        End If
    End If

End If
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

Call nul(txtperiode)
Call nul(lblkdteknisi)
Call nul(lblnmteknisi)

txttglspk1 = Date

O1.Value = True

DTPCari.Value = Date
DTPCari.Value = Null

'sqlK = "Select * from alasan_cek where kebutuhan='F' order by nmalasan"
'Set rsK = con.Execute(sqlK)
'
'rsK.MoveFirst
'
'Do While Not rsK.EOF
'DGKeterangan.AddItem rsK!nmalasan
'rsK.MoveNext
'Loop

If UTAMA.lblstatus = 0 Then
cmdT(5).Enabled = False
Else
cmdT(5).Enabled = True
End If


DGKeterangan.ListIndex = 0

TimerNR.Interval = 5000
TimerALL.Interval = 10
End Sub

Private Sub lblkdteknisi_Change()
Call nul(lblkdteknisi)
End Sub

Private Sub lblnmteknisi_Change()
Call nul(lblnmteknisi)
End Sub

Private Sub O1_Click(Value As Integer)
DGKeterangan.Clear

sqlK = "Select * from alasan_cek where kebutuhan='F' order by nmalasan"
Set rsK = con.Execute(sqlK)

rsK.MoveFirst

Do While Not rsK.EOF
DGKeterangan.AddItem rsK!nmalasan
rsK.MoveNext
Loop

Label8 = "TOTAL ROUTE PLAN"

TimerALL.Interval = 10

cmdT(0).Enabled = True
cmdT(2).Enabled = True
cmdT(5).Enabled = True

End Sub

Private Sub O1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub O2_Click(Value As Integer)
DGKeterangan.Clear

sqlK = "Select * from alasan_cek where kebutuhan='R' order by nmalasan"
Set rsK = con.Execute(sqlK)

rsK.MoveFirst

Do While Not rsK.EOF
DGKeterangan.AddItem rsK!nmalasan
rsK.MoveNext
Loop

Label8 = "TOTAL REALISASI"

TimerALL.Interval = 10
cmdT(0).Enabled = False
cmdT(2).Enabled = True
cmdT(5).Enabled = False

End Sub

Private Sub O2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub O3_Click(Value As Integer)
DGKeterangan.Clear

sqlK = "Select * from alasan_cek where kebutuhan='R' order by nmalasan"
Set rsK = con.Execute(sqlK)

rsK.MoveFirst

Do While Not rsK.EOF
DGKeterangan.AddItem rsK!nmalasan
rsK.MoveNext
Loop


TimerALL.Interval = 10
cmdT(0).Enabled = False
cmdT(2).Enabled = True
cmdT(5).Enabled = False

End Sub

Private Sub O3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub O4_Click(Value As Integer)
TimerALL.Interval = 10
cmdT(0).Enabled = True
cmdT(2).Enabled = True
cmdT(5).Enabled = True

End Sub

Private Sub O4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub


Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerALL.Interval = 0


MousePointer = vbDefault

End Sub


Private Sub Timerflood_Timer()
' Dim j%
'  Static i%
'
'  If i > 90 Then
'  i = 0
'
'  End If
'
'  i = i + 10
'
'  If i = 100 Then
'  Timerflood.Interval = 0
'  flood.Visible = False
'
'
''    If rs.RecordCount = 0 Then
''        SetTimer hwnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
''         MsgBox "Data gak ada Broo !", vbInformation, "Info !"
''    End If
'
'  End If
'
'
'  For j = 0 To 10
'    flood.FloodPercent = i
'    flood.Caption = i & "%"
'  Next j
'

End Sub

Private Sub TimerNR_Timer()
On Error Resume Next

Call cek_NR


End Sub

Private Sub TXTCARI_Change()
If TXTCARI = "" Then
TimerALL.Interval = 0
End If
End Sub

Private Sub TXTCARI_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTCARI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

TimerALL.Interval = 10
'    If rs.RecordCount <> 0 Then
'    datagrid1.SetFocus
'    Call LG
''    Else
''    CMBCARI.SetFocus
'    End If

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub


Private Sub txtperiode_Change()
Call nul(txtperiode)
End Sub

Private Sub txtperiode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtperiode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
Beep
KeyAscii = 0
End If
End Sub

Private Sub txtperiode_LostFocus()
txtperiode = UCase(txtperiode)
End Sub

Private Sub txttglspk1_Change()
Call nul(txttglspk1)
End Sub

Private Sub txttglSPK1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglSPK1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglspk1_KeyPress(KeyAscii As Integer)
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

Private Sub txttglspk1_LostFocus()
On Error GoTo hell

txttglspk1 = FormatDateTime(txttglspk1, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglspk1.SetFocus

End Sub


