VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Beli_D 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   15630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSJ 
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
      Left            =   1035
      TabIndex        =   2
      Top             =   1395
      Width           =   2175
   End
   Begin VB.TextBox txttglSJ 
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
      Left            =   4050
      TabIndex        =   3
      Top             =   1395
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
      Left            =   8100
      TabIndex        =   5
      Top             =   1755
      Width           =   4335
   End
   Begin VB.TextBox txttglBPB 
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
      Left            =   10395
      TabIndex        =   1
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
      Left            =   7335
      Top             =   585
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   8
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
      Left            =   5625
      TabIndex        =   0
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
      Picture         =   "Beli_D.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   735
      Left            =   14760
      TabIndex        =   6
      ToolTipText     =   "Simpan"
      Top             =   1215
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
      Picture         =   "Beli_D.frx":2832
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   9
      Top             =   2655
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
      Height          =   915
      Index           =   0
      Left            =   9855
      TabIndex        =   10
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
      Picture         =   "Beli_D.frx":529F
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   1
      Left            =   14760
      TabIndex        =   11
      ToolTipText     =   "Ubah"
      Top             =   2925
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
      Picture         =   "Beli_D.frx":7F13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   2
      Left            =   14760
      TabIndex        =   12
      ToolTipText     =   "Hapus"
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
      Picture         =   "Beli_D.frx":B110
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   3
      Left            =   14760
      TabIndex        =   13
      ToolTipText     =   "Refresh"
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
      Picture         =   "Beli_D.frx":E1A9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   735
      Index           =   4
      Left            =   14760
      TabIndex        =   14
      ToolTipText     =   "Cetak"
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
      Picture         =   "Beli_D.frx":11325
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
      Picture         =   "Beli_D.frx":14D82
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   6255
      TabIndex        =   4
      ToolTipText     =   "Simpan"
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
      Picture         =   "Beli_D.frx":1B5E4
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBatal 
      Height          =   735
      Left            =   14760
      TabIndex        =   7
      ToolTipText     =   "Batal"
      Top             =   1980
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
      Picture         =   "Beli_D.frx":1DE16
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   5010
      Left            =   180
      TabIndex        =   15
      Top             =   2925
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
      AllowUserResizing=   3
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
      FormatString    =   $"Beli_D.frx":210B5
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
      Left            =   1035
      TabIndex        =   37
      Top             =   2115
      Width           =   2040
   End
   Begin VB.Label Label5 
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
      Left            =   180
      TabIndex        =   36
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   285
      Left            =   5985
      TabIndex        =   35
      Top             =   8865
      Width           =   870
   End
   Begin VB.Label lblkdsupplier 
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
      Left            =   1035
      TabIndex        =   34
      Top             =   1755
      Width           =   1140
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER :"
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
      TabIndex        =   33
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label lblnmsupplier 
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
      Left            =   2205
      TabIndex        =   32
      Top             =   1755
      Width           =   4065
   End
   Begin VB.Label Label10 
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
      Left            =   315
      TabIndex        =   31
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL SJ :"
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
      TabIndex        =   30
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL BPB :"
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
      Left            =   9585
      TabIndex        =   29
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NO BPB :"
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
      Left            =   6390
      TabIndex        =   28
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblnoBPB 
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
      Left            =   7200
      TabIndex        =   27
      Top             =   1035
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
      Left            =   4050
      TabIndex        =   26
      Top             =   1035
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
      Left            =   6930
      TabIndex        =   25
      Top             =   1800
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
      Left            =   8370
      TabIndex        =   24
      Top             =   1395
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
      Left            =   6345
      TabIndex        =   23
      Top             =   1440
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
      Left            =   7200
      TabIndex        =   22
      Top             =   1395
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembelian Barang"
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
      Left            =   3330
      TabIndex        =   20
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
      Left            =   1035
      TabIndex        =   19
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
      Left            =   315
      TabIndex        =   18
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3690
      TabIndex        =   17
      Top             =   8775
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "Beli_D.frx":211E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15585
   End
End
Attribute VB_Name = "Beli_D"
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
Dim sqlACC As String
Dim color As Long, flag As Byte

Private Sub cmdBatal_Click()
On Error GoTo hell

 ms = MsgBox("Apakah anda ingin Membatalkan Pembelian ini ?", vbYesNo + vbQuestion, "Info")
 If ms = vbYes Then
        
    sql = "delete from beli_d where kdbeli='" & txtkdPO & "_" & lblnoBPB & "'"
    con.Execute (sql)
    
    sql = "delete from beli where kdbeli='" & txtkdPO & "_" & lblnoBPB & "'"
    con.Execute (sql)
    
    txtkdPO = ""
    txttglPO = ""
    cmdBR.Enabled = True
    txttglBPB = Date
    txttglBPB.Enabled = True
    txtSJ.Enabled = True
    txttglSJ.Enabled = True
    txtSJ = ""
    txttglSJ = Date
    lblkdgudang = ""
    lblnmgudang = ""
    lblkdsupplier = ""
    lblnmsupplier = ""
    cmdBR1.Enabled = True
    txtketerangan.Enabled = True
    txtketerangan = ""
    LBLKODE = 1
    
    
    TimerALL.Interval = 10
    Beli.TimerALL.Interval = 10
Else
    Exit Sub
End If


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub cmdBatal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR1_Click()
Supplier_BR.LBLKODE = "BELI_D"
Supplier_BR.Show vbModal

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

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Cetak()

Unload AR_LPB

sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from beli_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdbeli='" & txtkdPO & "_" & lblnoBPB & "' order by a.kdbarang"

Set rsX = con.Execute(sqlX)

With AR_LPB.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_LPB
.fldunit.DataField = "unit"
.fldnmbarang.DataField = "nmbarang"
.fldsatuan.DataField = "satuan"
.fldketerangan.DataField = "keterangan"
.fldkdbarang.DataField = "kdbarang"

.lblnoLPB = lblnoBPB
.lblsupplier = lblnmsupplier
.lbltglLPB = Format(txttglBPB, "dd/MM/yyyy")
.lblKET = "Note : " & txtketerangan

sqlACC = "select * from Signature where kdFrm='" & lblkdgudang & "'"
Set rsACC = con.Execute(sqlACC)

.lblAcc1 = rsACC!Acc1
.lblAcc4 = rsACC!Acc4



AR_LPB.Show vbModal

End With

End Sub


Private Sub Cetak1()

Unload AR_LPB

sqlX = "select a.kdbarang,b.kd1,b.nmbarang,abs(a.unit) as unit,b.satuan,a.keterangan from beli_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdbeli='" & txtkdPO & "_" & lblnoBPB & "' order by a.kdbarang"

Set rsX = con.Execute(sqlX)

With AR_LPB.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_LPB
.fldunit.DataField = "unit"
.fldnmbarang.DataField = "nmbarang"
.fldsatuan.DataField = "satuan"
.fldketerangan.DataField = "kd1"
.fldkdbarang.DataField = "kdbarang"
.lbljudul = "BUKTI PENGELUARAN BARANG"

.lblnoLPB = lblnoBPB
.lblsupplier = lblnmsupplier
.lbltglLPB = Format(txttglBPB, "dd/MM/yyyy")


sqlACC = "select * from Signature where kdFrm='" & lblkdgudang & "'"
Set rsACC = con.Execute(sqlACC)

.lblmengetahui = "Penerima, "
.lblAcc1 = "(                                 )"
.lblAcc4 = rsACC!Acc4



AR_LPB.Show vbModal

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
sql = "select a.kdbarang,b.kd1,b.nmbarang,a.unit,b.satuan,a.harga,a.rupiah,a.keterangan,a.kdbeli_d from beli_d a left join barang b " & vbCrLf & _
      "on a.kdbarang=b.kdbarang where a.kdbeli='" & txtkdPO & "_" & lblnoBPB & "' order by a.kdbarang "
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


Call LG
End Sub



Private Sub tbh()


End Sub


Private Sub ubh()
Call Cek_tglOD
If CDate(txttglBPB) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

Else

    Beli_DTU.LBLKODE = 2
    
    
    lblpos = rs.AbsolutePosition
    kode = 2
    
    
    Beli_DTU.lblkdbarang = rs!kdbarang
    Beli_DTU.lblnmbarang = rs!nmbarang
    Beli_DTU.lblsatuan = rs!satuan
    Beli_DTU.txtunit = FormatNumber(rs!unit, 0)
    Beli_DTU.txtharga = FormatNumber(rs!harga, 0)
    Beli_DTU.lblrupiah = FormatNumber(rs!rupiah, 0)
    Beli_DTU.txtketerangan = rs!keterangan
    Beli_DTU.lblkdbeli_d = rs!kdbeli_d
    Beli_DTU.lblunit_awal = rs!unit
      
    Beli_DTU.Show vbModal
End If
 
End Sub


Private Sub hps()
On Error GoTo hell

Call Cek_tglOD
If CDate(txttglBPB) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

Else


    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from beli_d where kdbeli_d ='" & rs!kdbeli_d & "'"
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
    sql = "select isnull(max(right(kdbeli,4)),0) as xx from beli where Month(tglbpb)='" & Month(txttglBPB) & "'  and year(tglbpb)='" & Year(txttglBPB) & "' and kdgudang= '" & lblkdgudang & "'"
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    lblnoBPB = lblkdgudang & "/B/" & Format(txttglBPB, "MMyy") & "/" & "000" & a
                Case 2
                    lblnoBPB = lblkdgudang & "/B/" & Format(txttglBPB, "MMyy") & "/" & "00" & a
                Case 3
                    lblnoBPB = lblkdgudang & "/B/" & Format(txttglBPB, "MMyy") & "/" & "0" & a
                Case 4
                    lblnoBPB = lblkdgudang & "/B/" & Format(txttglBPB, "MMyy") & "/" & a
        End Select
    
    Else
        lblnoBPB = lblkdgudang & "/B/" & Format(txttglBPB, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
lblnoBPB = lblkdgudang & "/B/" & Format(txttglBPB, "MMyy") & "/" & "0001"
End Sub




Private Sub cmdBR_Click()
PObeli_BR.LBLKODE = "BELI_D"
PObeli_BR.Show vbModal

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
    If rs!unit >= 0 Then
    Call Cetak
    Else
    Call Cetak1
    End If
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
    If rs!unit >= 0 Then
    Call Cetak
    Else
    Call Cetak1
    End If

End If
End Sub


Private Sub cmdsimpan_Click()

MousePointer = vbHourglass

Call Cek_tglOD
If CDate(txttglBPB) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    MousePointer = vbDefault
    Exit Sub

ElseIf txtkdPO = "" Or lblkdgudang = "" Or lblkdsupplier = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
    MousePointer = vbDefault
    Exit Sub
Else

    If LBLKODE = 1 Then
        Call nomer
        
        sqlA1 = "select a.kdbarang,b.nmbarang,a.unit,isnull(sum(d.unit),0) as Ubeli,b.satuan,a.keterangan,a.kdpobeli_d from pobeli_d a left join barang b " & vbCrLf & _
               "on a.kdbarang=b.kdbarang left join beli c on a.kdPObeli=c.kdPO left join beli_d d on c.kdbeli=d.kdbeli and a.kdbarang=d.kdbarang where a.kdpobeli='" & txtkdPO & "' " & vbCrLf & _
               "group by a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan,a.kdpobeli_d "
      
        sqlA = "select kdbarang,nmbarang,unit,Ubeli,unit-Ubeli as sisa,satuan,keterangan,kdpobeli_d from (" & sqlA1 & ") a "

        
        sql = "insert into beli values ('" & txtkdPO & "_" & lblnoBPB & "','" & lblnoBPB & "','" & Format(txttglBPB, "yyyy-MM-dd") & "','" & lblkdgudang & "','" & lblkdsupplier & "','" & UCase(txtketerangan) & "','" & UCase(txtSJ) & "','" & Format(txttglSJ, "yyyy/MM/dd") & "','" & txtkdPO & "','" & lblnoEASAP & "')"
        con.Execute (sql)
        
        sql = "insert into beli_d select kdbarang  + '" & txtkdPO & "_" & lblnoBPB & "','" & txtkdPO & "_" & lblnoBPB & "',kdbarang,SISA,0,0,keterangan from (" & sqlA & ") a where sisa<>0 "
        con.Execute (sql)
        
        txttglBPB.Enabled = False
        cmdBR.Enabled = False
        txtketerangan.Enabled = False
        cmdsimpan.Enabled = False
        cmdBR1.Enabled = False
        txtSJ.Enabled = False
        txttglSJ.Enabled = False
        cmdBatal.Enabled = True
    
    
    ElseIf LBLKODE = 2 Then
        sql = "Update beli set keterangan='" & UCase(txtketerangan) & "',noSJ='" & UCase(txtSJ) & "',tglsj='" & Format(txttglSJ, "yyyy/MM/dd") & "',kdsupplier='" & lblkdsupplier & "' where kdbeli='" & txtkdPO & "_" & lblnoBPB & "'"
        con.Execute (sql)
    
        txtketerangan.Enabled = False
        cmdsimpan.Enabled = False
        cmdBR1.Enabled = False
        txtSJ.Enabled = False
        txttglSJ.Enabled = False
    
        MsgBox "Header berhasil di Ubah ", vbInformation, "Info !"
    End If
 
End If
 
Beli.TimerALL.Interval = 10
Beli_D.TimerALL.Interval = 10

MousePointer = vbDefault

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
    If rs!unit >= 0 Then
    Call Cetak
    Else
    Call Cetak1
    End If

 
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0


txttglBPB = Date
txttglBPB.Enabled = True


txttglSJ = Date

TimerALL.Interval = 10
TimerNO.Interval = 10


Call nul(lblkdgudang)
Call nul(lblnmgudang)
Call nul(txtkdPO)
Call nul(lbltglPO)
Call nul(lblkdsupplier)
Call nul(lblnmsupplier)



End Sub


Private Sub Form_Unload(Cancel As Integer)
If cmdBR.Enabled = False And lblkdgudang <> "GD1" And UTAMA.lblstatus = 0 Then
sql2 = "select a.*,b.nobpb from beli_d a left join beli b on a.kdbeli=b.kdbeli where b.nobpb='" & lblnoBPB & "' and a.harga=0 and unit > 0"
Set rs2 = con.Execute(sql2)

    If rs2.RecordCount <> 0 Then
        MsgBox "tidak dapat keluar karena ada barang yg tidak ada harganya !!", vbCritical, "Error !"
        Cancel = 1
        Exit Sub
    Else
        Unload Me
    End If

End If

End Sub

Private Sub Label11_Click()

End Sub

Private Sub lblkdgudang_Change()
Call nul(lblkdgudang)
Call nomer
End Sub

Private Sub lblkdsupplier_Change()
Call nul(lblkdsupplier)
End Sub

Private Sub lblnmgudang_Change()
Call nul(lblnmgudang)
End Sub

Private Sub lblnmsupplier_Change()
Call nul(lblnmsupplier)
End Sub

Private Sub lbltglPO_Change()
Call nul(lbltglPO)
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all


If kode = 2 Then
rs.AbsolutePosition = lblpos
End If

 

TimerALL.Interval = 0

End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If


TimerNO.Interval = 0
End Sub



Private Sub txtkdPO_Change()
Call nul(txtkdPO)

sql1 = "select * from PObeli where kdPObeli='" & txtkdPO & "'"
Set rs1 = con.Execute(sql1)

If rs1.RecordCount <> 0 Then
lbltglPO = rs1!tglPObeli
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

Private Sub txtSJ_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtSJ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtSJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtSJ_LostFocus()
txtSJ = UCase(txtSJ)
End Sub

Private Sub txttglBPB_Change()
Call nul(txttglBPB)
Call nomer

End Sub

Private Sub txttglBPB_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglBPB_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglBPB_KeyPress(KeyAscii As Integer)
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

Private Sub txttglBPB_LostFocus()
On Error GoTo hell

txttglBPB = FormatDateTime(txttglBPB, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglBPB.SetFocus

End Sub

Private Sub txttglSJ_Change()
Call nul(txttglSJ)

End Sub

Private Sub txttglSJ_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglSJ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglSJ_KeyPress(KeyAscii As Integer)
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

Private Sub txttglSJ_LostFocus()
On Error GoTo hell

txttglSJ = FormatDateTime(txttglSJ, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglSJ.SetFocus

End Sub




