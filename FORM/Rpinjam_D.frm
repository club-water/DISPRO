VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Rpinjam_D 
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCHKrtr 
      Left            =   7695
      Top             =   225
   End
   Begin VB.TextBox txttglpengajuan 
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
      Left            =   4950
      TabIndex        =   0
      Top             =   990
      Width           =   1590
   End
   Begin VB.CheckBox CHKRtr 
      BackColor       =   &H00000000&
      Caption         =   "TGL RETUR :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6750
      TabIndex        =   1
      Top             =   990
      Width           =   1455
   End
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
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2070
      Width           =   1905
   End
   Begin VB.Timer TimerNO 
      Left            =   5670
      Top             =   225
   End
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
   End
   Begin VB.TextBox txttglRpinjam 
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
      Left            =   8190
      TabIndex        =   2
      Top             =   990
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
      Left            =   3735
      TabIndex        =   6
      Top             =   2070
      Width           =   10770
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   15
      Top             =   765
      Width           =   14280
      _Version        =   524288
      _ExtentX        =   25188
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   13995
      TabIndex        =   3
      ToolTipText     =   "Simpan"
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
      Picture         =   "Rpinjam_D.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   825
      Left            =   14625
      TabIndex        =   7
      ToolTipText     =   "Simpan"
      Top             =   1395
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1455
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
      Picture         =   "Rpinjam_D.frx":2832
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   16
      Top             =   2655
      Width           =   14505
      _Version        =   524288
      _ExtentX        =   25585
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
      TabIndex        =   8
      ToolTipText     =   "Tambah"
      Top             =   2745
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Rpinjam_D.frx":529F
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   1
      Left            =   14625
      TabIndex        =   9
      ToolTipText     =   "Ubah"
      Top             =   3555
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Rpinjam_D.frx":7F13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   2
      Left            =   14625
      TabIndex        =   10
      ToolTipText     =   "Hapus"
      Top             =   4365
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Rpinjam_D.frx":B110
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   3
      Left            =   14625
      TabIndex        =   11
      ToolTipText     =   "Refresh"
      Top             =   5175
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Rpinjam_D.frx":E1A9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   4
      Left            =   7830
      TabIndex        =   12
      ToolTipText     =   "Cetak Kwitansi"
      Top             =   7965
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Rpinjam_D.frx":11325
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   14
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
      Picture         =   "Rpinjam_D.frx":141F1
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand CmdBR1 
      Height          =   420
      Left            =   13005
      TabIndex        =   4
      ToolTipText     =   "Simpan"
      Top             =   1665
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
      Picture         =   "Rpinjam_D.frx":1AA53
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   5
      Left            =   14625
      TabIndex        =   13
      ToolTipText     =   "Cetak BPB"
      Top             =   5985
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Rpinjam_D.frx":1D285
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   5010
      Left            =   135
      TabIndex        =   37
      Top             =   2790
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
      FormatString    =   $"Rpinjam_D.frx":20CE2
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
   Begin VB.Label lbltotalrp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   12735
      TabIndex        =   36
      Top             =   7875
      Width           =   1545
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL RUPIAH :"
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
      Left            =   11475
      TabIndex        =   35
      Top             =   7920
      Width           =   1410
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   5805
      TabIndex        =   34
      Top             =   9000
      Width           =   1185
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
      Left            =   7425
      TabIndex        =   33
      Top             =   1350
      Width           =   6585
   End
   Begin VB.Label lblkode 
      Caption         =   "0"
      Height          =   285
      Left            =   3510
      TabIndex        =   32
      Top             =   9045
      Width           =   1545
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NO PJM :"
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
      Left            =   7560
      TabIndex        =   31
      Top             =   8910
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label txtkd 
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
      Left            =   8370
      TabIndex        =   30
      Top             =   8910
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL PJM :"
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
      Left            =   10755
      TabIndex        =   29
      Top             =   8955
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Retur Pinjaman"
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
      TabIndex        =   28
      Top             =   45
      Width           =   8025
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
      Left            =   1755
      TabIndex        =   27
      Top             =   1710
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MASUK GUDANG :"
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
      Top             =   1755
      Width           =   1500
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
      Left            =   2925
      TabIndex        =   25
      Top             =   1710
      Width           =   10050
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
      TabIndex        =   24
      Top             =   2115
      Width           =   1320
   End
   Begin VB.Label lbltgl 
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
      Left            =   11565
      TabIndex        =   23
      Top             =   8910
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblKDRPinjam 
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
      TabIndex        =   22
      Top             =   990
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
      TabIndex        =   21
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL PENGAJUAN :"
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
      Left            =   3510
      TabIndex        =   20
      Top             =   1035
      Width           =   1455
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
      TabIndex        =   19
      Top             =   1350
      Width           =   5145
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
      TabIndex        =   18
      Top             =   1395
      Width           =   1050
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
      TabIndex        =   17
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "Rpinjam_D.frx":20E12
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15495
   End
End
Attribute VB_Name = "Rpinjam_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK As ADODB.Recordset
Dim rsT As ADODB.Recordset
Dim a As Integer
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim color As Long, flag As Byte
Dim rsACC As ADODB.Recordset
Dim sqlACC As String
Dim rscek As ADODB.Recordset


Private Sub cek_dalem()
sqlcek = "select * from Rpinjam_d where kdRpinjam='" & lblKDRPinjam & "'"
Set rscek = con.Execute(sqlcek)
End Sub


Private Sub CHKRtr_Click()
txttglRpinjam = Date
TimerCHKrtr.Interval = 10
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
Gudang_BR.LBLKODE = "RPINJAM_D"
Gudang_BR.Show vbModal
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

'Unload AR_Kwitansi1
'
'With AR_Kwitansi1
'.fldnokwitansi = lblKDRPinjam
'.fldnmcustomer = lblnmcustomer
'.fldalamat = lblalamat
'.flduang = rs!rupiah
'.fldket1 = "PENGGANTIAN " & rs!nmbarang & " YG HILANG"
'.fldket2 = "JUMLAH = " & rs!UNIT & " ,HARGA = Rp " & Format(rs!harga, "#,###0") & " ( NO DISP : " & rs!kdbarang & " )"
'.fldjmlpiutang = Format(rs!rupiah, "#,###0")
'.fldtglposting = txttglRpinjam
'.lblKET = txtketerangan
'
'AR_Kwitansi1.Show vbModal
'
'
'End With
End Sub


Private Sub Cetak1()

Unload AR_LPB

sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from Rpinjam_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdRPinjam='" & lblKDRPinjam & "'  order by a.kdbarang"

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


.lblnoLPB = lblKDRPinjam
.lblsupplier = lblnmcustomer
.lbltglLPB = Format(txttglRpinjam, "dd/MM/yyyy")
.Lbljudul_sup = "Customer :"

sqlACC = "select * from Signature where kdFrm='" & lblkdgudang & "'"
Set rsACC = con.Execute(sqlACC)

.lblAcc1 = rsACC!Acc1
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
sql = "select a.kdbarang,b.kd1,b.nmbarang,a.unit,b.satuan,a.harga,a.rupiah,a.keterangan,a.kdRpinjam_d from Rpinjam_d a left join barang b " & vbCrLf & _
      "on a.kdbarang=b.kdbarang where a.kdRpinjam='" & lblKDRPinjam & "' order by a.kdbarang "
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


sqlT = "select kdRpinjam,sum(rupiah) as rupiah from Rpinjam_d  where kdRpinjam='" & lblKDRPinjam & "' group by kdRpinjam "
Set rsT = con.Execute(sqlT)

If rsT.RecordCount <> 0 Then

    lbltotalrp = FormatNumber(rsT!rupiah, 0)
    
    If lblkdgudang = "GH1" Then
        con.Execute ("delete from klaim_hilang where kdklaim='" & lblKDRPinjam & "'")
        con.Execute ("insert into klaim_hilang values ('" & lblKDRPinjam & "','" & lblkdcustomer & "','" & Format(txttglRpinjam, "yyyy/MM/dd") & "'," & CCur(lbltotalrp) & ",getdate(),'" & UTAMA.lblkduser & "')")
    End If

Else
    If lblkdgudang = "GH1" Then
        con.Execute ("delete from byrklaim where kdklaim='" & lblKDRPinjam & "'")
        con.Execute ("delete from klaim_hilang where kdklaim='" & lblKDRPinjam & "'")
    End If
    
    lbltotalrp = "0"
End If




Call LG
End Sub



Private Sub tbh()
Call Cek_tglOD
If CDate(txttglRpinjam) <= rstgl_OD!tglOD And CHKRtr.Value = 1 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

ElseIf CDate(txttglpengajuan) <= rstgl_OD!tglOD And CHKRtr.Value = 0 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else


    If cmdBR.Enabled = False Then
    Rpinjam_DTU.LBLKODE = 1
      
    
    Rpinjam_DTU.Show vbModal
    
    Else
    MsgBox "Kepala data belum disimpan !", vbCritical, "INfo !!"
    End If

End If


End Sub


Private Sub ubh()
Call Cek_tglOD
If CDate(txttglRpinjam) <= rstgl_OD!tglOD And CHKRtr.Value = 1 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

ElseIf CDate(txttglpengajuan) <= rstgl_OD!tglOD And CHKRtr.Value = 0 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    Rpinjam_DTU.LBLKODE = 2
    
    
    lblpos = rs.AbsolutePosition
    kode = 2
    
    
    Rpinjam_DTU.lblkdbarang = rs!kdbarang
    Rpinjam_DTU.lblnmbarang = rs!nmbarang
    Rpinjam_DTU.lblsatuan = rs!satuan
    Rpinjam_DTU.txtunit = FormatNumber(rs!unit, 0)
    Rpinjam_DTU.txtharga = FormatNumber(rs!harga, 0)
    Rpinjam_DTU.lblrupiah = FormatNumber(rs!rupiah, 0)
    Rpinjam_DTU.txtketerangan = rs!keterangan
    Rpinjam_DTU.lblkdRpinjam_d = rs!kdRpinjam_d
    Rpinjam_DTU.lblunit_awal = rs!unit
    Rpinjam_DTU.cmdBR.Enabled = False

End If
Rpinjam_DTU.Show vbModal
 
End Sub


Private Sub hps()
On Error GoTo hell

Call Cek_tglOD
If CDate(txttglRpinjam) <= rstgl_OD!tglOD And CHKRtr.Value = 1 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

ElseIf CDate(txttglpengajuan) <= rstgl_OD!tglOD And CHKRtr.Value = 0 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else

    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from Rpinjam_d where kdRpinjam_d ='" & rs!kdRpinjam_d & "'"
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
    sql = "select isnull(max(right(kdRpinjam,4)),0) as xx from Rpinjam where Month(tglpengajuan)='" & Month(txttglpengajuan) & "'  and year(tglpengajuan)='" & Year(txttglpengajuan) & "' and kdgudang= '" & lblkdgudang & "'"
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    lblKDRPinjam = lblkdgudang & "/H/" & Format(txttglpengajuan, "MMyy") & "/" & "000" & a
                Case 2
                    lblKDRPinjam = lblkdgudang & "/H/" & Format(txttglpengajuan, "MMyy") & "/" & "00" & a
                Case 3
                    lblKDRPinjam = lblkdgudang & "/H/" & Format(txttglpengajuan, "MMyy") & "/" & "0" & a
                Case 4
                    lblKDRPinjam = lblkdgudang & "/H/" & Format(txttglpengajuan, "MMyy") & "/" & a
        End Select
    
    Else
        lblKDRPinjam = lblkdgudang & "/H/" & Format(txttglpengajuan, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
lblKDRPinjam = lblkdgudang & "/H/" & Format(txttglpengajuan, "MMyy") & "/" & "0001"
End Sub




Private Sub cmdBR_Click()
PS_BR.LBLKODE = "RPINJAM_D"
PS_BR.lblkdkategori = "02"
PS_BR.Show vbModal

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
ElseIf Index = 5 Then
Call Cetak1
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
 Call Cetak1
End If
End Sub


Private Sub cmdsimpan_Click()
Call Cek_tglOD
If CDate(txttglRpinjam) <= rstgl_OD!tglOD And CHKRtr.Value = 1 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub

ElseIf CDate(txttglpengajuan) <= rstgl_OD!tglOD And CHKRtr.Value = 0 And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub


ElseIf CDate(txttglpengajuan) > CDate(txttglRpinjam) And CHKRtr.Value = 1 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat disimpan, Tgl pengajuan > Tgl Retur ", vbCritical, "Error !"
    Exit Sub
Else

    If lblkdgudang = "" Or lblkdcustomer = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
    Exit Sub
    Else
    
        If LBLKODE = 1 Then
            Call nomer
            
            sql = "insert into Rpinjam values ('" & lblKDRPinjam & "','" & Format(txttglpengajuan, "yyyy-MM-dd") & "'," & CHKRtr.Value & ",'" & Format(txttglRpinjam, "yyyy-MM-dd") & "','" & lblkdgudang & "','" & lblkdcustomer & "','" & UCase(txtketerangan) & "')"
            con.Execute (sql)
            
            txttglRpinjam.Enabled = False
            CHKRtr.Enabled = False
            txttglpengajuan.Enabled = False
            cmdBR.Enabled = False
            cmdBR1.Enabled = False
            CMBket.Enabled = False
            txtketerangan.Enabled = False
            cmdsimpan.Enabled = False
            
            
        
        
        ElseIf LBLKODE = 2 Then
            sql = "Update Rpinjam set keterangan='" & UCase(txtketerangan) & "',tglRpinjam='" & Format(txttglRpinjam, "yyyy-MM-dd") & "',rtr=" & CHKRtr.Value & " where kdRpinjam='" & lblKDRPinjam & "'"
            con.Execute (sql)
        
            txtketerangan.Enabled = False
            txttglRpinjam.Enabled = False
             CHKRtr.Enabled = False
            cmdBR1.Enabled = False
            CMBket.Enabled = False
            cmdsimpan.Enabled = False
            
            
            SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
            MsgBox "Header berhasil di Ubah ", vbInformation, "Info !"
        End If
     
    End If
     
    Rpinjam.TimerALL.Interval = 10
    TimerALL.Interval = 10


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
 Call Cetak1
 
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttglpengajuan = Date
txttglpengajuan.Enabled = True




CMBket.AddItem "RETUR"
CMBket.AddItem "PENGGANTIAN"
CMBket.AddItem "LAIN - LAIN"
CMBket.ListIndex = 0


TimerALL.Interval = 10
TimerNO.Interval = 10


Call nul(lblkdgudang)
Call nul(lblnmgudang)
Call nul(lbltgl)
Call nul(lblkdcustomer)
Call nul(lblnmcustomer)
Call nul(lblalamat)


End Sub



Private Sub Form_Unload(Cancel As Integer)
Call cek_dalem

If lbltotalrp = 0 And rscek.RecordCount <> 0 And lblkdgudang = "GH1" Then
MsgBox "Tidak bisa Keluar dari menu ini , Karena Tidak Ada Rupiah yg Akan di Klaimkan ?", vbCritical, "Error"
End If


If txttglpengajuan.Enabled = False And rscek.RecordCount = 0 Then
 ms = MsgBox("Tidak Ada Detail Retur, apa anda ingin membatalkan Header Retur ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = " delete from Rpinjam where kdRpinjam='" & lblKDRPinjam & "' "
        con.Execute (sql)
           
        Rpinjam.TimerALL.Interval = 10
           
        Unload Me
        
    Else
        Cancel = 1
    End If
End If



End Sub

Private Sub lblalamat_Change()
Call nul(lblalamat)
End Sub

Private Sub lblkdRpinjam_Change()
Call nul(lblKDRPinjam)
End Sub

Private Sub lblkdgudang_Change()
Call nul(lblkdgudang)
Call nomer
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



Private Sub lbltgl_Change()
Call nul(lbltgl)
End Sub


Private Sub Text1_Change()

End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all


If rs.RecordCount = 0 And txttglpengajuan.Enabled = False Then
cmdT(0).SetFocus
Else
datagrid1.SetFocus
End If


If kode = 2 Then
rs.AbsolutePosition = lblpos
End If

 

TimerALL.Interval = 0

End Sub

Private Sub TimerCHKrtr_Timer()
If CHKRtr.Value = 0 Then
    txttglRpinjam = "01/01/1900"
    txttglRpinjam.Enabled = False
    
Else
   
    txttglRpinjam.Enabled = True
    
    
End If

TimerCHKrtr.Interval = 0
End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If


TimerNO.Interval = 0
End Sub



Private Sub txtkd_Change()
'Call nul(txtkd)
'
'sql1 = "select * from Pinjam where kdpinjam='" & txtkd & "'"
'Set rs1 = con.Execute(sql1)
'
'If rs1.RecordCount <> 0 Then
'lbltgl = rs1!tglpinjam
'End If
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






Private Sub txttglpengajuan_Change()
Call nul(txttglpengajuan)
Call nomer

End Sub

Private Sub txttglpengajuan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglpengajuan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglpengajuan_KeyPress(KeyAscii As Integer)
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

Private Sub txttglpengajuan_LostFocus()
On Error GoTo hell

txttglpengajuan = FormatDateTime(txttglpengajuan, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglpengajuan.SetFocus

End Sub


Private Sub txttglRpinjam_Change()
Call nul(txttglRpinjam)
'Call nomer

End Sub

Private Sub txttglRpinjam_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglRpinjam_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglRpinjam_KeyPress(KeyAscii As Integer)
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

Private Sub txttglRpinjam_LostFocus()
On Error GoTo hell

txttglRpinjam = FormatDateTime(txttglRpinjam, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglRpinjam.SetFocus

End Sub












