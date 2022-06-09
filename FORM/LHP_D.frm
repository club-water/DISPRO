VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form LHP_D 
   BorderStyle     =   0  'None
   Caption         =   "LHP_D"
   ClientHeight    =   10230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19650
   LinkTopic       =   "Form2"
   ScaleHeight     =   10230
   ScaleWidth      =   19650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglCLR 
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
      Left            =   12735
      TabIndex        =   2
      Top             =   945
      Width           =   1590
   End
   Begin VB.CheckBox ChKCLEAR 
      BackColor       =   &H00000000&
      Caption         =   "LHP Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   14355
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   945
      Width           =   1365
   End
   Begin VB.Timer TimerG 
      Left            =   2385
      Top             =   4050
   End
   Begin VB.Timer TimerAll 
      Left            =   1890
      Top             =   4050
   End
   Begin VB.TextBox txttglLHP 
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
      Left            =   1305
      TabIndex        =   0
      Top             =   945
      Width           =   1590
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   315
      TabIndex        =   8
      Top             =   720
      Width           =   18060
      _Version        =   524288
      _ExtentX        =   31856
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
      TabIndex        =   9
      Top             =   1440
      Width           =   18105
      _Version        =   524288
      _ExtentX        =   31935
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
      Left            =   18720
      TabIndex        =   4
      ToolTipText     =   "Tambah"
      Top             =   1845
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
      Picture         =   "LHP_D.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   12330
      TabIndex        =   10
      ToolTipText     =   "Ubah"
      Top             =   -135
      Visible         =   0   'False
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
      Picture         =   "LHP_D.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   2
      Left            =   18720
      TabIndex        =   5
      ToolTipText     =   "Hapus"
      Top             =   2655
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
      Picture         =   "LHP_D.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   3
      Left            =   18720
      TabIndex        =   6
      ToolTipText     =   "Refresh"
      Top             =   3465
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
      Picture         =   "LHP_D.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   4
      Left            =   18720
      TabIndex        =   7
      ToolTipText     =   "Cetak"
      Top             =   4275
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
      Picture         =   "LHP_D.frx":C086
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   990
      TabIndex        =   11
      Top             =   9630
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
      Picture         =   "LHP_D.frx":FAE3
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   10305
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
      Picture         =   "LHP_D.frx":16345
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7440
      Left            =   225
      TabIndex        =   30
      Top             =   1800
      Width           =   18195
      _cx             =   32094
      _cy             =   13123
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
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14737632
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"LHP_D.frx":18B77
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
   Begin VB.Label lblhari 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL LHP :"
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
      Left            =   2970
      TabIndex        =   32
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL CLR :"
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
      Left            =   11880
      TabIndex        =   31
      Top             =   990
      Width           =   780
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "TANDA TERIMA :"
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
      Left            =   11385
      TabIndex        =   29
      Top             =   9405
      Width           =   1455
   End
   Begin VB.Label lblTT 
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
      Left            =   12690
      TabIndex        =   28
      Top             =   9360
      Width           =   1410
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TDK TERTAGIH :"
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
      TabIndex        =   27
      Top             =   9405
      Width           =   1455
   End
   Begin VB.Label lbltdk_tertagih 
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
      Left            =   9540
      TabIndex        =   26
      Top             =   9360
      Width           =   1410
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TERTAGIH :"
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
      TabIndex        =   25
      Top             =   9405
      Width           =   1455
   End
   Begin VB.Label lbltertagih 
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
      Left            =   6345
      TabIndex        =   24
      Top             =   9360
      Width           =   1410
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   18585
      Picture         =   "LHP_D.frx":18CF2
      Stretch         =   -1  'True
      Top             =   585
      Width           =   600
   End
   Begin VB.Label lblTotal 
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
      Left            =   16965
      TabIndex        =   23
      Top             =   9360
      Width           =   1410
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL TAGIHAN :"
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
      Left            =   15435
      TabIndex        =   22
      Top             =   9405
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "3. UNTUK  STATUS TT (TANDA TERIMA)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   9630
      TabIndex        =   21
      Top             =   1530
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2. UNTUK  STATUS TDK TERTAGIH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5715
      TabIndex        =   20
      Top             =   1530
      Width           =   3795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TEKAN TOMBOL :    1. UNTUK  STATUS TERTAGIH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   405
      TabIndex        =   19
      Top             =   1530
      Width           =   5010
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   6930
      TabIndex        =   18
      Top             =   9945
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   4725
      TabIndex        =   17
      Top             =   9855
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL LHP :"
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
      Left            =   540
      TabIndex        =   16
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rincian LHP"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   45
      Width           =   6000
   End
   Begin VB.Label lblnmkolektor 
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
      TabIndex        =   14
      Top             =   945
      Width           =   4020
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "KOLEKTOR :"
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
      Left            =   4140
      TabIndex        =   13
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label lblkdkolektor 
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
      Left            =   5130
      TabIndex        =   12
      Top             =   945
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   10185
      Left            =   0
      Picture         =   "LHP_D.frx":190B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19635
   End
End
Attribute VB_Name = "LHP_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK, rsT As ADODB.Recordset
Dim a As Integer
Dim KODE As Integer
Dim rsX As ADODB.Recordset
Dim sqlA As String
Dim color As Long, flag As Byte
Dim rscek As ADODB.Recordset
Dim rsTot As ADODB.Recordset
Dim rsTot1 As ADODB.Recordset
Dim rsTot2 As ADODB.Recordset
Dim rsTot3 As ADODB.Recordset
Dim sqlT, sqlT1, sqlT2, sqlT3 As String
Dim rsPPH As ADODB.Recordset



Private Sub total_LHP()
sqlT = "select tgllhp,kdkolektor, sum(rpLHP) as rpLHP from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "' group by tgllhp,kdkolektor"
Set rsTot = con.Execute(sqlT)

If rsTot.RecordCount = 0 Then
lblTotal = 0
Else
lblTotal = FormatNumber(rsTot!rpLHP, 0)
End If

'total tertagih
sqlT1 = "select tgllhp,kdkolektor, sum(rpLHP) as rpLHP from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "' and status='TERTAGIH' group by tgllhp,kdkolektor"
Set rsTot1 = con.Execute(sqlT1)

If rsTot1.RecordCount = 0 Then
lbltertagih = 0
Else
lbltertagih = FormatNumber(rsTot1!rpLHP, 0)
End If


'total tdk tertagih
sqlT2 = "select tgllhp,kdkolektor, sum(rpLHP) as rpLHP from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "' and status='TDK TERTAGIH' group by tgllhp,kdkolektor"
Set rsTot2 = con.Execute(sqlT2)

If rsTot2.RecordCount = 0 Then
lbltdk_tertagih = 0
Else
lbltdk_tertagih = FormatNumber(rsTot2!rpLHP, 0)
End If


'total tertagih
sqlT3 = "select tgllhp,kdkolektor, sum(rpLHP) as rpLHP from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "' and status='TANDA TERIMA' group by tgllhp,kdkolektor"
Set rsTot3 = con.Execute(sqlT3)

If rsTot3.RecordCount = 0 Then
lblTT = 0
Else
lblTT = FormatNumber(rsTot3!rpLHP, 0)
End If




End Sub

Private Sub cek_dalem()
sqlcek = "select * from PObeli_D where kdPObeli='" & txtkdPO & "'"
Set rscek = con.Execute(sqlcek)
End Sub


Private Sub cmdBR1_Click()

End Sub

Private Sub ChKCLEAR_Click()

If ChKCLEAR.Value = 0 Then
txttglCLR.Enabled = True
Else
txttglCLR.Enabled = False
End If


If lblkode = 1 Then


    If ChKCLEAR.Value = 1 Then
        sql = "update lhp set clr=" & ChKCLEAR.Value & ",tglCLR='" & Format(txttglCLR, "yyyy/MM/dd") & "',tglinput_clr=getdate() ,status='TERTAGIH',keterangan='' where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "'"
        con.Execute (sql)
        
        sql = "insert into byrpiutangsewa select convert(varchar,(ISNULL(b.urut,0) + 1)) + a.kdpiutang as kdbyrpiutang,ISNULL(b.urut,0) + 1 as urut,a.tglCLR," & vbCrLf & _
              "left(a.kdpiutang,6) as kdcustomer,a.kdkolektor,a.rpLHP,0,0,'LHP',a.kdpiutang,0  from LHP a left join byrpiutangSewa b " & vbCrLf & _
              "on a.kdpiutang=b.kdpiutang  where a.kdkolektor='" & lblkdkolektor & "' and a.tglLHP='" & Format(txttglLHP, "yyyy/MM/dd") & "'"
        con.Execute (sql)
        
       
    ElseIf ChKCLEAR.Value = 0 Then
        sql = "update lhp set clr=" & ChKCLEAR.Value & ",tglCLR='" & Format(txttglLHP, "yyyy/MM/dd") & "',tglinput_clr=getdate() ,status='',keterangan='' where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "'"
        con.Execute (sql)
        
        sql = "delete from byrpiutangsewa where kdpiutang in (select kdpiutang from lhp where kdkolektor='" & lblkdkolektor & "' and tglLHP='" & Format(txttglLHP, "yyyy/MM/dd") & "') and tglbayar ='" & Format(txttglCLR, "yyyy/MM/dd") & "' and keterangan='LHP' "
        con.Execute (sql)
        
        sql3 = "select * from Tanda_terima where kdpiutang in (select kdpiutang from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "') and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'"
        Set rs3 = con.Execute(sql3)
        
        If rs3.RecordCount <> 0 Then
        con.Execute ("delete from Tanda_terima where kdpiutang in (select kdpiutang from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "')  and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'")
        con.Execute ("update piutangsewa set tt=0 where kdpiutang in (select kdpiutang from LHP where tgllhp='" & Format(txttglLHP, "yyyy/MM/dd") & "' and kdkolektor='" & lblkdkolektor & "')")
        End If
    End If
    
    If UTAMA.lblstatus = 0 Then
    ChKCLEAR.Enabled = False
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    End If
    
    TimerAll.Interval = 10
    LHP.TimerAll.Interval = 10
End If

End Sub

Private Sub ChKCLEAR_KeyPress(KeyAscii As Integer)
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


Private Sub cmdsimpan_Click()

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

Unload AR_LHP

sqlX = "select a.kdpiutang,c.bln,c.tahun,c.kdcustomer,b.nmcustomer,b.alamat_TGH as alamat,a.rpLHP,c.TT,N_TT = case when c.TT=1 then 'X' else '' end,a.status,a.keterangan,a.kdlhp  from LHP a " & vbCrLf & _
      "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join customer b on c.kdcustomer=b.kdcustomer where a.tglLHP='" & Format(txttglLHP, "yyyy/MM/dd") & "' and a.kdkolektor='" & lblkdkolektor & "' order by c.kdcustomer,c.tahun,c.bln"

Set rsX = con.Execute(sqlX)

With AR_LHP.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_LHP
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldkdpiutang.DataField = "kdpiutang"
.fldrpLHP.DataField = "rplhp"
.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")
.lblnmkolektor = lblnmkolektor
.lblnmkolektor1 = "( " & lblnmkolektor & " )"
.lbltglLHP = Format(txttglLHP, "dd/MM/yyyy") & "   ( " & lblhari & " )"
.lblTotal = lblTotal

AR_LHP.Show vbModal

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
    txttglLHP.Enabled = True
    cmdBR.Enabled = True
    datagrid1.Enabled = False
    cmdT(2).Enabled = False
    txttglLHP.SetFocus
    

Else
    txttglLHP.Enabled = False
    cmdBR.Enabled = False
    datagrid1.Enabled = True
    cmdT(2).Enabled = True
    datagrid1.SetFocus
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

sql1 = "select a.kdpiutang,c.bln,c.tahun,c.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.rpLHP,c.TT,N_TT = case when c.TT=1 then 'X' else '' end,a.status,a.keterangan,a.kdlhp  from LHP a " & vbCrLf & _
      "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join customer b on c.kdcustomer=b.kdcustomer where a.tglLHP='" & Format(txttglLHP, "yyyy/MM/dd") & "' and a.kdkolektor='" & lblkdkolektor & "' order by c.kdcustomer,c.tahun,c.bln"
      
      
Set rs = con.Execute(sql1)

Set datagrid1.DataSource = rs

Call total_LHP

Call LG

For i = 1 To (datagrid1.Rows - 1)
For j = 1 To (datagrid1.Cols - 1)


If datagrid1.TextMatrix(i, 10) = "TERTAGIH" Then
datagrid1.Cell(flexcpForeColor, i, j) = vbRed
ElseIf datagrid1.TextMatrix(i, 10) = "TDK TERTAGIH" Then
datagrid1.Cell(flexcpForeColor, i, j) = &H80FF&
ElseIf datagrid1.TextMatrix(i, 10) = "TANDA TERIMA" Then
datagrid1.Cell(flexcpForeColor, i, j) = &HFF00&
End If

Next
Next


MousePointer = vbDefault
End Sub



Private Sub tbh()
If lblkdkolektor = "" Or txttglLHP = "" Then
    MsgBox "Header Belum Lengkap !!", vbCritical, "Error !"
    Exit Sub
Else
    If ChKCLEAR.Value = 0 Then
    Piutang_BR1.Show vbModal
    Else
    MsgBox "data tidak dapat ditambah karena LHP Sudah Clear", vbCritical, "Error !!"
    End If
End If

End Sub


Private Sub ubh()
End Sub


Private Sub hps()
On Error GoTo hell

If ChKCLEAR.Value = 1 Then
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "Data Tidak dapat dihapus karena LHP sudah Clear", vbCritical, "Error !"
Exit Sub
Else

    KODE = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        
        sql = "delete from LHP where kdLHP ='" & rs!kdLHP & "'"
        con.Execute (sql)
        
        sql3 = "select * from Tanda_terima where kdpiutang='" & rs!kdpiutang & "' and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'"
        Set rs3 = con.Execute(sql3)
        
        If rs3.RecordCount <> 0 Then
        con.Execute ("delete from Tanda_terima where kdpiutang='" & rs!kdpiutang & "' and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'")
        con.Execute ("update piutangsewa set tt=0 where kdpiutang='" & rs!kdpiutang & "'")
        End If
        
        TimerAll.Interval = 10
        LHP.TimerAll.Interval = 10
    End If

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub











Private Sub cmdBR_Click()
Kolektor_BR.lblkode = "LHP_D"
Kolektor_BR.Show vbModal

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
 txtcari = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
End If
End Sub



Private Sub DataGrid1_DblClick()
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
On Error GoTo hell

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
txtcari = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
ElseIf KeyAscii = Asc("1") And ChKCLEAR.Value = 1 Then
KODE = 2
lblpos = rs.AbsolutePosition
sql = "update LHP set status='TERTAGIH',keterangan='' where kdLHP='" & rs!kdLHP & "'"
con.Execute (sql)


con.Execute ("delete from byrpiutangsewa where kdpiutang='" & rs!kdpiutang & "' and tglbayar='" & Format(txttglCLR, "yyyy/MM/dd") & "' and keterangan='LHP' ")

con.Execute ("insert into byrpiutangsewa select convert(varchar,(ISNULL(b.urut,0) + 1)) + a.kdpiutang as kdbyrpiutang,ISNULL(b.urut,0) + 1 as urut,a.tglCLR," & vbCrLf & _
             "left(a.kdpiutang,6) as kdcustomer,a.kdkolektor,a.rpLHP,0,0,'LHP',a.kdpiutang,0  from LHP a left join byrpiutangSewa b " & vbCrLf & _
             "on a.kdpiutang=b.kdpiutang  where a.kdkolektor='" & lblkdkolektor & "' and a.tglLHP='" & Format(txttglLHP, "yyyy/MM/dd") & "' and a.kdpiutang='" & rs!kdpiutang & "' ")


 sql3 = "select * from Tanda_terima where kdpiutang='" & rs!kdpiutang & "' and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'"
 Set rs3 = con.Execute(sql3)
      
 If rs3.RecordCount <> 0 Then
 con.Execute ("delete from Tanda_terima where kdpiutang='" & rs!kdpiutang & "' and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'")
 con.Execute ("update piutangsewa set tt=0 where kdpiutang='" & rs!kdpiutang & "'")
 End If
 
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "Status : TERTAGIH", vbInformation, "Info !"
TimerAll.Interval = 10

ElseIf KeyAscii = Asc("2") And ChKCLEAR.Value = 1 Then

ms = InputBox("Masukkan Keterangan Tdk Tertagih !", "KETERANGAN TDK TERTAGIH")

    If ms = "" Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Tidak Boleh Kosong !!", vbCritical, "Error !"
        Exit Sub
    Else
        KODE = 2
        lblpos = rs.AbsolutePosition
        sql = "update LHP set status='TDK TERTAGIH',keterangan='" & UCase(ms) & "' where kdLHP='" & rs!kdLHP & "'"
        con.Execute (sql)
         
        con.Execute ("delete from byrpiutangsewa where kdpiutang='" & rs!kdpiutang & "' and tglbayar='" & Format(txttglCLR, "yyyy/MM/dd") & "' and keterangan='LHP' ")
         
        sql3 = "select * from Tanda_terima where kdpiutang='" & rs!kdpiutang & "' and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'"
        Set rs3 = con.Execute(sql3)
             
        If rs3.RecordCount <> 0 Then
        con.Execute ("delete from Tanda_terima where kdpiutang='" & rs!kdpiutang & "' and tglTT='" & Format(txttglLHP, "yyyy/MM/dd") & "'")
        con.Execute ("update piutangsewa set tt=0 where kdpiutang='" & rs!kdpiutang & "'")
        End If
 
         
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Status : TDK TERTAGIH", vbInformation, "Info !"
        TimerAll.Interval = 10
    End If

 
ElseIf KeyAscii = Asc("3") And ChKCLEAR.Value = 1 Then
KODE = 2
lblpos = rs.AbsolutePosition
 sql = "update LHP set status='TANDA TERIMA',keterangan='' where kdLHP='" & rs!kdLHP & "'"
 con.Execute (sql)
 
 con.Execute ("delete from byrpiutangsewa where kdpiutang='" & rs!kdpiutang & "' and tglbayar='" & Format(txttglCLR, "yyyy/MM/dd") & "' and keterangan='LHP' ")
 
 sql = "update piutangsewa set TT=1 where kdpiutang='" & rs!kdpiutang & "' "
 con.Execute (sql)
 
 sql = "insert into tanda_terima values('" & rs!kdpiutang & "','" & Format(txttglLHP, "yyyy/MM/dd") & "' ) "
 con.Execute (sql)
 
 SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
 MsgBox "Status : TANDA TERIMA", vbInformation, "Info !"
 TimerAll.Interval = 10

 
 
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub Form_Load()
GradientForm Me, 0


txttglLHP = Date
txttglCLR = Date

TimerAll.Interval = 10



Call nul(lblkdkolektor)
Call nul(lblnmkolektor)


End Sub


Private Sub lblkdgudang_Change()
Call nul(lblkdgudang)
End Sub

Private Sub lblnmgudang_Change()
Call nul(lblnmgudang)
End Sub

Private Sub lblalamat_Change()
Call nul(lblalamat)
End Sub

Private Sub Form_Unload(Cancel As Integer)
sqlP1 = "select kdpiutang from LHP a left join Customer b on left(a.kdpiutang,6)=b.kdcustomer where b.PPH23=1 and a.tglLHP='" & Format(LHP_D.txttglLHP, "yyyy/MM/dd") & "' and a.status='TERTAGIH' "

sqlP2 = "select * from byrpiutangSewa where kdpiutang in (" & sqlP1 & " ) and tglbayar='" & Format(LHP_D.txttglCLR, "yyyy/MM/dd") & "' and keterangan='LHP' and trf=0 and kdkolektor='" & LHP_D.lblkdkolektor & "'"

sqlP = "select a.kdbyrpiutang,a.urut,a.tglbayar,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlbayar,a.rpPPH23,a.potongan from (" & sqlP2 & ") a left join customer b on a.kdcustomer=b.kdcustomer"

Set rsPPH = con.Execute(sqlP)

If rsPPH.RecordCount <> 0 Then

 
 List_input_PPH23.Show vbModal

End If



End Sub

Private Sub lblkdkolektor_Change()
Call nul(lblkdkolektor)
End Sub

Private Sub lblnmkolektor_Change()
Call nul(lblnmkolektor)
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all


If KODE = 2 Then
rs.AbsolutePosition = lblpos
End If

lblkode = 1

TimerAll.Interval = 0
MousePointer = vbDefault

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

Private Sub txttglCLR_Change()
Call nul(txttglCLR)
End Sub

Private Sub txttglCLR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglCLR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txttglCLR_KeyPress(KeyAscii As Integer)
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

Private Sub txttglCLR_LostFocus()
On Error GoTo hell

txttglCLR = FormatDateTime(txttglCLR, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglCLR.SetFocus

End Sub

Private Sub txttglLHP_Change()
On Error GoTo hell

Call nul(txttglLHP)

If WeekdayName(Weekday(txttglLHP)) = "Sunday" Then
lblhari = "MINGGU"
ElseIf WeekdayName(Weekday(txttglLHP)) = "Monday" Then
lblhari = "SENIN"
ElseIf WeekdayName(Weekday(txttglLHP)) = "Tuesday" Then
lblhari = "SELASA"
ElseIf WeekdayName(Weekday(txttglLHP)) = "Wednesday" Then
lblhari = "RABU"
ElseIf WeekdayName(Weekday(txttglLHP)) = "Thursday" Then
lblhari = "KAMIS"
ElseIf WeekdayName(Weekday(txttglLHP)) = "Friday" Then
lblhari = "JUMAT"
ElseIf WeekdayName(Weekday(txttglLHP)) = "Saturday" Then
lblhari = "SABTU"
End If

Exit Sub
hell:
lblhari = ""


End Sub

Private Sub txttglLHP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglLHP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txttglLHP_KeyPress(KeyAscii As Integer)
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

Private Sub txttglLHP_LostFocus()
On Error GoTo hell

txttglLHP = FormatDateTime(txttglLHP, vbGeneralDate)
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglLHP.SetFocus

End Sub






