VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Piutang_D 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
   End
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin VB.Timer TimerNO 
      Left            =   1755
      Top             =   720
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   5
      Top             =   720
      Width           =   13920
      _Version        =   524288
      _ExtentX        =   24553
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   180
      TabIndex        =   6
      Top             =   2205
      Width           =   13920
      _Version        =   524288
      _ExtentX        =   24553
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   0
      Left            =   14265
      TabIndex        =   0
      ToolTipText     =   "Tambah"
      Top             =   2385
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1455
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
      Picture         =   "Piutang_D.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   1
      Left            =   14265
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   3240
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Piutang_D.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   2
      Left            =   14265
      TabIndex        =   2
      ToolTipText     =   "Hapus"
      Top             =   4095
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Piutang_D.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   3
      Left            =   14265
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   4950
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Piutang_D.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   7
      Top             =   7380
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
      Picture         =   "Piutang_D.frx":C086
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   3750
      Left            =   180
      TabIndex        =   4
      Top             =   2385
      Width           =   13965
      _cx             =   24633
      _cy             =   6615
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
      BackColorAlternate=   12648447
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Piutang_D.frx":128E8
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
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "PPH 23 :"
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
      TabIndex        =   32
      Top             =   6300
      Width           =   735
   End
   Begin VB.Label lblrpPPH23 
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
      Left            =   7695
      TabIndex        =   31
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   285
      Left            =   5895
      TabIndex        =   30
      Top             =   8280
      Width           =   1680
   End
   Begin VB.Label lblsisa 
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
      Left            =   12330
      TabIndex        =   29
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "SISA :"
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
      Left            =   11790
      TabIndex        =   28
      Top             =   6300
      Width           =   735
   End
   Begin VB.Label lblpotongan 
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
      Left            =   10260
      TabIndex        =   27
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "POTONGAN :"
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
      Left            =   9225
      TabIndex        =   26
      Top             =   6300
      Width           =   1230
   End
   Begin VB.Label lbljmlbyr 
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
      Left            =   5400
      TabIndex        =   25
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "JMLBAYAR :"
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
      Left            =   4410
      TabIndex        =   24
      Top             =   6300
      Width           =   1230
   End
   Begin VB.Label lbljmlpiutang 
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
      Left            =   2880
      TabIndex        =   23
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "JML PIUTANG :"
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
      Left            =   1710
      TabIndex        =   22
      Top             =   6300
      Width           =   1230
   End
   Begin VB.Label lblalamat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESEMBER"
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
      Left            =   1575
      TabIndex        =   21
      Top             =   1665
      Width           =   12570
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN :"
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
      Left            =   8415
      TabIndex        =   20
      Top             =   990
      Width           =   780
   End
   Begin VB.Label lbltahun 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2017"
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
      Left            =   9090
      TabIndex        =   19
      Top             =   945
      Width           =   915
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BULAN :"
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
      TabIndex        =   18
      Top             =   990
      Width           =   780
   End
   Begin VB.Label lblbln 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESEMBER"
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
      Left            =   7110
      TabIndex        =   17
      Top             =   945
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL POSTING :"
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
      Left            =   3825
      TabIndex        =   16
      Top             =   990
      Width           =   1230
   End
   Begin VB.Label lbltglposting 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "22/12/2017"
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
      Left            =   4995
      TabIndex        =   15
      Top             =   945
      Width           =   1365
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
      Left            =   2745
      TabIndex        =   14
      Top             =   1305
      Width           =   11400
   End
   Begin VB.Label Label3 
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
      Left            =   405
      TabIndex        =   13
      Top             =   1350
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
      Left            =   1575
      TabIndex        =   12
      Top             =   1305
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembayaran Piutang"
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
      TabIndex        =   11
      Top             =   45
      Width           =   6000
   End
   Begin VB.Label txtkdPiutang 
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
      Left            =   1575
      TabIndex        =   10
      Top             =   945
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NO KWITANSI :"
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
      TabIndex        =   9
      Top             =   990
      Width           =   1905
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3780
      TabIndex        =   8
      Top             =   8235
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   7980
      Left            =   0
      Picture         =   "Piutang_D.frx":12A51
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15090
   End
End
Attribute VB_Name = "Piutang_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK, rsT As ADODB.Recordset
Dim a As Integer
Dim KODE As Integer
Dim rsX As ADODB.Recordset
Dim color As Long, flag As Byte
Dim rsbp As ADODB.Recordset
Dim sqlbp As String
Dim rsC As ADODB.Recordset
Dim rpPPH_X As Currency


Private Sub jmlpiut()
On Error GoTo hell
sqlbp = "select isnull(sum(jmlbayar),0)as Tbyr,isnull(sum(rpPPH23),0)as Tpph23,isnull(sum(potongan),0) as Tpotongan from byrpiutangsewa where kdpiutang='" & txtkdPiutang & "' "
Set rsbp = con.Execute(sqlbp)

If rsbp!Tbyr <> 0 Then
lbljmlbyr = Format(rsbp!Tbyr, "###,###,###,###")
Else
lbljmlbyr = 0
End If

If rsbp!Tpph23 <> 0 Then
lblrpPPH23 = Format(rsbp!Tpph23, "###,###,###,###")
Else
lblrpPPH23 = 0
End If


If rsbp!Tpotongan <> 0 Then
lblpotongan = Format(rsbp!Tpotongan, "###,###,###,###")
Else
lblpotongan = 0
End If


lblsisa = CCur(lbljmlpiutang) - CCur(lbljmlbyr) - CCur(lblrpPPH23) - CCur(lblpotongan)
lblsisa = Format(lblsisa, "#,###0")
Exit Sub
hell:

'lbljmlbyr = "0"
'lblpotongan = "0"
'lblsisa = lbljmlpiut - lblretur
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

Unload AR_PObeli

sqlX = "select urut,tglbayar,jmlbayar,rpPPH23,potongan,Jnsbyr = case when trf=1 then 'TRANSFER' else 'TUNAI' end,trf,keterangan from byrpiutangsewa  " & vbCrLf & _
       "where a.kdpobeli='" & txtkdPO & "' order by a.kdbarang"

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
.lblnmgudang = lblnmgudang
.lbltglPO = Format(txttglPO, "dd/MM/yyyy")

If txtketerangan = "" Then
.lblNB = ""
Else
.lblNB = "NB : " & txtketerangan
End If

.lbljudul2.Visible = False
.lblkategori.Visible = False


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

With datagrid1.Columns(0)
.Caption = "ANGS"
.Width = 40
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "TGL BAYAR"
.Width = 85
.Alignment = dbgCenter
End With

With datagrid1.Columns(2)
.Caption = "JML BAYAR"
.Width = 100
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With

With datagrid1.Columns(3)
.Caption = "POTONGAN"
.Width = 80
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With


With datagrid1.Columns(4)
.Caption = "JNS BYR"
.Width = 80
.Alignment = dbgCenter
End With

With datagrid1.Columns(5)
.Caption = "TRF"
.Width = 0
End With

With datagrid1.Columns(6)
.Caption = "KETERANGAN"
.Width = 155
End With

With datagrid1.Columns(7)
.Caption = "Kd kolektor"
.Width = 0
End With

With datagrid1.Columns(8)
.Caption = "KOLEKTOR"
.Width = 80
End With

With datagrid1.Columns(9)
.Caption = "KD BYR"
.Width = 0
End With



Call tbl

Exit Sub
hell:
End Sub


Private Sub all()
sql = "select a.urut,a.tglbayar,a.jmlbayar,a.rpPPH23,a.potongan,Jnsbyr = case when a.trf=1 then 'TRANSFER' else 'TUNAI' end,a.trf,a.keterangan,a.kdkolektor,ISNULL(b.nmkolektor,'') AS nmkolektor,a.kdbyrpiutang from byrpiutangsewa a " & vbCrLf & _
      "left join kolektor b on a.kdkolektor=b.kdkolektor where a.kdpiutang='" & txtkdPiutang & "' order by a.urut"

Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


Call LG
End Sub



Private Sub tbh()

Piutang_DTU.LBLKODE = 1


sqlC = "select * from customer where kdcustomer ='" & lblkdcustomer & "' and PPH23=1"
Set rsC = con.Execute(sqlC)
    
If rsC.RecordCount <> 0 Then
rpPPH_X = FormatNumber((CCur(lblsisa) / 1.1) * 0.02, 0)
Piutang_DTU.txtjmlbayar = FormatNumber(CCur(lblsisa), 0)
Piutang_DTU.txtPPH23 = rpPPH_X
Piutang_DTU.lblsisa_awal = FormatNumber(CCur(lblsisa), 0)
Piutang_DTU.txtPPH23.Enabled = True
Piutang_DTU.Show vbModal

Else

rpPPH_X = FormatNumber((CCur(lblsisa) / 1.1) * 0.02, 0)
Piutang_DTU.txtjmlbayar = lblsisa
Piutang_DTU.lblsisa_awal = lblsisa
Piutang_DTU.txtPPH23.Enabled = False
Piutang_DTU.Show vbModal
End If

End Sub


Private Sub ubh()

Piutang_DTU.LBLKODE = 2


lblpos = rs.AbsolutePosition
KODE = 2



Piutang_DTU.lblkdkolektor = rs!kdkolektor
Piutang_DTU.lblnmkolektor = rs!nmkolektor
Piutang_DTU.lblurut = rs!urut
Piutang_DTU.txtjmlbayar = FormatNumber(rs!jmlbayar, 0)
Piutang_DTU.txtPPH23 = FormatNumber(rs!rpPPH23, 0)
Piutang_DTU.txtpotongan = FormatNumber(rs!potongan, 0)
Piutang_DTU.txtketerangan = rs!keterangan
Piutang_DTU.txttglbayar = rs!tglbayar
Piutang_DTU.CMBjenis.ListIndex = rs!trf
Piutang_DTU.lblkdbyrpiutang = rs!kdbyrpiutang

'Piutang_DTU.txtjmlbayar = CCur(lblsisa) + CCur(rs!jmlbayar) + CCur(rs!potongan)
Piutang_DTU.lblsisa_awal = FormatNumber(CCur(lblsisa) + CCur(rs!jmlbayar) + CCur(lblrpPPH23) + CCur(rs!potongan), 0)


  
Piutang_DTU.Show vbModal
 
End Sub


Private Sub hps()
On Error GoTo hell

KODE = 2
Call max


ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
If ms = vbYes Then
    sql = "delete from byrpiutangsewa where kdbyrpiutang ='" & rs!kdbyrpiutang & "'"
    con.Execute (sql)
    
    Piutang.TimerALL.Interval = 10
    TimerALL.Interval = 10
    
End If

         

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub









Private Sub nomer()
On Error GoTo hell

If LBLKODE = 1 Then
    sql = "select isnull(max(right(kdpobeli,4)),0) as xx from PObeli where Month(tglPObeli)='" & Month(txttglPO) & "'  and year(tglPObeli)='" & Year(txttglPO) & "' and kdgudang= '" & lblkdgudang & "'"
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "000" & a
                Case 2
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "00" & a
                Case 3
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "0" & a
                Case 4
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & a
        End Select
    
    Else
        txtnoPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
txtnoPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "0001"
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
 
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0


txttglPO = Date

TimerALL.Interval = 10
TimerNo.Interval = 10




End Sub




Private Sub TimerAll_Timer()
On Error Resume Next
Call all
Call jmlpiut


If KODE = 2 Then
rs.AbsolutePosition = lblpos
End If

 

TimerALL.Interval = 0

End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If


TimerNo.Interval = 0
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

