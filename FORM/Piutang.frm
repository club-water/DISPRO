VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Piutang 
   BorderStyle     =   0  'None
   ClientHeight    =   10260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtR 
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
      Left            =   17460
      TabIndex        =   27
      Text            =   "100"
      Top             =   315
      Width           =   735
   End
   Begin VB.CheckBox ChkR 
      BackColor       =   &H00000000&
      Caption         =   "TAMPILKAN :"
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
      Height          =   330
      Left            =   15885
      MaskColor       =   &H00000000&
      TabIndex        =   26
      Top             =   315
      Value           =   1  'Checked
      Width           =   1545
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
      Left            =   12510
      TabIndex        =   13
      Top             =   855
      Width           =   2535
   End
   Begin VB.ComboBox CMBjenis 
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
      Left            =   9855
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   855
      Width           =   1275
   End
   Begin VB.TextBox txttglbayar 
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
      Left            =   1215
      TabIndex        =   10
      Top             =   855
      Width           =   1590
   End
   Begin VB.Timer TimerG 
      Left            =   6165
      Top             =   4815
   End
   Begin VB.Timer TimerAll 
      Left            =   5625
      Top             =   4815
   End
   Begin VB.TextBox TXTCARI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   7
      Top             =   9585
      Width           =   2850
   End
   Begin VB.ComboBox CMBCARI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   9585
      Width           =   1860
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   360
      TabIndex        =   14
      Top             =   675
      Width           =   18870
      _Version        =   524288
      _ExtentX        =   33285
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
      TabIndex        =   0
      ToolTipText     =   "Tambah"
      Top             =   1305
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
      Picture         =   "Piutang.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   19395
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   2250
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
      Picture         =   "Piutang.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   13590
      TabIndex        =   15
      ToolTipText     =   "Hapus"
      Top             =   8325
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
      Picture         =   "Piutang.frx":594B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   19395
      TabIndex        =   2
      ToolTipText     =   "Refresh"
      Top             =   3195
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
      Picture         =   "Piutang.frx":89E4
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   19395
      TabIndex        =   4
      ToolTipText     =   "Cari Data"
      Top             =   5085
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
      Picture         =   "Piutang.frx":BB60
      ButtonStyle     =   4
   End
   Begin Threed.SSOption Oblunas 
      Height          =   330
      Left            =   135
      TabIndex        =   8
      Top             =   8595
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "Belum Lunas"
   End
   Begin Threed.SSOption Olunas 
      Height          =   330
      Left            =   1575
      TabIndex        =   9
      Top             =   8595
      Width           =   780
      _ExtentX        =   1376
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
      Caption         =   "Lunas"
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   7650
      TabIndex        =   11
      ToolTipText     =   "Simpan"
      Top             =   810
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
      Picture         =   "Piutang.frx":EA86
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   5
      Left            =   19395
      TabIndex        =   3
      ToolTipText     =   "Tampilkan Total Sisa Piutang"
      Top             =   4140
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
      Picture         =   "Piutang.frx":112B8
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7260
      Left            =   90
      TabIndex        =   5
      Top             =   1305
      Width           =   19050
      _cx             =   33602
      _cy             =   12806
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
      BackColorAlternate=   12632319
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Piutang.frx":14005
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RECORD"
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
      Left            =   18225
      TabIndex        =   28
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label6 
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
      Left            =   11295
      TabIndex        =   25
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "JNS PEMBAYARAN :"
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
      Left            =   8280
      TabIndex        =   24
      Top             =   900
      Width           =   1545
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
      Left            =   4725
      TabIndex        =   23
      Top             =   855
      Width           =   2940
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
      Left            =   3825
      TabIndex        =   22
      Top             =   855
      Width           =   870
   End
   Begin VB.Label Label5 
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
      Left            =   2880
      TabIndex        =   21
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL BAYAR :"
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
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   90
      TabIndex        =   19
      Top             =   9180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA TIDAK ADA"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9180
      TabIndex        =   18
      Top             =   9585
      Width           =   2220
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   9990
      Picture         =   "Piutang.frx":141A4
      Stretch         =   -1  'True
      Top             =   9090
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Piutang Sewa"
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
      Left            =   1260
      TabIndex        =   17
      Top             =   0
      Width           =   7395
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1350
      Top             =   9180
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Pencarian"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1530
      TabIndex        =   16
      Top             =   9225
      Width           =   4560
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6345
      Picture         =   "Piutang.frx":1A9F6
      Stretch         =   -1  'True
      Top             =   9540
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   19305
      Picture         =   "Piutang.frx":278A6
      Stretch         =   -1  'True
      Top             =   405
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   10230
      Left            =   0
      Picture         =   "Piutang.frx":27C66
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20445
   End
End
Attribute VB_Name = "Piutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori, sqlcek As String
Dim kode As Integer
Dim rsmax As ADODB.Recordset
Dim rscek As ADODB.Recordset
Dim rsL As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim sqlL As String
Dim l As Integer
Dim sqlJ, sqlJ1, sqlJ2 As String
Dim color As Long, flag As Byte
Dim rpPPH As Currency
Dim rpPPH_X As Currency
Dim rsC As ADODB.Recordset

Private Sub all_jml()
'On Error Resume Next
sqlJ2 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar, sum(rpPPH23) as rppph23,sum(potongan) as potongan," & vbCrLf & _
       "sum(jmlpiutang - jmlbayar - rpPPH23 - potongan) as sisa from (" & vbCrLf & _
       "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar, 0 as rpPPH23,0 as potongan from piutangsewa" & vbCrLf & _
       "Union" & vbCrLf & _
       "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(rpPPH23) as rpPPH23,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
       "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"

If ChkR.Value = 0 Then
    If TXTCARI = "" Then
        If Oblunas.Value = True Then
        sqlJ1 = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 "
        Else
        sqlJ1 = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 "
        End If
    Else
        If Oblunas.Value = True Then
        sqlJ1 = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 and " & kategori & " like '%" & TXTCARI & "%' "
        Else
        sqlJ1 = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 and " & kategori & " like '%" & TXTCARI & "%'"
        End If
    
    End If
Else
    If TXTCARI = "" Then
        If Oblunas.Value = True Then
        sqlJ1 = "select top " & CLng(txtR) & "  a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 "
        Else
        sqlJ1 = "select top " & CLng(txtR) & " a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 "
        End If
    Else
        If Oblunas.Value = True Then
        sqlJ1 = "select top " & CLng(txtR) & " a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 and " & kategori & " like '%" & TXTCARI & "%' "
        Else
        sqlJ1 = "select top " & CLng(txtR) & " a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun from (" & sqlJ2 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 and " & kategori & " like '%" & TXTCARI & "%'"
        End If
    
    End If
End If

sqlJ3 = "select '1' as kode,* from (" & sqlJ1 & ") a"

sqlJ = "select kode,sum(sisa) as sisa from (" & sqlJ3 & ") a group by kode"

Set rsJ = con.Execute(sqlJ)
MsgBox "Sisa Piutang Sewa = Rp " & Format(rsJ!sisa, "#,###0") & " ,-", vbInformation, "Info !!"


End Sub

Private Sub lunas()
On Error GoTo hell

If cmdT(1).Enabled = True Then

    Call Cek_tglOD
    If CDate(txttglbayar) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
        MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
        Exit Sub
    ElseIf lblkdkolektor = "" Then
        MsgBox "Tolong isi dulu kolektornya !!", vbCritical, "Error !"
        Exit Sub
    Else
    
        ms = MsgBox("Apakah anda ingin melunasi Piutang ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
                 
             sqlL = "select isnull(max(urut),0) as urut from byrpiutangsewa where kdpiutang='" & rs!kdpiutang & "' "
             Set rsL = con.Execute(sqlL)
             
             
             
             If rsL.RecordCount <> 0 Then
             l = CLng(rsL!urut) + 1
             Else
             l = 1
             End If
        
         If ms = vbYes Then
             If cmdT(2).Enabled = True Then
                 Call max
                 kode = 2
                 
                 sqlC = "select * from customer where kdcustomer ='" & rs!kdcustomer & "' and PPH23=1"
                 Set rsC = con.Execute(sqlC)
                 
                 If rsC.RecordCount <> 0 Then
                 
                 rpPPH_X = FormatNumber((CCur(rs!sisa) / 1.1) * 0.02, 0)
                 
                 ms = InputBox("Masukkan PPH 23 !", "Potongan PPH 23", rpPPH_X)
                    If ms = "" Then
                    rpPPH = 0
                    Else
                    rpPPH = FormatNumber(ms, 0)
                    End If
                 Else
                 rpPPH = 0
                 End If
                 
                 
                 sql = "insert into byrpiutangsewa values ('" & CStr(l) & rs!kdpiutang & "'," & l & " ,'" & Format(txttglbayar, "yyyy-MM-dd") & "','" & rs!kdcustomer & "','" & lblkdkolektor & "'," & CLng(rs!sisa) - rpPPH & "," & rpPPH & ",0,'" & txtketerangan & "','" & rs!kdpiutang & "'," & CMBjenis.ListIndex & ") "
                 con.Execute (sql)
                 
                 TimerALL.Interval = 10
             End If
         Else
             Exit Sub
         End If
     End If
End If

Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"
End Sub


Private Sub cek_dalem()
sqlcek = "select * from PObeli_D where kdPObeli='" & rs!kdPObeli & "'"
Set rscek = con.Execute(sqlcek)
End Sub

Private Sub ChkR_Click()
TimerALL.Interval = 10

If ChkR.Value = 0 Then
txtR.Enabled = False
Else
txtR.Enabled = True
End If

End Sub

Private Sub ChkR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub CMBjenis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdBR_Click()
Kolektor_BR.LBLKODE = "PIUTANG"
Kolektor_BR.Show vbModal

End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
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



'untuk set cursor pada saat dihapus
Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub


Private Sub tbl()
If rs.RecordCount = 0 Then
    cmdT(0).Enabled = False
    cmdT(1).Enabled = False
    datagrid1.Enabled = False
    img1.Visible = True
    lbl1.Visible = True
Else
    cmdT(0).Enabled = True
    
    If Oblunas.Value = True Then
    cmdT(1).Enabled = True
    Else
    cmdT(1).Enabled = False
    End If
    
    datagrid1.Enabled = True
    img1.Visible = False
    lbl1.Visible = False
End If
End Sub


Private Sub LG()
On Error GoTo hell

Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()
Piutang_D.LBLKODE = 2
lblpos = rs.AbsolutePosition
kode = 2

Piutang_D.txtkdPiutang = rs!kdpiutang
Piutang_D.lblalamat = rs!alamat
Piutang_D.lblkdcustomer = rs!kdcustomer
Piutang_D.lblnmcustomer = rs!nmcustomer
Piutang_D.lbltglposting = rs!tglposting
Piutang_D.lbltahun = rs!tahun
Piutang_D.lbljmlpiutang = Format(rs!jmlpiutang, "#,###0")

Select Case rs!bln
Case 1
Piutang_D.lblbln = "JANUARI"

Case 2
Piutang_D.lblbln = "FEBRUARI"

Case 3
Piutang_D.lblbln = "MARET"

Case 4
Piutang_D.lblbln = "APRIL"

Case 5
Piutang_D.lblbln = "MEI"

Case 6
Piutang_D.lblbln = "JUNI"

Case 7
Piutang_D.lblbln = "JULI"

Case 8
Piutang_D.lblbln = "AGUSTUS"

Case 9
Piutang_D.lblbln = "SEPTEMBER"

Case 10
Piutang_D.lblbln = "OKTOBER"

Case 11
Piutang_D.lblbln = "NOVEMBER"

Case 12
Piutang_D.lblbln = "DESEMBER"

End Select


'Piutang_D.txtketerangan = rs!keterangan
'Piutang_D.txtnoPP = rs!nopp
'Piutang_D.txttglkembali = rs!tglpengembalian
'Piutang_D.CMBStatus.Text = rs!Status
'
'Piutang_D.txttglpinjam.Enabled = False
'Piutang_D.cmdBR.Enabled = False
'



Piutang_D.Show vbModal
End Sub

Private Sub ubh()

'Piutang_D.lblkode = 2
'lblpos = rs.AbsolutePosition
'kode = 2
'
'Piutang_D.txtkdPO = rs!kdPO
'Piutang_D.lblkdgudang = rs!kdgudang
'Piutang_D.lblnmgudang = rs!nmgudang
'Piutang_D.lblkdcustomer = rs!kdcustomer
'Piutang_D.lblnmcustomer = rs!nmcustomer
'Piutang_D.lblalamat = rs!alamat
'Piutang_D.txttglpinjam = rs!tglpinjam
'Piutang_D.lblKDPinjam = rs!kdPinjam
'Piutang_D.txtketerangan = rs!keterangan
'Piutang_D.txtnoPP = rs!nopp
'Piutang_D.txttglkembali = rs!tglpengembalian
'Piutang_D.CMBStatus.Text = rs!Status
'
'Piutang_D.txttglpinjam.Enabled = False
'Piutang_D.cmdBR.Enabled = False




'Piutang_D.Show vbModal
End Sub

Private Sub hps()
'On Error GoTo hell
'kode = 3
'Call max
'
'
'    Call cek_dalem
'    If rscek.RecordCount = 0 Then
'        MsgBox "Data Tidak dapat dihapus, karena Detail PO masih ada", vbCritical, "Error !"
'        Exit Sub
'
'    Else
'        ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
'        If ms = vbYes Then
'            sql = "delete from PObeli where kdpobeli='" & rs!kdPObeli & "' "
'            con.Execute (sql)
'
'            TimerAll.Interval = 10
'        Else
'            Exit Sub
'        End If
'    End If
'
'Exit Sub
'hell:
'MsgBox err.Description
End Sub


Private Sub all()

MousePointer = vbHourglass

sql1 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar,sum(rpPPH23) as rpPPH23,sum(potongan) as potongan," & vbCrLf & _
       "sum(jmlpiutang - jmlbayar - rpPPH23 - potongan) as sisa from (" & vbCrLf & _
       "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar,0 as rpPPH23,0 as potongan from piutangsewa" & vbCrLf & _
       "Union" & vbCrLf & _
       "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(rpPPH23) as rpPPH23,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
       "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"


If ChkR.Value = 0 Then
    If TXTCARI = "" Then
        If Oblunas.Value = True Then
        sql = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 order by c.tahun desc,c.bln desc"
        Else
        sql = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 order by c.tahun desc,c.bln desc"
        End If
    Else
        If Oblunas.Value = True Then
        sql = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 and " & kategori & " like '%" & TXTCARI & "%' order by c.tahun desc,c.bln desc"
        Else
        sql = "select a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 and " & kategori & " like '%" & TXTCARI & "%'order by c.tahun desc,c.bln desc"
        End If
    
    End If

Else
    If TXTCARI = "" Then
        If Oblunas.Value = True Then
        sql = "select top " & CLng(txtR) & " a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 order by c.tahun desc,c.bln desc"
        Else
        sql = "select top " & CLng(txtR) & "a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 order by c.tahun desc,c.bln desc"
        End If
    Else
        If Oblunas.Value = True Then
        sql = "select top " & CLng(txtR) & " a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 and " & kategori & " like '%" & TXTCARI & "%' order by c.tahun desc,c.bln desc"
        Else
        sql = "select top " & CLng(txtR) & " a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlpiutang,a.jmlbayar,a.rpPPH23,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,TT1 = case when c.TT=1 then 'X' else '' end  from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
              "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa = 0 and " & kategori & " like '%" & TXTCARI & "%'order by c.tahun desc,c.bln desc"
        End If
    
    End If
End If

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

Call LG

MousePointer = vbDefault
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "a.kdpiutang"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "b.nmcustomer"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "b.alamat"
End If

TimerALL.Interval = 10
End Sub

Private Sub CMBCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("l") Or KeyAscii = Asc("L") Then
 If rs.RecordCount <> 0 Then
 Call lunas
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
     If rs.RecordCount <> 0 Then
     Call lunas
     End If
ElseIf Index = 2 Then
     If rs.RecordCount <> 0 Then
     Call hps
     End If
ElseIf Index = 3 Then
TXTCARI = ""
Call all
ElseIf Index = 4 Then
TXTCARI = ""
    If TXTCARI.Enabled = True Then
    Me.Height = Me.Height - 1170

    TXTCARI.Enabled = False
    CMBCARI.Enabled = False
    Else
    Me.Height = Me.Height + 1170

    TXTCARI.Enabled = True
    CMBCARI.Enabled = True
    End If
ElseIf Index = 5 Then
Call all_jml
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
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
ElseIf KeyAscii = Asc("j") Or KeyAscii = Asc("J") Then
 Call all_jml
 
End If
End Sub

Private Sub datagrid1_Click()
TimerG.Interval = 10
End Sub

Private Sub datagrid1_DblClick()
 If rs.RecordCount <> 0 Then
 Call tbh
 End If

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
ElseIf KeyAscii = Asc("l") Or KeyAscii = Asc("L") Then
 If rs.RecordCount <> 0 Then
 Call lunas
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
ElseIf KeyAscii = Asc("j") Or KeyAscii = Asc("J") Then
 Call all_jml

End If
End Sub


Private Sub Form_Load()

GradientForm Me, 0

Me.Top = Screen.Height / 3
Me.Height = Me.Height - 1170

Oblunas.Value = True
txttglbayar = Date

CMBjenis.AddItem "TUNAI"
CMBjenis.AddItem "TRANSFER"
CMBjenis.ListIndex = 0

CMBCARI.AddItem "NO KWITANSI"
CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "ALAMAT"
CMBCARI.ListIndex = 0

Call nul(lblkdkolektor)
Call nul(lblnmkolektor)




TimerALL.Interval = 10
End Sub

Private Sub lblkdkolektor_Change()
Call nul(lblkdkolektor)
End Sub

Private Sub lblnmkolektor_Change()
Call nul(lblnmkolektor)
End Sub

Private Sub Oblunas_Click(Value As Integer)
cmdT(1).Enabled = True
TimerALL.Interval = 10
End Sub

Private Sub Oblunas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Olunas_Click(Value As Integer)
cmdT(1).Enabled = False
TimerALL.Interval = 10
End Sub

Private Sub Olunas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub TimerAll_Timer()

On Error Resume Next
Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerALL.Interval = 0


End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerALL.Interval = 10
End Sub

Private Sub TXTCARI_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTCARI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    TimerG.Interval = 10
    Else
    SendKeys vbTab
    End If
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub txtketerangan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtketerangan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
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

Private Sub txtR_Change()
Call nul(txtR)
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii <> vbKeyBack Then

    cekTBL = InStr("1234567890.,", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txtR_LostFocus()
On Error GoTo hell

txtR = FormatNumber(txtR, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtR.SetFocus

End Sub

Private Sub txttglbayar_Change()
Call nul(txttglbayar)
End Sub

Private Sub txttglbayar_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglbayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglbayar_KeyPress(KeyAscii As Integer)
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

Private Sub txttglbayar_LostFocus()
On Error GoTo hell

txttglbayar = FormatDateTime(txttglbayar, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglbayar.SetFocus

End Sub

