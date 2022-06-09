VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Kwitansi 
   BorderStyle     =   0  'None
   ClientHeight    =   10275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerPDF 
      Left            =   16245
      Top             =   945
   End
   Begin VB.ComboBox CMBBLN 
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
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   765
      Width           =   915
   End
   Begin VB.TextBox txttahun 
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
      TabIndex        =   10
      Text            =   "2017"
      Top             =   765
      Width           =   960
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
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   9675
      Width           =   1860
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
      Left            =   3915
      TabIndex        =   8
      Top             =   9675
      Width           =   2850
   End
   Begin VB.Timer TimerAll 
      Left            =   5625
      Top             =   4815
   End
   Begin VB.Timer TimerG 
      Left            =   6165
      Top             =   4815
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   11
      Top             =   675
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
   Begin Threed.SSCommand cmdPDF 
      Height          =   915
      Left            =   19395
      TabIndex        =   0
      ToolTipText     =   "Ubah"
      Top             =   1395
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
      Picture         =   "Kwitansi.frx":0000
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   19395
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   4230
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
      Picture         =   "Kwitansi.frx":31E7
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   19395
      TabIndex        =   4
      ToolTipText     =   "Cari Data"
      Top             =   5175
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
      Picture         =   "Kwitansi.frx":6363
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   0
      Left            =   19395
      TabIndex        =   1
      ToolTipText     =   "Cetak"
      Top             =   2340
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
      Picture         =   "Kwitansi.frx":9289
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   19395
      TabIndex        =   5
      ToolTipText     =   "Cetak Bentuk List"
      Top             =   6120
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
      Picture         =   "Kwitansi.frx":CCE6
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   19395
      TabIndex        =   2
      ToolTipText     =   "Ubah"
      Top             =   3285
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
      Picture         =   "Kwitansi.frx":1006C
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7800
      Left            =   180
      TabIndex        =   6
      Top             =   1260
      Width           =   19050
      _cx             =   33602
      _cy             =   13758
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
      BackColorAlternate=   16761087
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
      FormatString    =   $"Kwitansi.frx":13269
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TAGIHAN BLN :"
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
      Left            =   450
      TabIndex        =   17
      Top             =   810
      Width           =   1320
   End
   Begin VB.Label Label3 
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
      Left            =   2880
      TabIndex        =   16
      Top             =   810
      Width           =   780
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   19305
      Picture         =   "Kwitansi.frx":133E4
      Stretch         =   -1  'True
      Top             =   405
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6840
      Picture         =   "Kwitansi.frx":137A4
      Stretch         =   -1  'True
      Top             =   9630
      Width           =   420
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
      Left            =   2025
      TabIndex        =   15
      Top             =   9315
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1845
      Top             =   9270
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kwitansi Tagihan Sewa"
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
      Left            =   1215
      TabIndex        =   14
      Top             =   0
      Width           =   7395
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   10890
      Picture         =   "Kwitansi.frx":20654
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   555
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
      Left            =   10080
      TabIndex        =   13
      Top             =   9765
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   585
      TabIndex        =   12
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   10230
      Left            =   0
      Picture         =   "Kwitansi.frx":26EA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20445
   End
End
Attribute VB_Name = "Kwitansi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsX As ADODB.Recordset
Dim kategori As String
Dim KODE As Integer
Dim rsmax As ADODB.Recordset
Dim kata As String
Dim sqlX As String
Dim color As Long, flag As Byte

Private Sub Cetak()

Unload AR_Kwitansi

 ms = MsgBox("Cetak Dengan Stempel TSP ?", vbYesNo + vbQuestion, "Info")
    If ms = vbNo Then
        AR_Kwitansi.IMG_STEMPEL.Visible = False
        AR_Kwitansi.lbltgl_STEMPEL.Visible = False
        AR_Kwitansi.Image2.Visible = True
    Else
        AR_Kwitansi.IMG_STEMPEL.Visible = True
        AR_Kwitansi.lbltgl_STEMPEL.Visible = True
        AR_Kwitansi.Image2.Visible = False
    End If

If txtcari = "" Then
sqlX = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,c.nmbank,c.norek,a.tglcetak,c.atas_nama from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer left join bank c on b.kdbank=c.kdbank where " & kata & " and a.tahun=" & txttahun & " order by a.kdpiutang,a.kdcustomer"
Else
sqlX = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,c.nmbank,c.norek,a.tglcetak,c.atas_nama from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer left join bank c on b.kdbank=c.kdbank where " & kata & " and a.tahun=" & txttahun & " and " & kategori & " like '%" & txtcari & "%' order by a.kdpiutang,a.kdcustomer"

End If

Set rsX = con.Execute(sqlX)

With AR_Kwitansi.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_Kwitansi
.fldnokwitansi.DataField = "kdpiutang"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.flduang.DataField = "jmlpiutang"
.fldbln.DataField = "bln"
.FLDTHN.DataField = "tahun"
.fldunit.DataField = "unit"
.fldharga.DataField = "harga"
.fldjmlpiutang.DataField = "jmlpiutang"
.fldtglposting.DataField = "tglcetak"
.fldnorek.DataField = "norek"
.fldnmbank.DataField = "nmbank"
.fldAtas_nama.DataField = "atas_nama"


Select Case CLng(Month(rsX!tglcetak))
       Case 1
      .lbltgl_STEMPEL = "23 JAN " & Kwitansi.txttahun
       Case 2
      .lbltgl_STEMPEL = "23 FEB " & Kwitansi.txttahun
       Case 3
      .lbltgl_STEMPEL = "23 MAR " & Kwitansi.txttahun
       Case 4
      .lbltgl_STEMPEL = "23 APR " & Kwitansi.txttahun
       Case 5
      .lbltgl_STEMPEL = "23 MEI " & Kwitansi.txttahun
       Case 6
      .lbltgl_STEMPEL = "23 JUN " & Kwitansi.txttahun
       Case 7
      .lbltgl_STEMPEL = "23 JUL " & Kwitansi.txttahun
       Case 8
      .lbltgl_STEMPEL = "23 AGS " & Kwitansi.txttahun
       Case 9
      .lbltgl_STEMPEL = "23 SEP " & Kwitansi.txttahun
       Case 10
      .lbltgl_STEMPEL = "23 OKT " & Kwitansi.txttahun
       Case 11
      .lbltgl_STEMPEL = "23 NOV " & Kwitansi.txttahun
       Case 12
      .lbltgl_STEMPEL = "23 DES " & Kwitansi.txttahun

End Select

.Zoom = 140


AR_Kwitansi.Show vbModal


End With

End Sub

Private Sub Cetak1()

Unload AR_Kwitansi

 ms = MsgBox("Cetak Dengan Stempel TSP ?", vbYesNo + vbQuestion, "Info")
    If ms = vbNo Then
        AR_Kwitansi.IMG_STEMPEL.Visible = False
        AR_Kwitansi.lbltgl_STEMPEL.Visible = False
        AR_Kwitansi.Image2.Visible = True
    Else
        AR_Kwitansi.IMG_STEMPEL.Visible = True
        AR_Kwitansi.lbltgl_STEMPEL.Visible = True
        AR_Kwitansi.Image2.Visible = False
    End If
    

If txtcari = "" Then
sqlX = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,c.nmbank,c.norek,a.tglcetak,c.atas_nama from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer left join bank c on b.kdbank=c.kdbank where " & kata & " and a.tahun=" & txttahun & " order by a.kdpiutang,a.kdcustomer"
Else
sqlX = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,c.nmbank,c.norek,a.tglcetak,c.atas_nama from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer left join bank c on b.kdbank=c.kdbank where " & kata & " and a.tahun=" & txttahun & " and " & kategori & " like '%" & txtcari & "%' order by a.kdpiutang,a.kdcustomer"

End If



Set rsX = con.Execute(sqlX)

With AR_Kwitansi.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_Kwitansi
.fldnokwitansi.DataField = "kdpiutang"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.flduang.DataField = "jmlpiutang"
.fldbln.DataField = "bln"
.FLDTHN.DataField = "tahun"
.fldunit.DataField = "unit"
.fldharga.DataField = "harga"
.fldjmlpiutang.DataField = "jmlpiutang"
.fldtglposting.DataField = "tglcetak"
.fldnorek.DataField = "norek"
.fldnmbank.DataField = "nmbank"
.fldAtas_nama.DataField = "atas_nama"

Select Case CLng(Month(rsX!tglcetak))
       Case 1
      .lbltgl_STEMPEL = "23 JAN " & Kwitansi.txttahun
       Case 2
      .lbltgl_STEMPEL = "23 FEB " & Kwitansi.txttahun
       Case 3
      .lbltgl_STEMPEL = "23 MAR " & Kwitansi.txttahun
       Case 4
      .lbltgl_STEMPEL = "23 APR " & Kwitansi.txttahun
       Case 5
      .lbltgl_STEMPEL = "23 MEI " & Kwitansi.txttahun
       Case 6
      .lbltgl_STEMPEL = "23 JUN " & Kwitansi.txttahun
       Case 7
      .lbltgl_STEMPEL = "23 JUL " & Kwitansi.txttahun
       Case 8
      .lbltgl_STEMPEL = "23 AGS " & Kwitansi.txttahun
       Case 9
      .lbltgl_STEMPEL = "23 SEP " & Kwitansi.txttahun
       Case 10
      .lbltgl_STEMPEL = "23 OKT " & Kwitansi.txttahun
       Case 11
      .lbltgl_STEMPEL = "23 NOV " & Kwitansi.txttahun
       Case 12
      .lbltgl_STEMPEL = "23 DES " & Kwitansi.txttahun

End Select

.Zoom = 140

AR_Kwitansi.Show vbModal

TimerPdf.Interval = 10



End With

End Sub



Private Sub CMBBLN_Click()


Select Case CMBBLN.ListIndex
    Case 0
        kata = "a.bln <=12"
    Case 1
        kata = "a.bln =1"
    Case 2
        kata = "a.bln =2"
    Case 3
        kata = "a.bln =3"
    Case 4
        kata = "a.bln =4"
    Case 5
        kata = "a.bln =5"
    Case 6
        kata = "a.bln =6"
    Case 7
        kata = "a.bln =7"
    Case 8
        kata = "a.bln =8"
    Case 9
        kata = "a.bln =9"
    Case 10
        kata = "a.bln =10"
    Case 11
        kata = "a.bln =11"
    Case 12
        kata = "a.bln =12"
End Select
 
 
TimerAll.Interval = 10
 
End Sub

Private Sub CMBBLN_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub cmdPDF_Click()
Call Cetak1
End Sub

Private Sub cmdPDF_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub DataGrid1_DblClick()
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

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
    cmdT(1).Enabled = False
    cmdT(2).Enabled = False
    datagrid1.Enabled = False
    img1.Visible = True
    lbl1.Visible = True
Else
    cmdT(1).Enabled = True
    cmdT(2).Enabled = True
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

End Sub

Private Sub ubh()
Kwitansi_TU.LBLKODE = 2
lblpos = rs.AbsolutePosition
KODE = 2

Kwitansi_TU.lblKDpiutang = rs!kdpiutang
Kwitansi_TU.lblnmcustomer = rs!nmcustomer
Kwitansi_TU.lblalamat = rs!alamat
Kwitansi_TU.txttglcetak = rs!tglcetak


Kwitansi_TU.Show vbModal
End Sub

Private Sub hps()
On Error GoTo hell
KODE = 3
Call max
    ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = "delete from BARANG where kdBARANG='" & rs!kdbarang & "' "
        con.Execute (sql)
        
        TimerAll.Interval = 10
    Else
        Exit Sub
    End If


Exit Sub
hell:
MsgBox err.Description
End Sub


Private Sub ALL()
MousePointer = vbHourglass

If txtcari = "" Then
sql = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,a.tglcetak from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & txttahun & " order by a.kdpiutang,a.kdcustomer"
Else
sql = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,a.tglcetak from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & txttahun & " and " & kategori & " like '%" & txtcari & "%' order by a.kdpiutang,a.kdcustomer"


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

TimerAll.Interval = 10
End Sub

Private Sub CMBCARI_KeyPress(KeyAscii As Integer)
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
 Call ALL
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call Cetak
ElseIf Index = 2 Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If

ElseIf Index = 3 Then
txtcari = ""
Call ALL
ElseIf Index = 4 Then
txtcari = ""
    If txtcari.Enabled = True Then
    Me.Height = Me.Height - 1170

    txtcari.Enabled = False
    CMBCARI.Enabled = False
    Else
    Me.Height = Me.Height + 1170

    txtcari.Enabled = True
    CMBCARI.Enabled = True
    End If
ElseIf Index = 1 Then
Kwitansi_LIST.lblfrm = "KWITANSI"
Kwitansi_LIST.Show vbModal
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
 Call ALL
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub

Private Sub datagrid1_Click()
TimerG.Interval = 10
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
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call ALL
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub


Private Sub Form_Load()

GradientForm Me, 0

Me.Height = Me.Height - 1170


txttahun = Year(Date)

CMBBLN.AddItem "ALL"
CMBBLN.AddItem "1"
CMBBLN.AddItem "2"
CMBBLN.AddItem "3"
CMBBLN.AddItem "4"
CMBBLN.AddItem "5"
CMBBLN.AddItem "6"
CMBBLN.AddItem "7"
CMBBLN.AddItem "8"
CMBBLN.AddItem "9"
CMBBLN.AddItem "10"
CMBBLN.AddItem "11"
CMBBLN.AddItem "12"
CMBBLN.ListIndex = Month(Date)


CMBCARI.AddItem "NO KWITANSI"
CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "ALAMAT"

CMBCARI.ListIndex = 0



TimerAll.Interval = 10
End Sub

Private Sub TimerALL_Timer()
On Error Resume Next
Call ALL

If KODE = 2 Or KODE = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerAll.Interval = 0

End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub


Private Sub TimerPDF_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF

out2 = out2 + 1

Call save_out
pdf.filename = alamat_save & "\outfile" & CStr(out2) & ".pdf"
pdf.Export AR_Kwitansi.Pages

Call EX_PDF(Me)
TimerPdf.Interval = 0

Exit Sub
hell:
TimerPdf.Interval = 0
If out2 < 10 Then
cmdPDF_Click
End If

End Sub



Private Sub TXTCARI_Change()
TimerAll.Interval = 10
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
'ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
' Call tbh
'ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
' If rs.RecordCount <> 0 Then
' Call ubh
' End If
'ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
' If rs.RecordCount <> 0 Then
' Call hps
' End If
'ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
' Call all
End If
End Sub




Private Sub txttahun_Change()
TimerAll.Interval = 10
End Sub

Private Sub txttahun_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
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

Private Sub txtharga_LostFocus()
On Error GoTo hell

txttahun = Format(txttahun, "####0")


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txttahun.SetFocus

End Sub
