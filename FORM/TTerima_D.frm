VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form TTerima_D 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglTT 
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
      TabIndex        =   0
      Top             =   945
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
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   6
      Top             =   720
      Width           =   9690
      _Version        =   524288
      _ExtentX        =   17092
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
      TabIndex        =   7
      Top             =   2205
      Width           =   9780
      _Version        =   524288
      _ExtentX        =   17251
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
      Left            =   10350
      TabIndex        =   2
      ToolTipText     =   "Tambah"
      Top             =   2295
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
      Picture         =   "TTerima_D.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   9000
      TabIndex        =   8
      ToolTipText     =   "Ubah"
      Top             =   6075
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
      Picture         =   "TTerima_D.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   10350
      TabIndex        =   3
      ToolTipText     =   "Hapus"
      Top             =   3240
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
      Picture         =   "TTerima_D.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   10350
      TabIndex        =   4
      ToolTipText     =   "Refresh"
      Top             =   4185
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
      Picture         =   "TTerima_D.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   7920
      TabIndex        =   9
      ToolTipText     =   "Cetak"
      Top             =   6075
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
      Picture         =   "TTerima_D.frx":C086
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   945
      TabIndex        =   10
      Top             =   6750
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
      Picture         =   "TTerima_D.frx":FAE3
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   9540
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   1260
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
      Picture         =   "TTerima_D.frx":16345
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   3750
      Left            =   90
      TabIndex        =   5
      Top             =   2295
      Width           =   10050
      _cx             =   17727
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"TTerima_D.frx":18B77
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
      Left            =   1215
      TabIndex        =   18
      Top             =   1305
      Width           =   1140
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
      Left            =   180
      TabIndex        =   17
      Top             =   1350
      Width           =   1005
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
      Left            =   2385
      TabIndex        =   16
      Top             =   1305
      Width           =   7125
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
      Left            =   1215
      TabIndex        =   15
      Top             =   1665
      Width           =   8835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanda Terima"
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
      TabIndex        =   14
      Top             =   45
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL TT :"
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
      TabIndex        =   13
      Top             =   990
      Width           =   735
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   4095
      TabIndex        =   12
      Top             =   7515
      Width           =   1545
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   6615
      TabIndex        =   11
      Top             =   7380
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   7260
      Left            =   -90
      Picture         =   "TTerima_D.frx":18C8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11490
   End
End
Attribute VB_Name = "TTerima_D"
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
Dim sqlA As String
Dim color As Long, flag As Byte
Dim rscek As ADODB.Recordset

Private Sub cek_dalem()
sqlcek = "select * from PObeli_D where kdPObeli='" & txtkdPO & "'"
Set rscek = con.Execute(sqlcek)
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

Unload AR_PObeli

sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from pobeli_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdpobeli='" & txtkdPO & "' order by a.kdbarang"

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
.lbltglPO = Format(txttglTT, "dd/MM/yyyy")
.lbljudul1 = "CUSTOMER"

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
    txttglTT.Enabled = True
    cmdBR.Enabled = True
    datagrid1.Enabled = False
    cmdT(2).Enabled = False
    txttglTT.SetFocus
    

Else
    txttglTT.Enabled = False
    cmdBR.Enabled = False
    datagrid1.Enabled = True
    cmdT(2).Enabled = True
    datagrid1.SetFocus
End If
End Sub


Private Sub LG()
'On Error GoTo hell
'
'With datagrid1.Columns(0)
'.Width = 120
'.Caption = "NO KWITANSI"
'.Alignment = dbgCenter
'End With
'
'With datagrid1.Columns(1)
'.Caption = "BLN"
'.Width = 40
'.Alignment = dbgCenter
'End With
'
'With datagrid1.Columns(2)
'.Caption = "TAHUN"
'.Width = 60
'.Alignment = dbgCenter
'End With
'
'With datagrid1.Columns(3)
'.Caption = "kdcustomer"
'.Width = 0
'.Alignment = dbgCenter
'End With
'
'With datagrid1.Columns(4)
'.Caption = "JML PIUTANG"
'.Width = 100
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'End With
'
'With datagrid1.Columns(5)
'.Caption = "JML BAYAR"
'.Width = 100
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'End With
'
'With datagrid1.Columns(6)
'.Caption = "POTONGAN"
'.Width = 100
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'End With
'
'
'With datagrid1.Columns(7)
'.Caption = "SISA PIUTANG"
'.Width = 100
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'End With
'
'
Call tbl
'
'Exit Sub
'hell:
End Sub


Private Sub all()
sql1 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
       "sum(jmlpiutang - jmlbayar - potongan) as sisa from (" & vbCrLf & _
       "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar,0 as potongan from piutangsewa" & vbCrLf & _
       "Union" & vbCrLf & _
       "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
       "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"


sql = "select a.kdpiutang,c.bln,c.tahun,a.kdcustomer,a.jmlpiutang,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a " & vbCrLf & _
      "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join Tanda_terima b on a.kdpiutang=b.kdpiutang where b.tgltt='" & Format(txttglTT, "yyyy/MM/dd") & "' and a.kdcustomer='" & lblkdcustomer & "' order by c.tahun,c.bln"
      
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


Call LG
End Sub



Private Sub tbh()
If lblkdcustomer = "" Or txttglTT = "" Then
    MsgBox "Header Belum Lengkap !!", vbCritical, "Error !"
    Exit Sub
Else
    Piutang_BR.Show vbModal
End If

End Sub


Private Sub ubh()
If rs!Ubeli <> 0 Then
    MsgBox "Data Tidak Dapat diubah karena sudah ada Pembelian !", vbCritical, "Error !"
    Exit Sub
Else
    PObeli_DTU.lblkode = 2
    
    
    lblpos = rs.AbsolutePosition
    KODE = 2
    
    PObeli_DTU.lblkdPObeli_d = rs!kdPObeli_d
    
    PObeli_DTU.lblkdbarang = rs!kdbarang
    PObeli_DTU.lblnmbarang = rs!nmbarang
    PObeli_DTU.lblsatuan = rs!satuan
    PObeli_DTU.txtunit = FormatNumber(rs!unit, 0)
    PObeli_DTU.txtketerangan = rs!keterangan
    PObeli_DTU.cmdBR.Enabled = False
    
      
    PObeli_DTU.Show vbModal
End If
End Sub


Private Sub hps()
On Error GoTo hell

If rs!sisa = 0 Then
    MsgBox "Data Tidak Dapat dihapus karena piutang sudah Lunas !", vbCritical, "Error !"
    Exit Sub
Else

    KODE = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        
        sql = "update piutangsewa set tt=0 where kdpiutang ='" & rs!kdpiutang & "'"
        con.Execute (sql)
        
        sql = "delete from tanda_terima where kdpiutang ='" & rs!kdpiutang & "'"
        con.Execute (sql)
        TimerAll.Interval = 10
        TTerima.TimerAll.Interval = 10
    End If

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub











Private Sub cmdBR_Click()
Customer_br.lblkode = "TTERIMA_D"
Customer_br.Show vbModal

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
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
 
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0


txttglTT = Date

TimerAll.Interval = 10



Call nul(lblkdcustomer)
Call nul(lblnmcustomer)
Call nul(lblalamat)


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

Private Sub lblkdcustomer_Change()
Call nul(lblkdcustomer)
End Sub

Private Sub lblnmcustomer_Change()
Call nul(lblnmcustomer)
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all


If KODE = 2 Then
rs.AbsolutePosition = lblpos
End If

 

TimerAll.Interval = 0


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

Private Sub txttglTT_Change()
Call nul(txttglTT)

End Sub

Private Sub txttglTT_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglTT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txttglTT_KeyPress(KeyAscii As Integer)
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

Private Sub txttglTT_LostFocus()
On Error GoTo hell

txttglTT = FormatDateTime(txttglTT, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglTT.SetFocus

End Sub




