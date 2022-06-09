VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Klaim 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19020
   LinkTopic       =   "Form2"
   ScaleHeight     =   9450
   ScaleWidth      =   19020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   8
      Top             =   8775
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
      Left            =   3420
      TabIndex        =   9
      Top             =   8775
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
      TabIndex        =   10
      Top             =   675
      Width           =   17565
      _Version        =   524288
      _ExtentX        =   30983
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
      Left            =   17955
      TabIndex        =   1
      ToolTipText     =   "Tambah"
      Top             =   1125
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
      Picture         =   "Klaim.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   14670
      TabIndex        =   11
      ToolTipText     =   "Ubah"
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
      Picture         =   "Klaim.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   17955
      TabIndex        =   5
      ToolTipText     =   "Detail Barang yg di Klaim"
      Top             =   4905
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
      Picture         =   "Klaim.frx":594B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   17955
      TabIndex        =   2
      ToolTipText     =   "Refresh"
      Top             =   2070
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
      Picture         =   "Klaim.frx":6BC99
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   17955
      TabIndex        =   4
      ToolTipText     =   "Cari Data"
      Top             =   3960
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
      Picture         =   "Klaim.frx":6EE15
      ButtonStyle     =   4
   End
   Begin Threed.SSOption Oblunas 
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   7875
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
      Left            =   1620
      TabIndex        =   7
      Top             =   7875
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
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   5
      Left            =   17955
      TabIndex        =   3
      ToolTipText     =   "Tampilkan Total Sisa Piutang"
      Top             =   3015
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
      Picture         =   "Klaim.frx":71D3B
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   6810
      Left            =   90
      TabIndex        =   0
      Top             =   1035
      Width           =   17790
      _cx             =   31380
      _cy             =   12012
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
      BackColorAlternate=   12632256
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
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
      FormatString    =   $"Klaim.frx":74A88
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
      Editable        =   2
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
   Begin VB.Image Image2 
      Height          =   345
      Left            =   17955
      Picture         =   "Klaim.frx":74BD8
      Stretch         =   -1  'True
      Top             =   315
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6345
      Picture         =   "Klaim.frx":74F98
      Stretch         =   -1  'True
      Top             =   8730
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
      Left            =   1530
      TabIndex        =   15
      Top             =   8415
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1350
      Top             =   8370
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Klaim Barang Hilang"
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
      Left            =   1035
      TabIndex        =   14
      Top             =   0
      Width           =   7395
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   10395
      Picture         =   "Klaim.frx":81E48
      Stretch         =   -1  'True
      Top             =   8370
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
      Left            =   9585
      TabIndex        =   13
      Top             =   8865
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   9180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   9465
      Left            =   0
      Picture         =   "Klaim.frx":8869A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19005
   End
End
Attribute VB_Name = "Klaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori, sqlcek As String
Dim KODE As Integer
Dim rsmax As ADODB.Recordset
Dim rscek As ADODB.Recordset
Dim rsL As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim sqlL As String
Dim l As Integer
Dim sqlJ, sqlJ1, sqlJ2 As String
Dim color As Long, flag As Byte

Private Sub all_jml()

sqlX1 = "select kdklaim, kdcustomer,sum(jmlklaim) as jmlklaim, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
       "sum(jmlklaim - jmlbayar - potongan) as sisa from (" & vbCrLf & _
       "select 'a' as kode,kdklaim,kdcustomer,jmlklaim, 0 as jmlbayar,0 as potongan from klaim_hilang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select 'b' as kode,kdklaim,kdcustomer,0 as jmlklaim,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrklaim" & vbCrLf & _
       "group by kdklaim,kdcustomer ) a group by kdklaim, kdcustomer"


If txtcari = "" Then
    If Oblunas.Value = True Then
    sqlJ1 = "select '1' as kode,a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sqlX1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
            "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa <> 0 "
    Else
    sqlJ1 = "select '1' as kode,a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sqlX1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
           "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa = 0 "
    End If
Else
    If Oblunas.Value = True Then
    sqlJ1 = "select '1' as kode,a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sqlX1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
            "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa <> 0 and " & kategori & " like '%" & txtcari & "%' "
    
    Else
    sqlJ1 = "select '1' as kode,a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat_tgh as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sqlX1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          " left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa = 0 and " & kategori & " like '%" & txtcari & "%' "
    
    End If

End If

sqlJ = "select kode,sum(sisa) as sisa from (" & sqlJ1 & ") a group by kode"

Set rsJ = con.Execute(sqlJ)
MsgBox "Sisa Klaim Sewa = Rp " & Format(rsJ!sisa, "#,###0") & " ,-", vbInformation, "Info !!"


End Sub

Private Sub lunas()
'On Error GoTo hell
'
'
'If cmdT(1).Enabled = True Then
'
'    If lblkdkolektor = "" Then
'        MsgBox "Tolong isi dulu kolektornya !!", vbCritical, "Error !"
'        Exit Sub
'    Else
'
'        ms = MsgBox("Apakah anda ingin melunasi Klaim ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
'
'             sqlL = "select isnull(max(urut),0) as urut from byrKlaimsewa where kdKlaim='" & rs!kdKlaim & "' "
'             Set rsL = con.Execute(sqlL)
'
'
'
'             If rsL.RecordCount <> 0 Then
'             l = CLng(rsL!urut) + 1
'             Else
'             l = 1
'             End If
'
'         If ms = vbYes Then
'             If cmdT(2).Enabled = True Then
'                 Call max
'                 KODE = 2
'
'                 sql = "insert into byrKlaimsewa values ('" & CStr(l) & rs!kdKlaim & "'," & l & " ,'" & Format(txttglbayar, "yyyy-MM-dd") & "','" & rs!kdcustomer & "','" & lblkdkolektor & "'," & CLng(rs!SISA) & ",0,'" & txtketerangan & "','" & rs!kdKlaim & "'," & CMBjenis.ListIndex & ") "
'                 con.Execute (sql)
'
'                 TimerAll.Interval = 10
'             End If
'         Else
'             Exit Sub
'         End If
'     End If
'End If
'
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !!"
End Sub


Private Sub cek_dalem()
sqlcek = "select * from PObeli_D where kdPObeli='" & rs!kdPObeli & "'"
Set rscek = con.Execute(sqlcek)
End Sub

Private Sub CMBjenis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
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
Klaim_D.LBLKODE = 2
lblpos = rs.AbsolutePosition
KODE = 2

Klaim_D.txtkdklaim = rs!kdklaim
Klaim_D.lblalamat = rs!alamat
Klaim_D.lblkdcustomer = rs!kdcustomer
Klaim_D.lblnmcustomer = rs!nmcustomer
Klaim_D.lbltglklaim = rs!tglKlaim
Klaim_D.lbljmlklaim = Format(rs!jmlKlaim, "#,###0")


Klaim_D.Show vbModal
End Sub

Private Sub ubh()

'Klaim_D.lblkode = 2
'lblpos = rs.AbsolutePosition
'kode = 2
'
'Klaim_D.txtkdPO = rs!kdPO
'Klaim_D.lblkdgudang = rs!kdgudang
'Klaim_D.lblnmgudang = rs!nmgudang
'Klaim_D.lblkdcustomer = rs!kdcustomer
'Klaim_D.lblnmcustomer = rs!nmcustomer
'Klaim_D.lblalamat = rs!alamat
'Klaim_D.txttglpinjam = rs!tglpinjam
'Klaim_D.lblKDPinjam = rs!kdPinjam
'Klaim_D.txtketerangan = rs!keterangan
'Klaim_D.txtnoPP = rs!nopp
'Klaim_D.txttglkembali = rs!tglpengembalian
'Klaim_D.CMBStatus.Text = rs!Status
'
'Klaim_D.txttglpinjam.Enabled = False
'Klaim_D.cmdBR.Enabled = False




'Klaim_D.Show vbModal
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


Private Sub ALL()

sql1 = "select kdklaim, kdcustomer,sum(jmlklaim) as jmlklaim, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
       "sum(jmlklaim - jmlbayar - potongan) as sisa from (" & vbCrLf & _
       "select 'a' as kode,kdklaim,kdcustomer,jmlklaim, 0 as jmlbayar,0 as potongan from klaim_hilang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select 'b' as kode,kdklaim,kdcustomer,0 as jmlklaim,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrklaim" & vbCrLf & _
       "group by kdklaim,kdcustomer ) a group by kdklaim, kdcustomer"


If txtcari = "" Then
    If Oblunas.Value = True Then
    sql = "select a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa <> 0 order by c.tglklaim"
    Else
    sql = "select a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa = 0 order by c.tglklaim"
    End If
Else
    If Oblunas.Value = True Then
    sql = "select a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa <> 0 and " & kategori & " like '%" & txtcari & "%' order by c.tglklaim"
    
    Else
    sql = "select a.kdklaim,c.tglklaim,a.kdcustomer,b.nmcustomer,b.alamat as alamat,a.jmlklaim,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join klaim_hilang c on a.kdklaim=c.kdklaim where a.sisa = 0 and " & kategori & " like '%" & txtcari & "%' order by c.tglklaim"
    
    End If

End If



Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs


For i = 1 To (datagrid1.Rows - 1)
For j = 1 To (datagrid1.Cols - 1)

If datagrid1.TextMatrix(i, 6) = 0 Then
datagrid1.Cell(flexcpForeColor, i, j) = vbRed
datagrid1.Cell(flexcpBackColor, i, j) = vbYellow
datagrid1.Cell(flexcpFontBold, i, j) = True
End If

Next
Next

Call LG
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "a.kdKlaim"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "a.kdcustomer"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "b.nmcustomer"
ElseIf CMBCARI.ListIndex = 3 Then
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
ElseIf KeyAscii = Asc("l") Or KeyAscii = Asc("L") Then
 If rs.RecordCount <> 0 Then
 Call lunas
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call ALL
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
     If rs.RecordCount <> 0 Then
     Call lunas
     End If
ElseIf Index = 2 Then
     If rs.RecordCount <> 0 Then
     Klaim_detail.lblKd = rs!kdklaim
     Klaim_detail.lblkdcustomer = rs!kdcustomer
     Klaim_detail.lblnmcustomer = rs!nmcustomer
     Klaim_detail.lblalamat = rs!alamat
     Klaim_detail.Show vbModal
     
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
 txtcari = ""
 Call ALL
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
ElseIf KeyAscii = Asc("j") Or KeyAscii = Asc("J") Then
 Call all_jml
 
End If
End Sub

Private Sub datagrid1_Click()
TimerG.Interval = 10
End Sub

Private Sub DataGrid1_DblClick()
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
txtcari = ""
 Call ALL
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
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

CMBCARI.AddItem "NO KLAIM"
CMBCARI.AddItem "KD CUSTOMER"
CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "ALAMAT"
CMBCARI.ListIndex = 0




TimerAll.Interval = 10
End Sub



Private Sub Oblunas_Click(Value As Integer)
cmdT(1).Enabled = True
TimerAll.Interval = 10
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
TimerAll.Interval = 10
End Sub

Private Sub Olunas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
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
End If
End Sub



