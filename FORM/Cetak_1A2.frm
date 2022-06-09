VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_1A2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18825
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   18825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00000000&
      Caption         =   "DETAIL"
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
      Left            =   1845
      TabIndex        =   5
      Top             =   1485
      Width           =   1050
   End
   Begin VB.OptionButton OPT1 
      BackColor       =   &H00000000&
      Caption         =   "SIMPLE "
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
      Left            =   630
      TabIndex        =   4
      Top             =   1485
      Width           =   1050
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Isi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9810
      TabIndex        =   7
      Top             =   2115
      Width           =   555
   End
   Begin VB.Timer Timerxls 
      Left            =   14400
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13950
      Top             =   2295
   End
   Begin VB.Timer TimerPdf 
      Left            =   14895
      Top             =   2295
   End
   Begin VB.TextBox txttgl1 
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
      Left            =   13995
      TabIndex        =   2
      Top             =   990
      Width           =   1590
   End
   Begin VB.TextBox txttgl2 
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
      Left            =   15930
      TabIndex        =   3
      Top             =   990
      Width           =   1590
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16065
      TabIndex        =   8
      Top             =   2070
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   262144
      ForeColor       =   255
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_1A2.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8220
      Left            =   360
      TabIndex        =   9
      Top             =   1980
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14499
      SectionData     =   "Cetak_1A2.frx":6862
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   6525
      TabIndex        =   0
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
      Picture         =   "Cetak_1A2.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   14
      Top             =   810
      Width           =   17205
      _Version        =   524288
      _ExtentX        =   30348
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   17730
      TabIndex        =   6
      ToolTipText     =   "Simpan"
      Top             =   1170
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_1A2.frx":90D0
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   12
      ToolTipText     =   "Simpan"
      Top             =   4590
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_1A2.frx":C986
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   10
      ToolTipText     =   "Simpan"
      Top             =   2970
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_1A2.frx":FB6D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17820
      TabIndex        =   11
      ToolTipText     =   "Simpan"
      Top             =   3780
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_1A2.frx":131B3
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   13
      Top             =   10440
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
      Picture         =   "Cetak_1A2.frx":16692
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   12915
      TabIndex        =   1
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
      Picture         =   "Cetak_1A2.frx":1CEF4
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label lblnmbarang 
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
      Left            =   9585
      TabIndex        =   24
      Top             =   990
      Width           =   3345
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
      Left            =   8055
      TabIndex        =   23
      Top             =   990
      Width           =   1500
   End
   Begin VB.Label Label5 
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
      Left            =   7200
      TabIndex        =   22
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   21
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
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
      Left            =   2475
      TabIndex        =   20
      Top             =   990
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
      Left            =   450
      TabIndex        =   19
      Top             =   1035
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
      Left            =   1305
      TabIndex        =   18
      Top             =   990
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kartu Stok"
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
      Top             =   135
      Width           =   6585
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL :"
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
      Left            =   13545
      TabIndex        =   16
      Top             =   1035
      Width           =   420
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SD"
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
      Left            =   15570
      TabIndex        =   15
      Top             =   1035
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_1A2.frx":1F726
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_1A2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte

Private Sub cmdBR1_Click()
Barang_BR.LBLKODE = UCase("1A2")
Barang_BR.Show vbModal

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
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub




Private Sub total()

sql1 = "select '1' AS KODE,a.tglRpinjam,a.kdRpinjam,a.kdcustomer,c.nmcustomer,c.alamat,b.kdbarang,d.nmbarang,b.unit,b.keterangan, a.kdpinjam,e.tglpinjam" & vbCrLf & _
      "from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam left join customer c on a.kdcustomer=c.kdcustomer left join barang d on b.kdbarang=d.kdbarang" & vbCrLf & _
      "left join pinjam e on a.kdpinjam=e.kdpinjam where a.kdgudang ='" & lblkdgudang & "' and a.tglRpinjam between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' "


sqlT = "select kode, sum(unit) as unit from (" & sql1 & ") a group by kode"
Set rs = con.Execute(sqlT)

End Sub




 





Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
If OPT1.Value = True Then
Call Cetak
Else
Call Cetak1
End If
End Sub

Private Sub CHK1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    If Chk1.Value = 1 Then
    Chk1.Value = 0
    Else
    Chk1.Value = 1
    End If
    
    Call Cetak
        
ElseIf KeyAscii = 27 Then
Unload Me
End If

End Sub


Private Sub Cetak()
On Error GoTo hell

Unload AR_1A2_01
Unload AR_1A2

sql1 = "select '0' As srt,'' as judul,'" & Format(txttgl1, "yyyy/MM/dd") & "' as tglbukti,'SALDO AWAL' as nobukti,'' as nosj,'' as nmcustomer,a.kdgudang,a.kdbarang,a.nmbarang,'' as keterangan,sum(a.masuk) as masuk,sum(a.keluar) as keluar,0 as H_masuk,0 as H_keluar,sum(rp_masuk) as rp_masuk,sum(rp_keluar) as rp_keluar,'' as x1  from rinci_stok a where a.kdbarang='" & lblkdbarang & "' and a.kdgudang='" & lblkdgudang & "' " & vbCrLf & _
       "and a.tglbukti < '" & Format(txttgl1, "yyyy/MM/dd") & "' group by a.kdgudang,a.kdbarang,a.nmbarang"
       
sql2 = "select a.*,(a.judul + ' - ' + a.keterangan) as x1  from rinci_stok a where a.kdbarang='" & lblkdbarang & "' and a.kdgudang='" & lblkdgudang & "' " & vbCrLf & _
       "and a.tglbukti between '" & Format(txttgl1, "yyyy/MM/dd") & "'  and '" & Format(txttgl2, "yyyy/MM/dd") & "'"


sql = "select * from (" & sql1 & " union " & sql2 & ") a order by tglbukti,srt"

With AR_1A2.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_1A2
.fldtglbukti.DataField = "tglbukti"
.fldkdbukti.DataField = "nobukti"
.fldnosj.DataField = "nosj"
.fldnmcustomer.DataField = "nmcustomer"
.fldketerangan.DataField = "X1"
.fldmasuk.DataField = "masuk"
.fldkeluar.DataField = "keluar"
.fldsrt.DataField = "srt"

.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2
.lblnmgudang = lblnmgudang
.lblnmbarang = lblnmbarang


If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False

.fldtglbukti.WordWrap = False
.fldkdbukti.WordWrap = False
.fldnosj.WordWrap = False
.fldnmcustomer.WordWrap = False
.fldketerangan.WordWrap = False
.fldmasuk.WordWrap = False
.fldkeluar.WordWrap = False

End If
'
Set Me.ARV1.ReportSource = AR_1A2
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub


Private Sub Cetak1()
On Error GoTo hell

Unload AR_1A2_01
Unload AR_1A2

sql1 = "select '0' As srt,'' as judul,'" & Format(txttgl1, "yyyy/MM/dd") & "' as tglbukti,'SALDO AWAL' as nobukti,'' as nosj,'' as nmcustomer,a.kdgudang,a.kdbarang,a.nmbarang,'' as keterangan,sum(a.masuk) as masuk,sum(a.keluar) as keluar,0 as H_masuk,0 as H_keluar,sum(rp_masuk) as rp_masuk,sum(rp_keluar) as rp_keluar,'' as x1  from rinci_stok a where a.kdbarang='" & lblkdbarang & "' and a.kdgudang='" & lblkdgudang & "' " & vbCrLf & _
       "and a.tglbukti < '" & Format(txttgl1, "yyyy/MM/dd") & "' group by a.kdgudang,a.kdbarang,a.nmbarang"
       
sql2 = "select a.*,(a.judul + ' - ' + a.keterangan) as x1  from rinci_stok a where a.kdbarang='" & lblkdbarang & "' and a.kdgudang='" & lblkdgudang & "' " & vbCrLf & _
       "and a.tglbukti between '" & Format(txttgl1, "yyyy/MM/dd") & "'  and '" & Format(txttgl2, "yyyy/MM/dd") & "'"


sql = "select * from (" & sql1 & " union " & sql2 & ") a order by tglbukti,srt"

With AR_1A2_01.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_1A2_01
.fldtglbukti.DataField = "tglbukti"
.fldkdbukti.DataField = "nobukti"
.fldnosj.DataField = "nosj"
.fldnmcustomer.DataField = "nmcustomer"
.fldketerangan.DataField = "X1"
.fldmasuk.DataField = "masuk"
.fldkeluar.DataField = "keluar"
.fldsrt.DataField = "srt"
.fldH_masuk.DataField = "H_masuk"
.fldrp_masuk.DataField = "rp_masuk"
.fldH_keluar.DataField = "H_keluar"
.fldrp_keluar.DataField = "rp_keluar"


.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2
.lblnmgudang = lblnmgudang
.lblnmbarang = lblnmbarang


If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False

.fldtglbukti.WordWrap = False
.fldkdbukti.WordWrap = False
.fldnosj.WordWrap = False
.fldnmcustomer.WordWrap = False
.fldketerangan.WordWrap = False
.fldmasuk.WordWrap = False
.fldkeluar.WordWrap = False
.fldH_masuk.WordWrap = False
.fldrp_masuk.WordWrap = False
.fldH_keluar.WordWrap = False
.fldrp_keluar.WordWrap = False

End If
'
Set Me.ARV1.ReportSource = AR_1A2_01
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub




Private Sub cmdBRKr_Click()
Karyawan_BR.LBLKODE = "LAD"
Karyawan_BR.Show vbModal

End Sub

Private Sub cmdBRKr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub





Private Sub cmdBR_Click()
Gudang_BR.LBLKODE = "1A2"
Gudang_BR.Show vbModal
End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub cmdfs_Click()
If OPT1.Value = True Then
AR_1A2.Show vbModal
Else
AR_1A2_01.Zoom = 110
AR_1A2_01.Show vbModal
End If
End Sub

Private Sub cmdfs_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_Click()
Call Cetak
ARV1.ToolbarVisible = False
ARV1.ToolbarVisible = True
End Sub

Private Sub cmdGO_Click()
If OPT1.Value = True Then
Call Cetak
Else
Call Cetak1
End If

End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdPDF_Click()
TimerPdf.Interval = 10
End Sub

Private Sub cmdPDF_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdrtf_Click()
TimerRtf.Interval = 10
End Sub

Private Sub cmdrtf_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub cmdxls_Click()
Timerxls.Interval = 10
End Sub


Private Sub cmdxls_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub Form_Load()
GradientForm Me, 0

OPT1.Value = True

txttgl1 = Date
txttgl2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub OPT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub Opt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If

End Sub

Private Sub TimerPDF_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF

out2 = out2 + 1

Call save_out
pdf.filename = alamat_save & "\outfile" & CStr(out2) & ".pdf"
pdf.Export ARV1.Pages

Call EX_PDF(Me)
TimerPdf.Interval = 0

Exit Sub
hell:
TimerPdf.Interval = 0
If out2 < 10 Then
cmdPDF_Click
End If

End Sub

Private Sub Timerrtf_Timer()
On Error GoTo hell
Dim rtf As New ActiveReportsRTFExport.ARExportRTF
out = out + 1

Call save_out
rtf.filename = alamat_save & "\outfile" & CStr(out) & ".rtf"
rtf.Export ARV1.Pages

Call EX_WORD(Me)
TimerRtf.Interval = 0

Exit Sub
hell:
TimerRtf.Interval = 0
If out < 10 Then
cmdrtf_Click
End If
End Sub

Private Sub Timerxls_Timer()
On Error GoTo hell
Dim xls As New ActiveReportsExcelExport.ARExportExcel



out1 = out1 + 1

Call save_out
xls.filename = alamat_save & "\outfile" & CStr(out1) & ".xls"
xls.Export ARV1.Pages

Call EX_EXEL(Me)
Timerxls.Interval = 0

Exit Sub
hell:
Timerxls.Interval = 0
If out1 < 10 Then
cmdxls_Click
End If
End Sub









Private Sub txttgl1_Change()
Call nul(txttgl1)
End Sub

Private Sub txttgl1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttgl1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttgl1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890-/", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If

End If

End Sub

Private Sub txttgl1_LostFocus()
On Error GoTo hell

txttgl1 = FormatDateTime(txttgl1, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttgl1.SetFocus
End Sub

Private Sub txttgl2_Change()
Call nul(txttgl2)
End Sub

Private Sub txttgl2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttgl2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttgl2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890-/", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If

End If

End Sub

Private Sub txttgl2_LostFocus()
On Error GoTo hell

txttgl2 = FormatDateTime(txttgl2, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttgl2.SetFocus
End Sub











