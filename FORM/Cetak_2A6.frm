VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_2A6 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18750
   LinkTopic       =   "Form2"
   ScaleHeight     =   10965
   ScaleWidth      =   18750
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
      Left            =   4410
      TabIndex        =   2
      Top             =   1260
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
      Left            =   3330
      TabIndex        =   1
      Top             =   1260
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
      Left            =   1575
      TabIndex        =   0
      Top             =   1260
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
      Picture         =   "Cetak_2A6.frx":0000
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
      SectionData     =   "Cetak_2A6.frx":6862
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   10
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
      TabIndex        =   3
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
      Picture         =   "Cetak_2A6.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   6
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
      Picture         =   "Cetak_2A6.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   4
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
      Picture         =   "Cetak_2A6.frx":D33B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17820
      TabIndex        =   5
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
      Picture         =   "Cetak_2A6.frx":10981
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   11
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
      Picture         =   "Cetak_2A6.frx":13E60
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   14
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report AR (Format HO)"
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
      TabIndex        =   13
      Top             =   135
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PER TANGGAL :"
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
      Left            =   360
      TabIndex        =   12
      Top             =   1305
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_2A6.frx":1A6C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_2A6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT1, sqlT2, sqlT3, sqlT4, sql1 As String
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim sqlA As String
Dim color As Long, flag As Byte




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




Private Sub total()

'sql1 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
'       "sum(jmlpiutang - jmlbayar - potongan) as sisa from (" & vbCrLf & _
'       "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar,0 as potongan from piutangsewa" & vbCrLf & _
'       "where tglposting <= '" & Format(txttgl1, "yyyy/MM/dd") & "'" & vbCrLf & _
'       "Union" & vbCrLf & _
'       "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
'       "where tglbayar <= '" & Format(txttgl1, "yyyy/MM/dd") & "'" & vbCrLf & _
'       "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"
'
'sqlA2 = "select '1' as kode,a.kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,a.jmlbayar,a.potongan,a.sisa,c.tglposting,c.bln,c.tahun,convert(float,'" & Format(txttgl1, "yyyy/MM/dd") & "' - c.tglposting) as umur from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
'      "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa <> 0 "
'
'
'If CMBjenis.ListIndex = 0 Then
'sql2 = "select * from (" & sqlA2 & ") a "
'ElseIf CMBjenis.ListIndex = 1 Then
'sql2 = "select * from (" & sqlA2 & ") a where umur <= 30"
'ElseIf CMBjenis.ListIndex = 2 Then
'sql2 = "select * from (" & sqlA2 & ") a where umur > 30 "
'End If
'
'
'sqlT = "select kode, sum(sisa) as sisa from (" & sql2 & ") a  group by kode"
'Set rs = con.Execute(sqlT)
'
'sqlT1 = "select kode, sum(sisa) as sisa from (" & sql2 & ") a  where umur <= 30 group by kode"
'Set rs1 = con.Execute(sqlT1)
'
'sqlT2 = "select kode, sum(sisa) as sisa from (" & sql2 & ") a  where umur >= 31  and umur <= 60 group by kode"
'Set rs2 = con.Execute(sqlT2)
'
'sqlT3 = "select kode, sum(sisa) as sisa from (" & sql2 & ") a  where umur >= 61  and umur <= 90 group by kode"
'Set rs3 = con.Execute(sqlT3)
'
'sqlT4 = "select kode, sum(sisa) as sisa from (" & sql2 & ") a  where umur >= 91 group by kode"
'Set rs4 = con.Execute(sqlT4)
'
'
'


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
Call CetakD
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


Unload AR_2A6

sql = "exec SP_AR_HO @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "'"


With AR_2A6.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_2A6
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcus.DataField = "nmcustomer"
.fldtglposting.DataField = "tglposting"
.fldbln.DataField = "bln"
.fldtahun.DataField = "tahun"
.fldalamat.DataField = "alamat"
.fldpiutang.DataField = "jmlpiutang"
.fldrupiah.DataField = "sisa"
.fldkdpiutang.DataField = "kdpiutang"
.fldumur.DataField = "umur"
.fldunit.DataField = "unit"
.fldharga.DataField = "harga"
.fldtglbyr1.DataField = "tglbayar1"
.fldtglbyr2.DataField = "tglbayar2"
.fldtglbyr3.DataField = "tglbayar3"

.fldbyr1.DataField = "jmlbayar1"
.fldbyr2.DataField = "jmlbayar2"
.fldbyr3.DataField = "jmlbayar3"


.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1

.GroupHeader1.Visible = False


If Chk1.Value = 1 Then
.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = False

.fldkdcustomer.WordWrap = False
.fldnmcus.WordWrap = False
.fldtglposting.WordWrap = False
.fldbln.WordWrap = False
.fldtahun.WordWrap = False
.fldalamat.WordWrap = False
.fldpiutang.WordWrap = False
.fldrupiah.WordWrap = False
.fldrupiah1.WordWrap = False
.fldrupiah2.WordWrap = False
.fldrupiah3.WordWrap = False
.fldrupiah4.WordWrap = False
.fldkdpiutang.WordWrap = False
.fldumur.WordWrap = False
.fldunit.WordWrap = False
.fldharga.WordWrap = False
.fldtglbyr1.WordWrap = False
.fldtglbyr2.WordWrap = False
.fldtglbyr3.WordWrap = False

.fldbyr1.WordWrap = False
.fldbyr2.WordWrap = False
.fldbyr3.WordWrap = False

End If

Set Me.ARV1.ReportSource = AR_2A6
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub


Private Sub CetakD()
On Error GoTo hell


Unload AR_2A6_1
Unload AR_2A6

sql = "exec SP_AR_HO1 @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "'"


With AR_2A6_1.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_2A6_1
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcus.DataField = "nmcustomer"
.fldtglposting.DataField = "tglposting"
.fldbln.DataField = "bln"
.fldtahun.DataField = "tahun"
.fldalamat.DataField = "alamat"
.fldpiutang.DataField = "jmlpiutang"
.fldrupiah.DataField = "sisa"
.fldkdpiutang.DataField = "kdpiutang"
.fldumur.DataField = "umur"
.fldunit.DataField = "unit"
.fldharga.DataField = "harga"
.fldtglbyr1.DataField = "tglbayar1"
.fldtglbyr2.DataField = "tglbayar2"
.fldtglbyr3.DataField = "tglbayar3"

.fldbyr1.DataField = "jmlbayar1"
.fldbyr2.DataField = "jmlbayar2"
.fldbyr3.DataField = "jmlbayar3"

.fldrpPPH23_1.DataField = "rpPPH23_1"
.fldrpPPH23_2.DataField = "rpPPH23_2"
.fldrpPPH23_3.DataField = "rpPPH23_3"

.fldpotongan1.DataField = "potongan1"
.fldpotongan2.DataField = "potongan1"
.fldpotongan3.DataField = "potongan1"



.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1

.GroupHeader1.Visible = False


If Chk1.Value = 1 Then
.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = False

.fldkdcustomer.WordWrap = False
.fldnmcus.WordWrap = False
.fldtglposting.WordWrap = False
.fldbln.WordWrap = False
.fldtahun.WordWrap = False
.fldalamat.WordWrap = False
.fldpiutang.WordWrap = False
.fldrupiah.WordWrap = False
.fldrupiah1.WordWrap = False
.fldrupiah2.WordWrap = False
.fldrupiah3.WordWrap = False
.fldrupiah4.WordWrap = False
.fldkdpiutang.WordWrap = False
.fldumur.WordWrap = False
.fldunit.WordWrap = False
.fldharga.WordWrap = False
.fldtglbyr1.WordWrap = False
.fldtglbyr2.WordWrap = False
.fldtglbyr3.WordWrap = False

.fldbyr1.WordWrap = False
.fldbyr2.WordWrap = False
.fldbyr3.WordWrap = False

.fldrpPPH23_1.WordWrap = False
.fldrpPPH23_2.WordWrap = False
.fldrpPPH23_3.WordWrap = False

.fldpotongan1.WordWrap = False
.fldpotongan2.WordWrap = False
.fldpotongan3.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_2A6_1
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub








Private Sub cmdfs_Click()
If OPT1.Value = True Then
AR_2A6.Show vbModal
Else
AR_2A6_1.Show vbModal
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
Call CetakD
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub SSCommand1_Click()

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





