VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Real_cek_List 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   18795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCetak 
      Left            =   13275
      Top             =   2340
   End
   Begin VB.Timer TimerPdf 
      Left            =   14895
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13950
      Top             =   2295
   End
   Begin VB.Timer Timerxls 
      Left            =   14400
      Top             =   2295
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
      Left            =   10395
      TabIndex        =   4
      Top             =   1575
      Width           =   555
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
      Left            =   2205
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
      Left            =   4140
      TabIndex        =   3
      Top             =   990
      Width           =   1590
   End
   Begin VB.CheckBox chktgl 
      BackColor       =   &H00000000&
      Caption         =   "TANGGAL CEK :"
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
      Left            =   405
      TabIndex        =   1
      Top             =   1035
      Width           =   1770
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16065
      TabIndex        =   5
      Top             =   1575
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
      Picture         =   "Real_cek_List.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
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
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   9
      ToolTipText     =   "Simpan"
      Top             =   4815
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
      Picture         =   "Real_cek_List.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   7
      ToolTipText     =   "Simpan"
      Top             =   3195
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
      Picture         =   "Real_cek_List.frx":9A49
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17820
      TabIndex        =   8
      ToolTipText     =   "Simpan"
      Top             =   4005
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
      Picture         =   "Real_cek_List.frx":D08F
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
      Picture         =   "Real_cek_List.frx":1056E
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8715
      Left            =   360
      TabIndex        =   0
      Top             =   1485
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   15372
      SectionData     =   "Real_cek_List.frx":16DD0
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   17730
      TabIndex        =   6
      ToolTipText     =   "Simpan"
      Top             =   1125
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
      Picture         =   "Real_cek_List.frx":16E0C
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Real Cek List"
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
      TabIndex        =   14
      Top             =   135
      Width           =   4560
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   13
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
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
      Left            =   3780
      TabIndex        =   12
      Top             =   1035
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Real_cek_List.frx":1A6C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Real_cek_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte
Dim kata, kategori As String
Dim rs1 As ADODB.Recordset


Private Sub chktgl_Click()
If chktgl.Value = 0 Then
    txttgl1.Enabled = False
    txttgl2.Enabled = False
Else
    txttgl1.Enabled = True
    txttgl2.Enabled = True
    txttgl1 = Date
    txttgl2 = Date
End If
End Sub

Private Sub chktgl_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 27 Then
Unload Me
End If

End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdGO_Click()
Call Cetak
End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub

Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
Call Cetak
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

MousePointer = vbHourglass

Unload AR_Real_Cek_list

    sqlX = "select kdcustomer,kdteknisi,nmrute,tglplan from route_plan where kdteknisi='" & Real_cek_TU.lblkdteknisi & "' and  nmrute= '" & Real_cek_TU.txtperiode & "' "

    sql1 = "select a.idcek,e.tglplan,a.tglcek,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,a.kdbarang,d.kd1,d.nmbarang,a.unit,a.keterangan from Real_cek a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join barang d on a.kdbarang=d.kdbarang left join (" & sqlX & ") e on a.kdcustomer=e.kdcustomer  where a.kdteknisi='" & Real_cek_TU.lblkdteknisi & "' and  a.nmrute= '" & Real_cek_TU.txtperiode & "' "
    

        
    If chktgl.Value = 0 Then
    sql = "select * from (" & sql1 & ") a  order by a.tglcek ,a.nmcustomer ,a.alamat"
    
    Else
    sql = "select * from (" & sql1 & ") a where a.tglcek between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "'  order by a.tglcek ,a.nmcustomer ,a.alamat"
    
    End If


Text1 = sqlX

With AR_Real_Cek_list.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_Real_Cek_list
.fldnmcustomer.DataField = "nmcustomer"
.fldkdcustomer.DataField = "kdcustomer"
.fldalamat.DataField = "alamat"
.fldtglplan.DataField = "tglplan"
.fldtglCek.DataField = "tglcek"
.fldnmareaC.DataField = "nmareaC"
.fldkdbarang.DataField = "kdbarang"
.fldkd1.DataField = "kd1"
.fldnmbarang.DataField = "nmbarang"
.fldketerangan.DataField = "keterangan"
.fldunit.DataField = "unit"

.lblnmteknisi = fixrute_TU.lblnmteknisi
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2

.lblcetak = Format(Date, "dd/MM/yyyy")
'

.GroupHeader1.Visible = False

If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True

.fldnmcustomer.WordWrap = False
.fldkdcustomer.WordWrap = False
.fldalamat.WordWrap = False
.fldtglplan.WordWrap = False
.fldtglCek.WordWrap = False
.fldnmareaC.WordWrap = False
.fldkdbarang.WordWrap = False
.fldkd1.WordWrap = False
.fldnmbarang.WordWrap = False
.fldketerangan.WordWrap = False
.fldunit.WordWrap = False
.fldNo.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_Real_Cek_list
End With

MousePointer = vbDefault

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub







Private Sub cmdfs_Click()
AR_Real_Cek_list.Show vbModal
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



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttgl1 = Date
txttgl2 = Date

txttgl1.Enabled = False
txttgl2.Enabled = False

'TimerCetak.Interval = 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
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








