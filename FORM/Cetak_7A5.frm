VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_7A5 
   BorderStyle     =   0  'None
   ClientHeight    =   10920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10920
   ScaleWidth      =   18795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbbln 
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
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   990
      Width           =   690
   End
   Begin VB.ComboBox cmbtahun 
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
      Left            =   9135
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Width           =   1095
   End
   Begin VB.ComboBox CMB1 
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
      Left            =   4185
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   1500
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
      Left            =   10080
      TabIndex        =   7
      Top             =   1575
      Width           =   555
   End
   Begin VB.Timer Timerxls 
      Left            =   14310
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13860
      Top             =   2295
   End
   Begin VB.Timer TimerPdf 
      Left            =   14805
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
      Left            =   1620
      TabIndex        =   0
      Top             =   990
      Width           =   1590
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00000000&
      Caption         =   "REKAP CUSTOMER IAP"
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
      Left            =   11655
      TabIndex        =   5
      Top             =   1035
      Width           =   2355
   End
   Begin VB.OptionButton OPT1 
      BackColor       =   &H00000000&
      Caption         =   "RINCIAN"
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
      Left            =   10440
      TabIndex        =   4
      Top             =   1035
      Width           =   1185
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16200
      TabIndex        =   8
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
      Picture         =   "Cetak_7A5.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   315
      TabIndex        =   13
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
      Top             =   1080
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
      Picture         =   "Cetak_7A5.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17730
      TabIndex        =   11
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
      Picture         =   "Cetak_7A5.frx":A118
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17730
      TabIndex        =   9
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
      Picture         =   "Cetak_7A5.frx":D2FF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17730
      TabIndex        =   10
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
      Picture         =   "Cetak_7A5.frx":10945
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1485
      TabIndex        =   14
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
      Picture         =   "Cetak_7A5.frx":13E24
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8715
      Left            =   450
      TabIndex        =   12
      Top             =   1485
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   15372
      SectionData     =   "Cetak_7A5.frx":1A686
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "OMSET PERIODE :"
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
      Left            =   6030
      TabIndex        =   20
      Top             =   1080
      Width           =   1545
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
      Left            =   8415
      TabIndex        =   19
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "ANALISA :"
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
      Left            =   3375
      TabIndex        =   18
      Top             =   1035
      Width           =   870
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10440
      TabIndex        =   17
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Analisa Dispencer dan Showcase"
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
      Left            =   1170
      TabIndex        =   16
      Top             =   135
      Width           =   7665
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
      Left            =   405
      TabIndex        =   15
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_7A5.frx":1A6C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_7A5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte
Dim kategori As String


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

Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
If CMB1.ListIndex = 0 Then
Call Cetak_DS
Else
Call Cetak_SH
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
    
    If CMB1.ListIndex = 0 Then
    Call Cetak_DS
    Else
    Call Cetak_SH
    End If
        
ElseIf KeyAscii = 27 Then
Unload Me
End If

End Sub


Private Sub Cetak_DS()
On Error GoTo hell

Unload AR_ANALISASH
Unload AR_ANALISADS

If Opt1.Value = True Then
sql = "exec SP_Analisa_DS1 @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "',@bln=" & cmbbln.Text & ",@thn=" & CMbtahun.Text & " "
Else
sql = "exec SP_Analisa_DS2 @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "',@bln=" & cmbbln.Text & ",@thn=" & CMbtahun.Text & " "
End If


With AR_ANALISADS.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_ANALISADS

If Opt1.Value = True Then
.fldkdcust.DataField = "kdcustomer"
.fldnmcust.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldarea.DataField = "nmareaC"
.fldPIC.DataField = "nmpic"
.flddl.DataField = "keterangan"
.fldcp.DataField = "cp"
.fldtelp.DataField = "telp"

Else
.fldkdcust = ""
.fldnmcust = ""
.fldalamat = ""
.fldarea = ""
.fldPIC = ""
.flddl.DataField = "jmlcust"
.fldcp = ""
.fldtelp = ""

End If


.fldkdcust_iap.DataField = "kdcustomer_iap"
.fldnmcust_iap.DataField = "nmcustomer_iap"
.fldalamat_iap.DataField = "alamat_iap"
.fldbp.DataField = "D1"
.fldtsp.DataField = "D2"
.fldhn.DataField = "D3"
.fldhc.DataField = "D4"
.fldrakgln.DataField = "D5"
.fldtotal.DataField = "total"
.fldnmsp.DataField = "nmSP"
.fldomsetgln.DataField = "qty_gln"



.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")
.lbltgl1 = txttgl1
.lbljudul = "ANALISA PEMAKAIAN DISPENCER"
'
'
Set rs = con.Execute("exec SP_Analisa_DS_total @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "'")

If rs.RecordCount <> 0 Then
.LBLBP = Format(rs!D1, "#,###0")
.lbltsp = Format(rs!D2, "#,###0")
.lblhn = Format(rs!D3, "#,###0")
.lblhc = Format(rs!D4, "#,###0")
.lblrakgln = Format(rs!D5, "#,###0")
.lbltotal = Format(rs!total, "#,###0")
Else
.LBLBP = 0
.lbltsp = 0
.lblhn = 0
.lblhc = 0
.lblrakgln = 0
.lbltotal = 0

End If
''
.GroupHeader1.Visible = False
'
'
If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = True

.fldkdcust_iap.WordWrap = False
.fldnmcust_iap.WordWrap = False
.fldalamat_iap.WordWrap = False
.fldbp.WordWrap = False
.fldtsp.WordWrap = False
.fldhn.WordWrap = False
.fldhc.WordWrap = False
.fldrakgln.WordWrap = False
.fldtotal.WordWrap = False
.fldkdcust.WordWrap = False
.fldnmcust.WordWrap = False
.fldalamat.WordWrap = False
.fldarea.WordWrap = False
.fldPIC.WordWrap = False
.flddl.WordWrap = False
.fldno.WordWrap = False
.fldnmsp.WordWrap = False
.fldomsetgln.WordWrap = False
.fldcp.WordWrap = False
.fldtelp.WordWrap = False

End If

Set Me.ARV1.ReportSource = AR_ANALISADS
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub


Private Sub Cetak_SH()
On Error GoTo hell

Unload AR_ANALISASH
Unload AR_ANALISADS

If Opt1.Value = True Then
sql = "exec SP_Analisa_SH1 @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "',@bln=" & cmbbln.Text & ",@thn=" & CMbtahun.Text & " "
Else
sql = "exec SP_Analisa_SH2 @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "',@bln=" & cmbbln.Text & ",@thn=" & CMbtahun.Text & " "
End If


With AR_ANALISASH.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_ANALISASH

If Opt1.Value = True Then
.fldkdcust.DataField = "kdcustomer"
.fldnmcust.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldarea.DataField = "nmareaC"
.fldPIC.DataField = "nmpic"
.flddl.DataField = "keterangan"
.fldcp.DataField = "cp"
.fldtelp.DataField = "telp"
Else
.fldkdcust = ""
.fldnmcust = ""
.fldalamat = ""
.fldarea = ""
.fldPIC = ""
.flddl.DataField = "jmlcust"
.fldcp = ""
.fldtelp = ""
End If


.fldkdcust_iap.DataField = "kdcustomer_iap"
.fldnmcust_iap.DataField = "nmcustomer_iap"
.fldalamat_iap.DataField = "alamat_iap"
.fldkecil.DataField = "S1"
.fldbesar.DataField = "S2"
.fldtotal.DataField = "total"
.fldnmsp.DataField = "nmSP"
.fldomsetSPS.DataField = "qty_sps"



.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbljudul = "ANALISA PEMAKAIAN SHOWCASE"
'
'
Set rs = con.Execute("exec SP_Analisa_SH_total @tgl1='" & Format(txttgl1, "yyyy/MM/dd") & "'")

If rs.RecordCount <> 0 Then
.lblkecil = Format(rs!S1, "#,###0")
.lblbesar = Format(rs!S2, "#,###0")
.lbltotal = Format(rs!total, "#,###0")
Else
.lblkecil = 0
.lblbesar = 0
.lbltotal = 0

End If
'
.GroupHeader1.Visible = False
'
'
If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = True

.fldkdcust_iap.WordWrap = False
.fldnmcust_iap.WordWrap = False
.fldalamat_iap.WordWrap = False
.fldbesar.WordWrap = False
.fldkecil.WordWrap = False
.fldtotal.WordWrap = False
.fldkdcust.WordWrap = False
.fldnmcust.WordWrap = False
.fldalamat.WordWrap = False
.fldarea.WordWrap = False
.fldPIC.WordWrap = False
.flddl.WordWrap = False
.fldno.WordWrap = False
.fldnmsp.WordWrap = False
.fldomsetSPS.WordWrap = False
.fldcp.WordWrap = False
.fldtelp.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_ANALISASH
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub









Private Sub cmdfs_Click()
If Opt1.Value = True Then
AR_ANALISADS.Zoom = 110
AR_ANALISADS.Show vbModal

Else
AR_ANALISADS.Zoom = 110
AR_ANALISADS.Show vbModal
End If

End Sub

Private Sub cmdfs_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_Click()
If CMB1.ListIndex = 0 Then
Call Cetak_DS
Else
Call Cetak_SH
End If

ARV1.ToolbarVisible = False
ARV1.ToolbarVisible = True
End Sub

Private Sub cmdGO_Click()
If CMB1.ListIndex = 0 Then
Call Cetak_DS
Else
Call Cetak_SH
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



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0

Opt1.Value = True

CMB1.AddItem "DISPENCER"
CMB1.AddItem "SHOWCASE"
CMB1.ListIndex = 0

CMbtahun.AddItem Year(Date) - 3
CMbtahun.AddItem Year(Date) - 2
CMbtahun.AddItem Year(Date) - 1
CMbtahun.AddItem Year(Date)
CMbtahun.AddItem Year(Date) + 1
CMbtahun.AddItem Year(Date) + 2
CMbtahun.AddItem Year(Date) + 3
CMbtahun.ListIndex = 3

cmbbln.AddItem "1"
cmbbln.AddItem "2"
cmbbln.AddItem "3"
cmbbln.AddItem "4"
cmbbln.AddItem "5"
cmbbln.AddItem "6"
cmbbln.AddItem "7"
cmbbln.AddItem "8"
cmbbln.AddItem "9"
cmbbln.AddItem "10"
cmbbln.AddItem "11"
cmbbln.AddItem "12"

If Month(Date) > 1 Then
cmbbln.ListIndex = CLng(Month(Date)) - 2
Else
cmbbln.ListIndex = 0
End If

txttgl1 = Date
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









