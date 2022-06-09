VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Kwitansi_LIST 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18735
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   18735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCetak 
      Left            =   11205
      Top             =   540
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
      Left            =   9855
      TabIndex        =   1
      Top             =   1080
      Width           =   555
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16065
      TabIndex        =   2
      Top             =   1080
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
      Picture         =   "Kwitansi_LIST.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   9210
      Left            =   360
      TabIndex        =   0
      Top             =   990
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   16245
      SectionData     =   "Kwitansi_LIST.frx":6862
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   9
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
      Left            =   17775
      TabIndex        =   5
      ToolTipText     =   "Simpan"
      Top             =   3060
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
      Picture         =   "Kwitansi_LIST.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17775
      TabIndex        =   3
      ToolTipText     =   "Simpan"
      Top             =   1440
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
      Picture         =   "Kwitansi_LIST.frx":9A85
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17775
      TabIndex        =   4
      ToolTipText     =   "Simpan"
      Top             =   2250
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
      Picture         =   "Kwitansi_LIST.frx":D0CB
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   8
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
      Picture         =   "Kwitansi_LIST.frx":105AA
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSOption Opt1 
      Height          =   330
      Left            =   12555
      TabIndex        =   6
      Top             =   360
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
      Caption         =   "List Kwitansi"
   End
   Begin Threed.SSOption Opt2 
      Height          =   330
      Left            =   13995
      TabIndex        =   7
      Top             =   360
      Width           =   3570
      _ExtentX        =   6297
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
      Caption         =   "List Kwitansi by NPWP (Customer PKP)"
   End
   Begin VB.Label lblFRM 
      Caption         =   "lblFRM"
      Height          =   285
      Left            =   9585
      TabIndex        =   12
      Top             =   405
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Kwitansi"
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
      TabIndex        =   11
      Top             =   135
      Width           =   4560
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   10
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Kwitansi_LIST.frx":16E0C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Kwitansi_LIST"
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

Private Sub Bk()
Select Case Kwitansi.cmbbln.ListIndex
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

If Kwitansi.CMBCARI.ListIndex = 0 Then
kategori = "a.kdpiutang"
ElseIf Kwitansi.CMBCARI.ListIndex = 1 Then
kategori = "b.nmcustomer"
ElseIf Kwitansi.CMBCARI.ListIndex = 2 Then
kategori = "b.alamat"
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

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub total()
If lblfrm = "KWITANSI" Then
    Call Bk
    
    If Kwitansi.TXTCARI = "" Then
    sql1 = "select '1' as kode,a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting from piutangsewa a " & vbCrLf & _
          "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & Kwitansi.txttahun & " "
    Else
    sql1 = "select '1' as kode,a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting from piutangsewa a " & vbCrLf & _
          "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & Kwitansi.txttahun & " and " & kategori & " like '%" & Kwitansi.TXTCARI & "%' "
    End If
ElseIf lblfrm = "POSTING" Then
    sqlA1 = "select kdcustomer,sum(unit) as unit from (" & vbCrLf & _
           "select 'A' as kode,a.kdsewa,b.kdcustomer,a.kdbarang,a.unit from sewa_d a left join  sewa b on a.kdsewa=b.kdsewa where b.tglsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "'" & vbCrLf & _
           "Union" & vbCrLf & _
           "select 'B' as kode,b.kdsewa,b.kdcustomer,a.kdbarang,-sum(a.unit) as unit from Rsewa_d a left join Rsewa b on a.kdRsewa =b.kdRsewa" & vbCrLf & _
           "where b.tglRsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' group by b.kdsewa,b.kdcustomer,A.kdbarang" & vbCrLf & _
           " ) a group by kdcustomer"
           
     sql1 = "select '1' as kode,a.kdcustomer +'/' + '" & Posting.lblbln & "' + '/' + '" & CStr(Posting.txttahun) & "' as kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat," & CCur(Posting.cmbbln.Text) & " as bln," & CCur(Posting.txttahun) & " as tahun,a.unit,b.hrgsewa as harga,(a.unit * b.hrgsewa) as jmlpiutang,'" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' as tglposting from  " & vbCrLf & _
          "(" & sqlA1 & ") a left join customer b on a.kdcustomer=b.kdcustomer where a.unit<>0 "


End If


sqlT = "select kode,sum(unit) as unit, sum(jmlpiutang) as jmlpiutang from (" & sql1 & ") a group by kode"
Set rs = con.Execute(sqlT)

End Sub




 





Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
If Opt1.Value = True Then
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
On Error Resume Next


Unload AR_Kwitansi_list
Unload AR_Kwitansi_NPWP


If lblfrm = "KWITANSI" Then
    Call Bk
    
    If Kwitansi.TXTCARI = "" Then
    sql = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting from piutangsewa a " & vbCrLf & _
          "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & Kwitansi.txttahun & " order by a.kdpiutang,a.kdcustomer"
    Else
    sql = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting from piutangsewa a " & vbCrLf & _
          "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & Kwitansi.txttahun & " and " & kategori & " like '%" & Kwitansi.TXTCARI & "%' order by a.kdpiutang,a.kdcustomer"
    End If

ElseIf lblfrm = "POSTING" Then

    sqlA = "select kdcustomer,sum(unit) as unit from (" & vbCrLf & _
           "select 'A' as kode,a.kdsewa,b.kdcustomer,a.kdbarang,a.unit from sewa_d a left join  sewa b on a.kdsewa=b.kdsewa where b.tglsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "'" & vbCrLf & _
           "Union" & vbCrLf & _
           "select 'B' as kode,a.kdsewa,b.kdcustomer,a.kdbarang,-sum(a.unit) as unit from Rsewa_d a left join Rsewa b on a.kdRsewa =b.kdRsewa" & vbCrLf & _
           "where b.tglRsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' group by a.kdsewa,b.kdcustomer,a.kdbarang" & vbCrLf & _
           " ) a group by kdcustomer"
           
    sql = "select a.kdcustomer +'/' + '" & Posting.lblbln & "' + '/' + '" & CStr(Posting.txttahun) & "' as kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat," & CCur(Posting.cmbbln.Text) & " as bln," & CCur(Posting.txttahun) & " as tahun,a.unit,b.hrgsewa as harga,(a.unit * b.hrgsewa) as jmlpiutang,'" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' as tglposting from  " & vbCrLf & _
          "(" & sqlA & ") a left join customer b on a.kdcustomer=b.kdcustomer where a.unit<>0 order by a.kdcustomer"

End If

With AR_Kwitansi_list.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_Kwitansi_list
.fldkdpiutang.DataField = "kdpiutang"
.fldtglposting.DataField = "tglposting"
.fldnmcus.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldbln.DataField = "bln"
.fldtahun.DataField = "tahun"
.fldunit.DataField = "Unit"
.fldharga.DataField = "harga"
.fldjmlpiutang.DataField = "jmlpiutang"

If lblfrm = "POSTING" Then
.lbljudul = "LIST KWITANSI SEBELUM POSTING"
Else
.lbljudul = "LIST KWITANSI"
End If
.lblcetak = Format(Date, "dd/MM/yyyy")

Call total
If rs.RecordCount <> 0 Then
.lblunit = Format(rs!unit, "#,###0")
.lbljmlpiutang = Format(rs!jmlpiutang, "#,###0")
Else
.lblunit = 0
.lbljmlpiutang = 0
End If


If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupFooter1.Visible = False
.GroupHeader1.Visible = False

.fldkdpiutang.WordWrap = False
.fldtglposting.WordWrap = False
.fldnmcus.WordWrap = False
.fldalamat.WordWrap = False
.fldbln.WordWrap = False
.fldtahun.WordWrap = False
.fldunit.WordWrap = False
.fldharga.WordWrap = False
.fldjmlpiutang.WordWrap = False
.fldno.WordWrap = False
End If

Set Me.ARV1.ReportSource = AR_Kwitansi_list
End With


End Sub


Private Sub Cetak1()
On Error Resume Next


Unload AR_Kwitansi_list
Unload AR_Kwitansi_NPWP

If lblfrm = "KWITANSI" Then
    Call Bk
       
    sql = "select a.kdpiutang,b.npwp,b.nmNPWP,b.alamatNPWP,a.unit,(a.jmlpiutang/a.unit) as harga,(a.jmlpiutang/a.unit) / 1.11 as harga1 from piutangsewa a " & vbCrLf & _
          "left join customer b on a.kdcustomer=b.kdcustomer where " & kata & " and a.tahun=" & Kwitansi.txttahun & " order by b.npwp"
          
        
'ElseIf lblFRM = "POSTING" Then
'
'    sqlA = "select kdcustomer,sum(unit) as unit from (" & vbCrLf & _
'           "select 'A' as kode,a.kdsewa,b.kdcustomer,a.kdbarang,a.unit from sewa_d a left join  sewa b on a.kdsewa=b.kdsewa where b.tglsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "'" & vbCrLf & _
'           "Union" & vbCrLf & _
'           "select 'B' as kode,b.kdsewa,b.kdcustomer,a.kdbarang,-sum(a.unit) as unit from Rsewa_d a left join Rsewa b on a.kdRsewa =b.kdRsewa" & vbCrLf & _
'           "where b.tglRsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' group by b.kdsewa,b.kdcustomer,a.kdbarang" & vbCrLf & _
'           " ) a group by kdcustomer"
'
'    sql = "select a.kdcustomer +'/' + '" & Posting.lblbln & "' + '/' + '" & CStr(Posting.txttahun) & "' as kdpiutang,a.kdcustomer,b.nmcustomer,b.alamat," & CCur(Posting.CMBBLN.Text) & " as bln," & CCur(Posting.txttahun) & " as tahun,a.unit,b.hrgsewa as harga,(a.unit * b.hrgsewa) as jmlpiutang,'" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' as tglposting from  " & vbCrLf & _
'          "(" & sqlA & ") a left join customer b on a.kdcustomer=b.kdcustomer where a.unit<>0 order by a.kdcustomer"

End If

With AR_Kwitansi_NPWP.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_Kwitansi_NPWP
.fldnmcus.DataField = "nmNPWP"
.fldalamat.DataField = "alamatNPWP"
.fldunit.DataField = "Unit"
.fldharga.DataField = "harga"
.fldharga1.DataField = "harga1"
.fldkdpiutang.DataField = "kdpiutang"
.fldnoNPWP.DataField = "npwp"

.lblcetak = Format(Now, "dd/MM/yyyy")


.GroupHeader1.Visible = False

If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True

.fldnmcus.WordWrap = False
.fldalamat.WordWrap = False
.fldunit.WordWrap = False
.fldharga.WordWrap = False
.fldharga1.WordWrap = False
.fldkdpiutang.WordWrap = False
.fldnoNPWP.WordWrap = False

End If

Set Me.ARV1.ReportSource = AR_Kwitansi_NPWP
End With


End Sub





Private Sub cmdBRKr_Click()
Karyawan_BR.LBLKODE = "LAD"
Karyawan_BR.Show vbModal

End Sub

Private Sub cmdBRKr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub




Private Sub cmdfs_Click()
If Opt1.Value = True Then
AR_Kwitansi_list.Show vbModal
Else
AR_Kwitansi_NPWP.Show vbModal
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

Opt1.Value = True
TimerCetak.Interval = 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub OPT1_Click(Value As Integer)
Call Cetak
End Sub

Private Sub OPT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Opt2_Click(Value As Integer)
Call Cetak1
End Sub

Private Sub Opt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub TimerCetak_Timer()
Call Cetak
TimerCetak.Interval = 0
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







