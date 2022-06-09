VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_9A2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   18810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OPT1 
      BackColor       =   &H00000000&
      Caption         =   "REKAP"
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
      Left            =   495
      TabIndex        =   3
      Top             =   1440
      Width           =   1050
   End
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
      Left            =   1575
      TabIndex        =   4
      Top             =   1440
      Width           =   1050
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
      Left            =   3465
      TabIndex        =   1
      Top             =   1035
      Width           =   1590
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1035
      Width           =   1590
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
      TabIndex        =   6
      Top             =   1890
      Width           =   555
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16155
      TabIndex        =   7
      Top             =   1890
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
      Picture         =   "Cetak_9A2.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8355
      Left            =   405
      TabIndex        =   11
      Top             =   1800
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14737
      SectionData     =   "Cetak_9A2.frx":6862
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   12
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
      Left            =   17775
      TabIndex        =   5
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
      Picture         =   "Cetak_9A2.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   10
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
      Picture         =   "Cetak_9A2.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   8
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
      Picture         =   "Cetak_9A2.frx":D33B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17820
      TabIndex        =   9
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
      Picture         =   "Cetak_9A2.frx":10981
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
      Picture         =   "Cetak_9A2.frx":13E60
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   17055
      TabIndex        =   2
      Top             =   990
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
      Picture         =   "Cetak_9A2.frx":1A6C2
      Caption         =   "&s"
      ButtonStyle     =   4
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
      Left            =   3105
      TabIndex        =   21
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL CEK :"
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
      Left            =   540
      TabIndex        =   20
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "History Kunjungan Per Customer"
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
      TabIndex        =   19
      Top             =   135
      Width           =   8970
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   18
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
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
      Left            =   11430
      TabIndex        =   17
      Top             =   1035
      Width           =   5640
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
      Left            =   6165
      TabIndex        =   16
      Top             =   1035
      Width           =   1140
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
      Left            =   5175
      TabIndex        =   15
      Top             =   1080
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
      Left            =   7335
      TabIndex        =   14
      Top             =   1035
      Width           =   4065
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_9A2.frx":1CEF4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_9A2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim kata As String
Dim color As Long, flag As Byte

Private Sub cmdBR_Click()
Customer_br.LBLKODE = "9A2"
Customer_br.Show vbModal
End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
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

Private Sub cmdCLR_Click()

End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Sawal()
End Sub


Private Sub total()

If lblkdcustomer = "" Then
sql1 = "select '1' as kode,isnull(d.nmkolektor,'') as nmkolektor,a.tglLHP,a.kdpiutang,c.bln,c.tahun,c.kdcustomer,b.nmcustomer,b.alamat_TGH as alamat,a.rpLHP,c.TT,N_TT = case when c.TT=1 then 'X' else '' end,a.status,a.keterangan,a.kdlhp,(CASE WHEN DATENAME(dw, a.tglLHP)='Sunday' then 'MINGGU'" & vbCrLf & _
      "WHEN DATENAME(dw, a.tglLHP)='Monday' THEN 'SENIN' WHEN DATENAME(dw, a.tglLHP)='Tuesday' THEN 'SELASA' WHEN DATENAME(dw, a.tglLHP)='Wednesday' THEN 'RABU'" & vbCrLf & _
      "WHEN DATENAME(dw, a.tglLHP)='Thursday' THEN 'KAMIS' WHEN DATENAME(dw, a.tglLHP)='Friday' THEN 'JUMAT' ELSE 'SABTU' END ) as hari  from LHP a " & vbCrLf & _
      "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join customer b on c.kdcustomer=b.kdcustomer  left join kolektor d on a.kdkolektor=d.kdkolektor where a.tglLHP between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' and a.clr=1 "
Else
sql1 = "select '1' as kode,isnull(d.nmkolektor,'') as nmkolektor,a.tglLHP,a.kdpiutang,c.bln,c.tahun,c.kdcustomer,b.nmcustomer,b.alamat_TGH as alamat,a.rpLHP,c.TT,N_TT = case when c.TT=1 then 'X' else '' end,a.status,a.keterangan,a.kdlhp,(CASE WHEN DATENAME(dw, a.tglLHP)='Sunday' then 'MINGGU'" & vbCrLf & _
      "WHEN DATENAME(dw, a.tglLHP)='Monday' THEN 'SENIN' WHEN DATENAME(dw, a.tglLHP)='Tuesday' THEN 'SELASA' WHEN DATENAME(dw, a.tglLHP)='Wednesday' THEN 'RABU'" & vbCrLf & _
      "WHEN DATENAME(dw, a.tglLHP)='Thursday' THEN 'KAMIS' WHEN DATENAME(dw, a.tglLHP)='Friday' THEN 'JUMAT' ELSE 'SABTU' END ) as hari  from LHP a " & vbCrLf & _
      "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join customer b on c.kdcustomer=b.kdcustomer  left join kolektor d on a.kdkolektor=d.kdkolektor where a.tglLHP between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' and c.kdcustomer='" & lblkdcustomer & "'  and a.clr=1 "
End If

sqlT = "select kode,sum(rpLHP) as rpLHP from (" & sql1 & ") a group by kode"
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

MousePointer = vbHourglass

Unload AR_9A2_A
Unload AR_9A2_B

sql = "select a.idrute,a.kdcustomer,a.tglcek,b.nmteknisi,a.nmrute,a.keterangan,a.det_keterangan,c.disp1,c.showC1,c.RG from Real_Cek a left join Teknisi b on a.kdteknisi=b.kdteknisi" & vbCrLf & _
      "left join V_real_CEK c on a.idrute=c.idrute where a.tglcek between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' and a.kdcustomer='" & lblkdcustomer & "' " & vbCrLf & _
      "group by a.idrute,a.kdcustomer,a.tglcek,b.nmteknisi,a.nmrute,a.keterangan,a.det_keterangan,c.disp1,c.showC1,c.RG order by a.tglcek "


Set rs = con.Execute(sql)



With AR_9A2_A.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_9A2_A
.fldtglcek.DataField = "tglcek"
.fldnmteknisi.DataField = "nmteknisi"
.fldnmrute.DataField = "nmrute"
.flddisp.DataField = "disp1"
.fldshow.DataField = "showC1"
.fldrg.DataField = "RG"
.flddet_keterangan.DataField = "det_keterangan"
.fldketerangan.DataField = "keterangan"

.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2
.lblnmCUSTOMER = lblkdcustomer & " - " & lblnmCUSTOMER
.lblalamat = lblalamat


.GroupHeader1.Visible = False

If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True

.fldtglcek.WordWrap = False
.fldnmrute.WordWrap = False
.fldnmteknisi.WordWrap = False
.flddisp.WordWrap = False
.fldshow.WordWrap = False
.fldrg.WordWrap = False
.flddet_keterangan.WordWrap = False
.fldketerangan.WordWrap = False

End If

Set Me.ARV1.ReportSource = AR_9A2_A
End With


MousePointer = vbDefault

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
MousePointer = vbDefault
End Sub


Private Sub Cetak1()
On Error GoTo hell

MousePointer = vbHourglass

Unload AR_9A2_A
Unload AR_9A2_B

sql = "select a.idrute,a.kdcustomer,a.tglcek,b.nmteknisi,a.nmrute,a.keterangan,a.det_keterangan,a.kdbarang,c.kd1,d.nmkategori,c.merk from Real_Cek a left join Teknisi b on a.kdteknisi=b.kdteknisi" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.tglcek between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' and a.kdcustomer='" & lblkdcustomer & "' order by a.tglcek"


Set rs = con.Execute(sql)


With AR_9A2_B.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_9A2_B
.fldtglcek.DataField = "tglcek"
.fldnmteknisi.DataField = "nmteknisi"
.fldnmrute.DataField = "nmrute"
.fldkdbarang.DataField = "kdbarang"
.fldkd1.DataField = "kd1"
.fldnmkategori.DataField = "nmkategori"
.fldmerk.DataField = "merk"
.flddet_keterangan.DataField = "det_keterangan"
.fldketerangan.DataField = "keterangan"

.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2
.lblnmCUSTOMER = lblkdcustomer & " - " & lblnmCUSTOMER
.lblalamat = lblalamat


.GroupHeader1.Visible = False

If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True

.fldtglcek.WordWrap = False
.fldnmteknisi.WordWrap = False
.fldnmrute.WordWrap = False
.fldkdbarang.WordWrap = False
.fldkd1.WordWrap = False
.fldnmkategori.WordWrap = False
.fldmerk.WordWrap = False
.flddet_keterangan.WordWrap = False
.fldketerangan.WordWrap = False

End If

Set Me.ARV1.ReportSource = AR_9A2_B
End With


MousePointer = vbDefault

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
MousePointer = vbDefault
End Sub




Private Sub cmdfs_Click()
If OPT1.Value = True Then
AR_9A2_A.Show vbModal
Else
AR_9A2_B.Show vbModal
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



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
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













