VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form fixrute_list 
   BorderStyle     =   0  'None
   Caption         =   "Fixrute_list"
   ClientHeight    =   10935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   18750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CHKBTA 
      BackColor       =   &H00000000&
      Caption         =   "Format Lampiran BTA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   6525
      TabIndex        =   3
      Top             =   945
      Width           =   2445
   End
   Begin VB.CheckBox chktgl 
      BackColor       =   &H00000000&
      Caption         =   "TANGGAL ROUTE PLAN :"
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
      TabIndex        =   0
      Top             =   1035
      Width           =   2445
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
      Left            =   4770
      TabIndex        =   2
      Top             =   990
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
      Left            =   2835
      TabIndex        =   1
      Top             =   990
      Width           =   1590
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
      TabIndex        =   14
      Top             =   1575
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
   Begin VB.Timer TimerCetak 
      Left            =   13275
      Top             =   2340
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
      Picture         =   "fixrute_list.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   11
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
      TabIndex        =   8
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
      Picture         =   "fixrute_list.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   6
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
      Picture         =   "fixrute_list.frx":9A49
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17820
      TabIndex        =   7
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
      Picture         =   "fixrute_list.frx":D08F
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   10
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
      Picture         =   "fixrute_list.frx":1056E
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8715
      Left            =   360
      TabIndex        =   9
      Top             =   1485
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   15372
      SectionData     =   "fixrute_list.frx":16DD0
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   17730
      TabIndex        =   4
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
      Picture         =   "fixrute_list.frx":16E0C
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
      Left            =   4410
      TabIndex        =   15
      Top             =   1035
      Width           =   420
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   13
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Route Plan"
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
      TabIndex        =   12
      Top             =   135
      Width           =   4560
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "fixrute_list.frx":1A6C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "fixrute_list"
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

Private Sub cmdUpload_Click()
Call con_mysql
sql1 = "select * from tbfixrute"
Set rs1 = con1.Execute(sql1)
Set datagrid1.DataSource = rs1
End Sub

Private Sub cmdGO_Click()
If CHKBTA.Value = 0 Then
Call Cetak
Else
Call Cetak1
End If
End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
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
If CHKBTA.Value = 0 Then
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
'On Error GoTo hell


Unload AR_fixrute_list

'planing
sqlQ = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
            "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
            "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
            "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= '" & Format(fixrute_TU.txttglspk1, "yyyy/MM/dd") & "' group by kdcustomer,kdkategori" & vbCrLf & _
            ") a group by kdcustomer " & vbCrLf & _
       ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2 +RG <>0"
       
'realisasi
sqlR = "select * from V_real_cek where kdteknisi='" & fixrute_TU.lblkdteknisi & "' and nmrute='" & fixrute_TU.txtperiode & "'"
       

'tidak dimasukkan ke fixrute
sqlS = "select kdcustomer from route_plan where kdteknisi='" & fixrute_TU.lblkdteknisi & "' and nmrute='" & fixrute_TU.txtperiode & "' union all select kdcustomer from V_real_cek where nmrute='" & fixrute_TU.txtperiode & "'"


'belum dikunjungi
sqlA1 = "select a.idrute,a.tglplan,d.tglcek,c.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,a.keterangan,a.det_keterangan,a.jmlunit,isnull(d.unit,0) as U_kunjungan,a.tglinput  from ROUTE_PLAN a left join Customer b " & vbCrLf & _
        "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (select idrute,min(tglcek) as tglcek,sum(unit) as unit from real_cek group by idrute) d on a.idrute=d.idrute where a.kdteknisi='" & fixrute_TU.lblkdteknisi & "' and  nmrute= '" & fixrute_TU.txtperiode & "' "
       
sqlA = "select *,'BLM DIKUNJUNGI' as Status from (" & sqlA1 & " ) a  where U_kunjungan = 0"


'Sudah dikunjungi


sqlB1 = "select a.idrute,a.tglplan,d.tglcek,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,d.keterangan,d.det_keterangan,a.jmlunit,isnull(d.unit,0) as U_kunjungan,a.tglinput  from ROUTE_PLAN a left join Customer b " & vbCrLf & _
        "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (select idrute,min(tglcek) as tglcek,sum(unit) as unit,keterangan,det_keterangan from real_cek group by idrute,keterangan,det_keterangan) d on a.idrute=d.idrute where a.kdteknisi='" & fixrute_TU.lblkdteknisi & "' and  nmrute= '" & fixrute_TU.txtperiode & "'"


sqlB = "select *,'SDH DIKUNJUNGI' as status  from (" & sqlB1 & ") a  where U_kunjungan <> 0 "


'Non Route

sqlC1 = "select idrute from route_plan where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi='" & fixrute_TU.lblkdteknisi & "'"
    
sqlC2 = "select a.idrute,a.tglcek as tglplan,a.tglcek,C.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,'' as keterangan,'' as det_keterangan,0 as jmlunit,sum(a.unit) as U_kunjungan,a.tglinput from" & vbCrLf & _
       "Real_Cek a left join customer b on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareac=c.kdareaC where (a.kdteknisi='" & fixrute_TU.lblkdteknisi & "' and  a.nmrute= '" & fixrute_TU.txtperiode & "') and a.idrute not in (" & sqlC1 & ") group by" & vbCrLf & _
       "a.idrute , a.tglcek, C.nmareaC, a.kdcustomer, b.nmcustomer,b.alamat, b.cp, b.telp,a.tglinput"
    
sqlC = "select *,'NON ROUTE' as status from (" & sqlC2 & ") a  where U_kunjungan <> 0 "




'hasil
sql1 = "select '1' as kode,a.* ,b.disp,b.showc,b.RG,isnull(c.disp1,0) as disp1,isnull(c.showc1,0) as showc1,isnull(c.RG,0) as RG1 from (" & sqlA & " union all " & sqlB & " ) a left join (" & sqlQ & ") b  on a.kdcustomer=b.kdcustomer left join (" & sqlR & ") c on  a.kdcustomer=c.kdcustomer "

sql2 = "select '1' as kode,a.* ,0 as disp,0 as showc,0 as RG,isnull(c.disp1,0) as disp1,isnull(c.showc1,0) as showc1,isnull(c.RG,0) as RG1 from (" & sqlC & " ) a left join (" & sqlR & ") c on  a.kdcustomer=c.kdcustomer "

sql3 = "select '1' as kode,'-' as idrute,'1900/01/01' as tglplan,'1900/01/01' as tglCEK,C.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,'' as keterangan,'' as det_keterangan,0 as jmlunit,0 as U_kunjungan,'1900/01/01' as tglinput,'BLM MSK ROUTE',a.disp,a.showc,a.rg,0 as disp1,0 as showc1, 0 as rg1 from" & vbCrLf & _
       "(" & sqlQ & ") a left join customer b on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareac=c.kdareaC where b.kdteknisi='" & fixrute_TU.lblkdteknisi & "' and a.kdcustomer not in (" & sqlS & ") "

'sql4 = "select '1' as kode,a.idrute_S as idrute,a.tglrute_S as tglplan,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,'' as keterangan,0 as jmlunit,0 as u_kunjungan,'BLM DIKUNJUNGI' as status,0 as disp,0 as showc,0 as RG,0 as disp1,0 as showc1,0 as RG1 from ROUTE_PLAN_S a left join Customer b " & vbCrLf & _
'        "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC where (a.kdteknisi='" & fixrute_TU.lblkdteknisi & "' and  a.nmrute= '" & fixrute_TU.txtperiode & "') and a.kdcustomer not in (select kdcustomer from real_cek where kdteknisi='" & fixrute_TU.lblkdteknisi & "' and  nmrute= '" & fixrute_TU.txtperiode & "')"

If chktgl.Value = 0 Then
sql = "select * from (" & sql1 & " union all " & sql2 & " union all " & sql3 & "  ) a order by tglplan,tglinput,nmcustomer,alamat "
sqlT = "select kode,sum(disp) as disp,sum(showc) as showC,sum(rg) as rg,sum(disp1) as disp1,sum(showc1) as showC1,sum(rg1) as rg1  from (" & sql1 & " union all " & sql2 & " union all " & sql3 & "  ) a group by kode"
Else
sql = "select * from (" & sql1 & " union all " & sql2 & " union all " & sql3 & "  ) a where tglplan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' order by tglplan,tglinput,nmcustomer,alamat "
sqlT = "select kode,sum(disp) as disp,sum(showc) as showC,sum(rg) as rg,sum(disp1) as disp1,sum(showc1) as showC1,sum(rg1) as rg1  from (" & sql1 & " union all " & sql2 & " union all " & sql3 & " ) a where tglplan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by kode"
End If


Set rs = con.Execute(sqlT)



With AR_fixrute_list.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_fixrute_list
.fldnmcustomer.DataField = "nmcustomer"
.fldkdcustomer.DataField = "kdcustomer"
.fldalamat.DataField = "alamat"
.fldtglplan.DataField = "tglplan"
.fldnmareaC.DataField = "nmareaC"
.fldshowC.DataField = "showC"
.flddisp.DataField = "disp"
.fldRG.DataField = "RG"
.fldshowc1.DataField = "showC1"
.flddisp1.DataField = "disp1"
.fldRG1.DataField = "RG1"
.fldstatus.DataField = "status"
.fldketerangan.DataField = "keterangan"
.flddet_keterangan.DataField = "det_keterangan"
.fldtglCek.DataField = "tglcek"


.lblnmteknisi = fixrute_TU.lblnmteknisi
.lbltgl1 = fixrute_TU.txttglspk1
.lblcetak = Format(Date, "dd/MM/yyyy")

If rs.RecordCount <> 0 Then
.lbldisp = FormatNumber(rs!disp, 0)
.lblshowC = FormatNumber(rs!showC, 0)
.lblRG = FormatNumber(rs!RG, 0)
.lbldisp1 = FormatNumber(rs!disp1, 0)
.lblshowc1 = FormatNumber(rs!showC1, 0)
.lblRG1 = FormatNumber(rs!RG1, 0)

Else
.lbldisp = "0"
.lblshowC = "0"
.lblRG = "0"
.lbldisp1 = "0"
.lblshowc1 = "0"
.lblRG1 = "0"
End If

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
.flddisp.WordWrap = False
.fldshowC.WordWrap = False
.flddisp1.WordWrap = False
.fldshowc1.WordWrap = False
.fldstatus.WordWrap = False
.fldNO.WordWrap = False
.fldRG.WordWrap = False
.fldRG1.WordWrap = False
.fldNO.WordWrap = False
.lbldisp.WordWrap = False
.lblshowC.WordWrap = False
.lblRG.WordWrap = False
.lbldisp1.WordWrap = False
.lblshowc1.WordWrap = False
.lblRG1.WordWrap = False
.fldketerangan.WordWrap = False
.flddet_keterangan.WordWrap = False
.fldNO.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_fixrute_list
End With
'''
''
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !"
End Sub


Private Sub Cetak1()
'On Error GoTo hell


Unload AR_Fixrute_list1

'planing
sqlQ = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
            "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
            "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
            "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= '" & Format(fixrute_TU.txttglspk1, "yyyy/MM/dd") & "' group by kdcustomer,kdkategori" & vbCrLf & _
            ") a group by kdcustomer " & vbCrLf & _
       ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2 +RG <>0"
       


'belum dikunjungi
sqlA = "select a.tglplan,c.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,d.jam_in,d.jam_out,a.tglinput from ROUTE_PLAN a left join Customer b " & vbCrLf & _
       "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (select * from jam_kunjungan where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi='" & fixrute_TU.lblkdteknisi & "') d on a.kdcustomer=d.kdcustomer and a.nmrute=d.nmrute and a.kdteknisi=d.kdteknisi where a.nmrute='" & fixrute_TU.txtperiode & "' and a.kdteknisi='" & fixrute_TU.lblkdteknisi & "' "
       
'hasil
sql1 = "select '1' as kode,a.* ,b.disp,b.showc,b.RG from (" & sqlA & " ) a left join (" & sqlQ & ") b  on a.kdcustomer=b.kdcustomer"

If chktgl.Value = 0 Then
sql = "select * from (" & sql1 & " ) a order by tglplan,tglinput,jam_in"
sqlT = "select kode,sum(disp) as disp,sum(showc) as showC,sum(rg) as rg  from (" & sql1 & ") a group by kode"
Else
sql = "select * from (" & sql1 & ") a where tglplan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' order by tglplan,tglinput,jam_in "
sqlT = "select kode,sum(disp) as disp,sum(showc) as showC,sum(rg) as rg from (" & sql1 & " ) a where tglplan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by kode"
End If


Set rs = con.Execute(sqlT)



With AR_Fixrute_list1.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_Fixrute_list1
.fldnmcustomer.DataField = "nmcustomer"
.fldkdcustomer.DataField = "kdcustomer"
.fldalamat.DataField = "alamat"
.fldtglplan.DataField = "tglplan"
.fldnmareaC.DataField = "nmareaC"
.fldshowC.DataField = "showC"
.flddisp.DataField = "disp"
.fldRG.DataField = "RG"
.fldin.DataField = "jam_in"
.fldout.DataField = "jam_out"


.lblnmteknisi = fixrute_TU.lblnmteknisi
.lbltgl1 = fixrute_TU.txttglspk1
.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")

If rs.RecordCount <> 0 Then
.lbldisp = FormatNumber(rs!disp, 0)
.lblshowC = FormatNumber(rs!showC, 0)
.lblRG = FormatNumber(rs!RG, 0)

Else
.lbldisp = "0"
.lblshowC = "0"
.lblRG = "0"
End If

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
.fldnmareaC.WordWrap = False
.flddisp.WordWrap = False
.fldshowC.WordWrap = False
.fldNO.WordWrap = False
.fldRG.WordWrap = False
.lbldisp.WordWrap = False
.lblshowC.WordWrap = False
.lblRG.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_Fixrute_list1
End With
'''
''
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !"
End Sub






Private Sub cmdfs_Click()
If CHKBTA.Value = 0 Then
AR_fixrute_list.Show vbModal
Else
AR_Fixrute_list1.Show vbModal
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



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttgl1 = Date
txttgl2 = Date

txttgl1.Enabled = False
txttgl2.Enabled = False

TimerCetak.Interval = 10

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






