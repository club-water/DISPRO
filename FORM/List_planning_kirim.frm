VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form List_planning_kirim 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10920
   ScaleWidth      =   18825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChKF 
      BackColor       =   &H00000000&
      Caption         =   "Format Laporan Pengiriman"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   14400
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   1305
      Width           =   2985
   End
   Begin VB.ComboBox CMbkolom 
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
      Left            =   12735
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1260
      Width           =   690
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
      Top             =   2160
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
      Left            =   2160
      TabIndex        =   0
      Top             =   1260
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
      Left            =   4095
      TabIndex        =   1
      Top             =   1260
      Width           =   1590
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16065
      TabIndex        =   7
      Top             =   2115
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
      Picture         =   "List_planning_kirim.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8220
      Left            =   360
      TabIndex        =   11
      Top             =   2025
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14499
      SectionData     =   "List_planning_kirim.frx":6862
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
      Picture         =   "List_planning_kirim.frx":689E
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
      Picture         =   "List_planning_kirim.frx":A154
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
      Picture         =   "List_planning_kirim.frx":D33B
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
      Picture         =   "List_planning_kirim.frx":10981
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
      Picture         =   "List_planning_kirim.frx":13E60
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   10485
      TabIndex        =   2
      ToolTipText     =   "Simpan"
      Top             =   1215
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
      Picture         =   "List_planning_kirim.frx":1A6C2
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "KOLOM KOSONG :"
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
      Left            =   11205
      TabIndex        =   22
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "BARIS"
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
      Left            =   13500
      TabIndex        =   21
      Top             =   1350
      Width           =   600
   End
   Begin VB.Label lblnmteknisi 
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
      Left            =   7560
      TabIndex        =   20
      Top             =   1260
      Width           =   2940
   End
   Begin VB.Label lblkdteknisi 
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
      Left            =   6660
      TabIndex        =   19
      Top             =   1260
      Width           =   870
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SOPIR :"
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
      Height          =   465
      Left            =   5850
      TabIndex        =   18
      Top             =   1305
      Width           =   825
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   17
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lbljudul 
      BackStyle       =   0  'Transparent
      Caption         =   "List Customer"
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
      Top             =   90
      Width           =   12705
   End
   Begin VB.Label lbltgl 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL RENCANA KIRIM :"
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
      TabIndex        =   15
      Top             =   1305
      Width           =   1905
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
      Left            =   3735
      TabIndex        =   14
      Top             =   1305
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   45
      Picture         =   "List_planning_kirim.frx":1CEF4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "List_planning_kirim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte
Dim kata_masalah, kata_sisa As String
Dim kata, kata1 As String



Private Sub CMbkolom_Click()
If CMbkolom.ListIndex = 0 Then
kata = "convert(int,kode1) < 2"
kata1 = "convert(int,urut1) < 3"
ElseIf CMbkolom.ListIndex = 1 Then
kata = "convert(int,kode1) < 3"
kata1 = "convert(int,urut1) < 4"
ElseIf CMbkolom.ListIndex = 2 Then
kata = "convert(int,kode1) < 4"
kata1 = "convert(int,urut1) < 5"
ElseIf CMbkolom.ListIndex = 3 Then
kata = "convert(int,kode1) < 5"
kata1 = "convert(int,urut1) < 6"
ElseIf CMbkolom.ListIndex = 4 Then
kata = "convert(int,kode1) < 6"
kata1 = "convert(int,urut1) < 7"
ElseIf CMbkolom.ListIndex = 5 Then
kata = "convert(int,kode1) < 7"
kata1 = "convert(int,urut1) < 8"

End If

End Sub

Private Sub cmdBR1_Click()
Teknisi_BR.LBLKODE = "LIST_PLANNING_KIRIM"
Teknisi_BR.Show vbModal

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

End Sub




 





Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
If ChKF.Value = 0 Then
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


Unload AR_JADWAL_KIRIM

If Planning_kirim.Opt1.Value = False Then
sql1 = "select '1' as kode1,a.kdteknisi,c.nmteknisi, a.tglPK,b.*,a.Uraian from planning_kirim a left join V_ALL_PO_RETUR b on a.kdPK=b.kode left join teknisi c on a.kdteknisi=c.kdteknisi where a.tglPK between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' and a.kdteknisi='" & lblkdteknisi & "' Union All" & vbCrLf & _
       "select '2' as kode1,'','',getdate(),'',getdate(),'','','','','','',0,0,0,'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx' Union All " & vbCrLf & _
       "select '3' as kode1,'','',getdate(),'',getdate(),'','','','','','',0,0,0,'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx' Union All " & vbCrLf & _
       "select '4' as kode1,'','',getdate(),'',getdate(),'','','','','','',0,0,0,'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx' Union All " & vbCrLf & _
       "select '5' as kode1,'','',getdate(),'',getdate(),'','','','','','',0,0,0,'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx' Union All " & vbCrLf & _
       "select '6' as kode1,'','',getdate(),'',getdate(),'','','','','','',0,0,0,'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'  "
       
sql = "select * from (" & sql1 & ") x where " & kata & " order by kode1,tglPK,nmteknisi"
Else
sql1 = "select '1' as kode1,'' as kdteknisi,'' as nmteknisi, '' as tglPK,*,'' as uraian from V_tanggungan_kirim where kode not in (select kdPK from planning_kirim) "
sql = "select * from (" & sql1 & ") x where tglpengajuan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' order by tglPengajuan"
End If

With AR_JADWAL_KIRIM.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_JADWAL_KIRIM

If Planning_kirim.Opt1.Value = False Then
.fldtglPK.DataField = "tglPK"
.lbljudul = "RENCANA PENGIRIMAN"
.lbltgl = "TGL KIRIM"
.lbltglX = "TGL KIRIM"

Else
.lbljudul = "OUTSTANDING PENGIRIMAN"
.fldtglPK.DataField = "tglpengajuan"
.lbltgl = "PENGAJUAN"
.lbltglX = "PENGAJUAN"
End If

.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldnmkategori.DataField = "nmkategori"
.fldketerangan.DataField = "keterangan"
.fldnmarea.DataField = "nmareaC"
'.flduraian.DataField = "Uraian"
.fldQty_DISP.DataField = "qty_Disp"
.fldQty_SHW.DataField = "qty_sh"
.FldQTY_lain.DataField = "qty_lain"



.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2
.GroupHeader1.Visible = False
.lblnmteknisi = lblnmteknisi

'
If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True

.fldtglPK.WordWrap = False
.fldnmcustomer.WordWrap = False
.fldalamat.WordWrap = False
.fldnmkategori.WordWrap = False
.fldketerangan.WordWrap = False
.flduraian.WordWrap = False
.fldnmarea.WordWrap = False
.fldQty_DISP.WordWrap = False
.fldQty_SHW.WordWrap = False
.FldQTY_lain.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_JADWAL_KIRIM
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub Cetak1()
On Error GoTo hell


Unload AR_JADWAL_KIRIM
Unload AR_LAP_PENGIRIMAN

       
sql1 = "select '1'as  urut1,b.nosj, a.kdpk,a.tglPk,c.nmcustomer + ' - ' + c.alamat as nmcustomer,c.nmkategori + ' - ' + c.keterangan as keterangan ," & vbCrLf & _
       "b.kdbarang1, b.kdbarang2 from Planning_kirim a left join V_Detail_Barang_kiriman b on a.kdPK=b.kode left join V_ALL_PO_RETUR c on a.kdPK=c.kode where a.kdteknisi='" & lblkdteknisi & "' and a.tglpk between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "'"

sql2 = "select '2'as  urut1,'' as nosj, a.kdpk,a.tglPk,c.nmcustomer + ' - ' + c.alamat as nmcustomer,c.nmkategori + ' - ' + c.keterangan as keterangan ," & vbCrLf & _
       "'' as kdbarang1, '' as kdbarang2 from Planning_kirim a left join V_ALL_PO_RETUR c on a.kdPK=c.kode where a.kdteknisi='" & lblkdteknisi & "' and a.tglpk between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' "
       
sql3 = "select '3' as urut1,'','Z',getdate(),'','','','' Union All " & vbCrLf & _
       "select '4' as urut1,'','Z',getdate(),'','','','' Union All " & vbCrLf & _
       "select '5' as urut1,'','Z',getdate(),'','','','' Union All " & vbCrLf & _
       "select '6' as urut1,'','Z',getdate(),'','','','' Union All " & vbCrLf & _
       "select '7' as urut1,'','Z',getdate(),'','','','' "


sql = "select row_number() over(partition by kdpk order by kdpk,tglpk,urut1) as urut, * from (" & sql1 & " union all " & sql2 & " union all " & sql3 & ") x where " & kata1 & ""


With AR_LAP_PENGIRIMAN.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_LAP_PENGIRIMAN


.fldkdPK.DataField = "nosj"
.fldcustomer.DataField = "nmcustomer"
.fldketerangan.DataField = "keterangan"
.fldkdbarang1.DataField = "kdbarang1"
.fldkdbarang2.DataField = "kdbarang2"
.fldurut.DataField = "urut"




.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2
.GroupHeader1.Visible = False
.lblnmteknisi = lblnmteknisi

'
If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True

.fldkdPK.WordWrap = False
.fldcustomer.WordWrap = False
.fldketerangan.WordWrap = False
.fldkdbarang1.WordWrap = False
.fldkdbarang2.WordWrap = False
.fldurut.WordWrap = False




End If

Set Me.ARV1.ReportSource = AR_LAP_PENGIRIMAN
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





Private Sub cmdfs_Click()
If ChKF.Value = 0 Then
AR_JADWAL_KIRIM.Show vbModal
Else
AR_LAP_PENGIRIMAN.Show vbModal
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
If ChKF.Value = 0 Then
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

txttgl1 = Date
txttgl2 = Date

CMbkolom.AddItem "0"
CMbkolom.AddItem "1"
CMbkolom.AddItem "2"
CMbkolom.AddItem "3"
CMbkolom.AddItem "4"
CMbkolom.AddItem "5"

CMbkolom.ListIndex = 0

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















