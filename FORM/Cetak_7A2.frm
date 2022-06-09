VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_7A2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18795
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   18795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   270
      TabIndex        =   4
      Top             =   1620
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
      Left            =   1350
      TabIndex        =   5
      Top             =   1620
      Width           =   1050
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
      Left            =   1530
      TabIndex        =   12
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Timer TimerPdf 
      Left            =   14805
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13860
      Top             =   2295
   End
   Begin VB.Timer Timerxls 
      Left            =   14310
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
      Left            =   9720
      TabIndex        =   6
      Top             =   2115
      Width           =   555
   End
   Begin VB.CheckBox Chk2 
      BackColor       =   &H00000000&
      Caption         =   "TAMPILKAN NAMA CUSTOMER YG SAMA"
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
      Height          =   600
      Left            =   15390
      TabIndex        =   2
      Top             =   1125
      Width           =   2085
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   15975
      TabIndex        =   7
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
      Picture         =   "Cetak_7A2.frx":0000
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
      Left            =   17685
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
      Picture         =   "Cetak_7A2.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17730
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
      Picture         =   "Cetak_7A2.frx":A118
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17730
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
      Picture         =   "Cetak_7A2.frx":D2FF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17730
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
      Picture         =   "Cetak_7A2.frx":10945
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
      Picture         =   "Cetak_7A2.frx":13E24
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   14265
      TabIndex        =   0
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
      Picture         =   "Cetak_7A2.frx":1A686
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCLR 
      Height          =   420
      Left            =   14760
      TabIndex        =   1
      ToolTipText     =   "Kosongi customer untuk menampilkan semuanya"
      Top             =   1215
      Width           =   555
      _ExtentX        =   979
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
      Picture         =   "Cetak_7A2.frx":1CEB8
      ButtonStyle     =   4
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8130
      Left            =   270
      TabIndex        =   11
      Top             =   1980
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14340
      SectionData     =   "Cetak_7A2.frx":1F502
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
      Left            =   315
      TabIndex        =   21
      Top             =   1305
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rekap Pinjaman dan Sewa"
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
      TabIndex        =   20
      Top             =   135
      Width           =   7665
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10440
      TabIndex        =   19
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
      Left            =   9495
      TabIndex        =   18
      Top             =   1260
      Width           =   4785
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
      Left            =   4230
      TabIndex        =   17
      Top             =   1260
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
      Left            =   3240
      TabIndex        =   16
      Top             =   1305
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
      Left            =   5400
      TabIndex        =   15
      Top             =   1260
      Width           =   4065
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_7A2.frx":1F53E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_7A2"
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


Private Sub Chk2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub


Private Sub cmdBR_Click()
Customer_br.lblkode = "7A2"
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
lblkdcustomer = ""
lblnmcustomer = ""
lblalamat = ""
End Sub

Private Sub cmdCLR_KeyPress(KeyAscii As Integer)
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




Private Sub total()

sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"


If lblkdcustomer = "" Then
sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and c.kdkategori between '04' and '10'"
Else
    If Chk2.Value = 0 Then
    sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and a.kdcustomer='" & lblkdcustomer & "' and c.kdkategori between '04' and '10'"
    Else
    sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and b.nmcustomer='" & lblnmcustomer & "' and c.kdkategori between '04' and '10'"
    End If
End If

sqlB = "select kdcustomer,nmcustomer,alamat,sum(case kdkategori when '04' then pjm else 0 end) as P1,sum(case kdkategori when '05' then pjm else 0 end) as P2,sum(case kdkategori when '06' then pjm else 0 end) as P3,sum(case kdkategori when '07' then pjm else 0 end) as P4,sum(case kdkategori when '08' then pjm else 0 end) as P5,sum(case kdkategori when '09' then pjm else 0 end) as P6,sum(case kdkategori when '10' then pjm else 0 end) as P7," & vbCrLf & _
      "sum(case kdkategori when '04' then swa else 0 end) as S1,sum(case kdkategori when '05' then swa else 0 end) as S2,sum(total) as total from (" & sqlA & ") a group by kdcustomer,nmcustomer,alamat"

sqlX = "select '1' as kode,a.*,(a.p1+a.p2+a.p3+a.p4+a.p5+a.p6+a.p7) as P_total,(a.S1+a.S2) as S_total from (" & sqlB & ") a "


sqlT = "select kode,sum(p1) as p1,sum(p2) as p2,sum(p3) as p3,sum(p4) as p4,sum(p5) as p5,sum(p6) as p6,sum(p7) as p7,sum(p_total) as p_total,sum(S1) as S1,sum(S2) as S2,sum(S_total) as S_total, sum(total) as total from (" & sqlX & ") a group by kode"
Set rs = con.Execute(sqlT)

End Sub


Private Sub total1()

sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"


If lblkdcustomer = "" Then
sqlA = "select '1' as kode,a.kdcustomer,b.nmcustomer,b.alamat,b.nospk,b.tglspk1,b.tglspk2,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,d.nmkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.pjm+a.swa <>0 and c.kdkategori between '04' and '10'"
Else
    If Chk2.Value = 0 Then
    sqlA = "select '1' as kode,a.kdcustomer,b.nmcustomer,b.alamat,b.nospk,b.tglspk1,b.tglspk2,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,d.nmkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
           "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.pjm+a.swa <>0 and a.kdcustomer='" & lblkdcustomer & "' and c.kdkategori between '04' and '10'"
    Else
    sqlA = "select '1' as kode,a.kdcustomer,b.nmcustomer,b.alamat,b.nospk,b.tglspk1,b.tglspk2,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,d.nmkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
           "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.pjm+a.swa <>0 and b.nmcustomer='" & lblnmcustomer & "' and c.kdkategori between '04' and '10'"
    End If
End If

sqlT = "select kode,sum(pjm) as pjm,sum(swa) as swa from (" & sqlA & ") a group by kode"
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
On Error GoTo hell

Unload AR_7A2_01
Unload AR_7A2

sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"


If lblkdcustomer = "" Then
sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and c.kdkategori between '04' and '10'"
Else
    If Chk2.Value = 0 Then
    sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and a.kdcustomer='" & lblkdcustomer & "' and c.kdkategori between '04' and '10'"
    Else
    sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
          "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and b.nmcustomer='" & lblnmcustomer & "' and c.kdkategori between '04' and '10'"
    End If
End If

sqlB = "select kdcustomer,nmcustomer,alamat,sum(case kdkategori when '04' then pjm else 0 end) as P1,sum(case kdkategori when '05' then pjm else 0 end) as P2,sum(case kdkategori when '06' then pjm else 0 end) as P3,sum(case kdkategori when '07' then pjm else 0 end) as P4,sum(case kdkategori when '08' then pjm else 0 end) as P5,sum(case kdkategori when '09' then pjm else 0 end) as P6,sum(case kdkategori when '10' then pjm else 0 end) as P7," & vbCrLf & _
      "sum(case kdkategori when '04' then swa else 0 end) as S1,sum(case kdkategori when '05' then swa else 0 end) as S2,sum(total) as total from (" & sqlA & ") a group by kdcustomer,nmcustomer,alamat"

sql = "select a.*,(a.p1+a.p2+a.p3+a.p4+a.p5+a.p6+a.p7) as P_total,(a.S1+a.S2) as S_total from (" & sqlB & ") a order by nmcustomer,alamat"

With AR_7A2.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_7A2
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcus.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldP1.DataField = "p1"
.fldP2.DataField = "p2"
.fldP3.DataField = "p3"
.fldP4.DataField = "p4"
.fldP5.DataField = "p5"
.fldP6.DataField = "p6"
.fldP7.DataField = "p7"
.fldS1.DataField = "S1"
.fldS2.DataField = "S2"
.fldP_total.DataField = "P_total"
.fldS_total.DataField = "S_total"
.fldtotal.DataField = "Total"


.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbljudul = "REKAP PINJAMAN & SEWA"


Call total
If rs.RecordCount <> 0 Then
.lblP1 = Format(rs!p1, "#,###0")
.lblP2 = Format(rs!p2, "#,###0")
.lblP3 = Format(rs!p3, "#,###0")
.lblP4 = Format(rs!p4, "#,###0")
.lblP5 = Format(rs!p5, "#,###0")
.lblP6 = Format(rs!p6, "#,###0")
.lblP7 = Format(rs!p7, "#,###0")
.lblP_total = Format(rs!p_total, "#,###0")
.lblS1 = Format(rs!S1, "#,###0")
.lblS2 = Format(rs!S2, "#,###0")
.lblS_total = Format(rs!S_total, "#,###0")
.lbltotal = Format(rs!total, "#,###0")
Else
.lblP1 = 0
.lblP2 = 0
.lblP3 = 0
.lblP4 = 0
.lblP5 = 0
.lblP6 = 0
.lblP7 = 0
.lblP_total = 0
.lblS1 = 0
.lblS2 = 0
.lblS_total = 0
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

.fldkdcustomer.WordWrap = False
.fldnmcus.WordWrap = False
.fldalamat.WordWrap = False
.fldP1.WordWrap = False
.fldP2.WordWrap = False
.fldP3.WordWrap = False
.fldP4.WordWrap = False
.fldP5.WordWrap = False
.fldP6.WordWrap = False
.fldP7.WordWrap = False
.fldS1.WordWrap = False
.fldS2.WordWrap = False
.fldP_total.WordWrap = False
.fldS_total.WordWrap = False
.fldtotal.WordWrap = False
.fldNO.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_7A2
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub


Private Sub Cetak1()
On Error GoTo hell


Unload AR_7A2
Unload AR_7A2_01

sqlX = "select a.kdbarang,max(b.tglbpb) as tglbpb from beli_d a left join beli b on a.kdbeli=b.kdbeli where a.unit > 0 and b.tglbpb <= '" & Format(txttgl1, "yyyy/MM/dd") & "' group by a.kdbarang"

sqlY = "select kdbarang,max(tglsj) as tglsj from (" & vbCrLf & _
       "select a.kdbarang,b.tglpinjam as tglSJ from pinjam_d a left join pinjam b on a.kdpinjam=b.kdpinjam Union all select a.kdbarang,b.tglsewa as tglSJ from sewa_d a left join sewa b on a.kdsewa=b.kdsewa" & vbCrLf & _
       ") a where tglsj <= '" & Format(txttgl1, "yyyy/MM/dd") & "' group by kdbarang"

sqlZ = "select a.kdpinjam as nobukti,a.nosj as nosj,tglpinjam as tgl1, a.kdcustomer,b.kdbarang,b.unit as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' " & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdsewa as nobukti,a.nosj as nosj,tglsewa as tgl1,a.kdcustomer,b.kdbarang,0 as pjm,b.unit as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  "


sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"


If lblkdcustomer = "" Then
sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,b.nospk,b.tglspk1,b.tglspk2,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,d.nmkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.pjm+a.swa <>0 and c.kdkategori between '04' and '10'"
Else
    If Chk2.Value = 0 Then
    sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,b.nospk,b.tglspk1,b.tglspk2,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,d.nmkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
           "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.pjm+a.swa <>0 and a.kdcustomer='" & lblkdcustomer & "' and c.kdkategori between '04' and '10'"
    Else
    sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,b.nospk,b.tglspk1,b.tglspk2,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,d.nmkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
           "left join barang c on a.kdbarang=c.kdbarang left join kategoriBRG d on c.kdkategori=d.kdkategori where a.pjm+a.swa <>0 and b.nmcustomer='" & lblnmcustomer & "' and c.kdkategori between '04' and '10'"
    End If
End If

sqlB = "select a.kdbarang,a.nmbarang,a.kdkategori,a.nmkategori,a.kd1,b.tglbpb,a.kdcustomer,a.nmcustomer,a.alamat,a.nospk,a.tglspk1,a.tglspk2,c.tglsj,a.pjm,a.swa,a.total from " & vbCrLf & _
       "(" & sqlA & ") a left join (" & sqlX & ") b on a.kdbarang=b.kdbarang left join (" & sqlY & ") c on a.kdbarang=c.kdbarang "
       
sql = "select a.* from (" & sqlB & ") a order by a.nmcustomer,a.alamat "



With AR_7A2_01.DC1
.ConnectionString = koneksi
.Source = sql
End With
''
With AR_7A2_01
.fldkdbarang.DataField = "kdbarang"
.fldnmbarang.DataField = "nmbarang"
.fldkd1.DataField = "kd1"
.fldkdkategori.DataField = "kdkategori"
.fldnmkategori.DataField = "nmkategori"
.fldnoSPK.DataField = "noSPK"
.fldtglSPK1.DataField = "tglspk1"
.fldtglSPK2.DataField = "tglspk2"
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcus.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldtglbpb.DataField = "tglBPB"
.fldtglSJ.DataField = "tglSJ"
.fldpjm.DataField = "pjm"
.fldswa.DataField = "swa"

'

.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1

Call total1
If rs.RecordCount <> 0 Then
.lblpjm = Format(rs!pjm, "#,###0")
.lblswa = Format(rs!swa, "#,###0")
Else
.lblpjm = 0
.lblswa = 0
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

.fldkdbarang.WordWrap = False
.fldnmbarang.WordWrap = False
.fldkd1.WordWrap = False
.fldkdkategori.WordWrap = False
.fldnmkategori.WordWrap = False
.fldnoSPK.WordWrap = False
.fldtglSPK1.WordWrap = False
.fldtglSPK2.WordWrap = False
.fldkdcustomer.WordWrap = False
.fldnmcus.WordWrap = False
.fldalamat.WordWrap = False
.fldtglbpb.WordWrap = False
.fldtglSJ.WordWrap = False
.fldpjm.WordWrap = False
.fldswa.WordWrap = False
.fldNO.WordWrap = False
.lblpjm.WordWrap = False
.lblswa.WordWrap = False
End If

Set Me.ARV1.ReportSource = AR_7A2_01
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub









Private Sub cmdfs_Click()
If Opt1.Value = True Then
AR_7A2.Zoom = 110
AR_7A2.Show vbModal

Else
AR_7A2_01.Zoom = 110
AR_7A2_01.Show vbModal
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
If Opt1.Value = True Then
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

Opt1.Value = True

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







