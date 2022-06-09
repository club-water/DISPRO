VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_frm_kunjungan 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10920
   ScaleWidth      =   18750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbbln3 
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
      Left            =   12555
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1485
      Width           =   690
   End
   Begin VB.ComboBox cmbtahun3 
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
      Left            =   14130
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1485
      Width           =   1095
   End
   Begin VB.ComboBox cmbbln2 
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
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1485
      Width           =   690
   End
   Begin VB.ComboBox cmbtahun2 
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
      Left            =   8775
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1485
      Width           =   1095
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
      Left            =   3420
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1485
      Width           =   1095
   End
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
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1485
      Width           =   690
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
      Left            =   14310
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Width           =   690
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
      Left            =   11025
      TabIndex        =   2
      Top             =   1035
      Width           =   1590
   End
   Begin VB.Timer Timerxls 
      Left            =   14355
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13905
      Top             =   2295
   End
   Begin VB.Timer TimerPdf 
      Left            =   14850
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
      Left            =   9135
      TabIndex        =   1
      Top             =   1035
      Width           =   1590
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16020
      TabIndex        =   14
      Top             =   2025
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
      Picture         =   "Cetak_frm_kunjungan.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8175
      Left            =   315
      TabIndex        =   15
      Top             =   1935
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14420
      SectionData     =   "Cetak_frm_kunjungan.frx":6862
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   360
      TabIndex        =   16
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
      TabIndex        =   10
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
      Picture         =   "Cetak_frm_kunjungan.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17775
      TabIndex        =   13
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
      Picture         =   "Cetak_frm_kunjungan.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17775
      TabIndex        =   11
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
      Picture         =   "Cetak_frm_kunjungan.frx":D33B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17775
      TabIndex        =   12
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
      Picture         =   "Cetak_frm_kunjungan.frx":10981
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1530
      TabIndex        =   17
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
      Picture         =   "Cetak_frm_kunjungan.frx":13E60
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand CmdBR 
      Height          =   420
      Left            =   990
      TabIndex        =   0
      ToolTipText     =   "Simpan"
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
      Picture         =   "Cetak_frm_kunjungan.frx":1A6C2
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label Label10 
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
      Left            =   11025
      TabIndex        =   34
      Top             =   1575
      Width           =   1545
   End
   Begin VB.Label Label9 
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
      Left            =   13410
      TabIndex        =   33
      Top             =   1575
      Width           =   1545
   End
   Begin VB.Label Label8 
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
      Left            =   5670
      TabIndex        =   32
      Top             =   1575
      Width           =   1545
   End
   Begin VB.Label Label6 
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
      Left            =   8055
      TabIndex        =   31
      Top             =   1575
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
      Left            =   2700
      TabIndex        =   30
      Top             =   1575
      Width           =   1545
   End
   Begin VB.Label Label1 
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
      Left            =   315
      TabIndex        =   29
      Top             =   1575
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
      Left            =   15075
      TabIndex        =   28
      Top             =   1080
      Width           =   600
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
      Left            =   12780
      TabIndex        =   27
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label lblnmrute 
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
      Left            =   1485
      TabIndex        =   26
      Top             =   1035
      Width           =   1365
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
      Left            =   10665
      TabIndex        =   25
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ROUTE :"
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
      TabIndex        =   24
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CHEKER :"
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
      Left            =   3060
      TabIndex        =   23
      Top             =   1080
      Width           =   735
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
      Left            =   3780
      TabIndex        =   22
      Top             =   1035
      Width           =   870
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
      Left            =   4680
      TabIndex        =   21
      Top             =   1035
      Width           =   2940
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10485
      TabIndex        =   20
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Cetak_frm_kunjungan 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Kunjungan Cheker"
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
      Left            =   1215
      TabIndex        =   19
      Top             =   135
      Width           =   7665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL PLANNING :"
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
      Left            =   7740
      TabIndex        =   18
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_frm_kunjungan.frx":1CEF4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_frm_kunjungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim sqlNR As String
Dim color As Long, flag As Byte
Dim rsNR As ADODB.Recordset
Dim kata As String

Private Sub cek_NR()
    sqlNRX = "select kdcustomer from route_plan where nmrute='" & lblnmrute & "' and kdteknisi ='" & lblkdteknisi & "' union all" & vbCrLf & _
             "select kdcustomer from real_cek where nmrute='" & lblnmrute & "' and kdteknisi ='" & lblkdteknisi & "' and kdcustomer not in (select kdcustomer from Route_plan  where nmrute='" & txtperiode & "' and kdteknisi ='" & lblkdteknisi & "')"
    
    sqlNR1 = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
                "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
                "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
                "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                    "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= getdate() group by kdcustomer,kdkategori" & vbCrLf & _
                ") a group by kdcustomer " & vbCrLf & _
           ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2+RG <>0"
    
    
    sqlNR2 = "select d.nmareaC,e.nmteknisi,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,a.disp,a.showC,a.RG,b.kdsp + '/' + b.kdcustomer_iap as kdcust_iap from (" & sqlNR1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
             "left join  area_cheker d on b.kdareaC=d.kdareaC left join teknisi e on b.kdteknisi= e.kdteknisi where b.kdteknisi='" & lblkdteknisi & "'"
          
    sqlNR3 = "select kdcustomer,tglplan,tglinput from plan_non_route where nmrute='" & lblnmrute & "' and kdteknisi='" & lblkdteknisi & "'"
          
    sqlNR = "select b.tglplan,a.nmareaC,a.nmteknisi,a.kdcustomer,a.nmcustomer,a.alamat,a.cp,a.telp,a.disp,a.showC,a.RG,a.disp+a.showC+a.RG as total,b.tglinput,a.kdcust_iap from (" & sqlNR2 & ") a left join (" & sqlNR3 & ") b on a.kdcustomer=b.kdcustomer where a.kdcustomer not in (" & sqlNRX & ") and b.tglplan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "'"
    
  
    
End Sub





Private Sub CMbkolom_Click()
If CMbkolom.ListIndex = 0 Then
kata = "convert(int,kode) < 3"
ElseIf CMbkolom.ListIndex = 1 Then
kata = "convert(int,kode) < 4"
ElseIf CMbkolom.ListIndex = 2 Then
kata = "convert(int,kode) < 5"
ElseIf CMbkolom.ListIndex = 3 Then
kata = "convert(int,kode) < 6"
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


Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub




Private Sub Cetak()
On Error GoTo hell


Unload AR_FRM_KUNJUNGAN2

'non route
Call cek_NR

'planing
sqlQ = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC ,RG from ( " & vbCrLf & _
            "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
            "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
            "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl <= getdate() group by kdcustomer,kdkategori" & vbCrLf & _
            ") a group by kdcustomer " & vbCrLf & _
       ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2 + RG <>0"


'realisasi
sqlR1 = "select idrute,kdteknisi,nmrute,kdcustomer,keterangan,det_keterangan,min(tglcek) as tglcek from Real_Cek group by idrute,kdteknisi,nmrute,kdcustomer,keterangan,det_keterangan"

sqlR = "select a.*,isnull(b.tglcek,'1900/01/01') as tglcek,isnull(b.keterangan,'') as keterangan,isnull(b.det_keterangan,'') as det_keterangan from V_real_cek a left join (" & sqlR1 & ") b on a.idrute=b.idrute  where a.kdteknisi='" & lblkdteknisi & "' and a.nmrute='" & lblnmrute & "'"

sql1 = "select '1' as kode,a.idrute,a.tglplan,d.tglcek,a.tglinput,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,b.cp,b.telp,a.jmlunit,e.disp,e.showC,e.RG,isnull(d.disp1,0) as disp1,isnull(d.showC1,0) as showC1,isnull(d.RG,0) as RG1,a.keterangan,a.det_keterangan,b.kdSP + '/' + b.kdcustomer_IAP as kdcust_IAP from ROUTE_PLAN a left join Customer b " & vbCrLf & _
       "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join (" & sqlR & ") d on a.idrute=d.idrute and a.kdcustomer=d.kdcustomer left join (" & sqlQ & ") e on a.kdcustomer=e.kdcustomer where (a.kdteknisi='" & lblkdteknisi & "' and  a.nmrute= '" & lblnmrute & "')  "
   
sql2 = "select kode,kdcustomer,nmcustomer,alamat,cp + ' (' + telp + ')' as telp,disp,showC,rg,tglplan,tglinput,keterangan + '  ' + det_keterangan as keterangan,kdcust_iap from (" & sql1 & " ) a  where disp + showC + RG <> 0 and disp1 + showC1 + RG1 = 0 and a.tglplan between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' union all" & vbCrLf & _
      "select '3' as kode,'','','','',0,0,0,getdate(),getdate(),'','' union all " & vbCrLf & _
      "select '4' as kode,'','','','',0,0,0,getdate(),getdate(),'','' union all " & vbCrLf & _
      "select '5' as kode,'','','','',0,0,0,getdate(),getdate(),'','' Union all " & vbCrLf & _
      "select '2' as kode,kdcustomer,nmcustomer,alamat,cp + ' (' + telp + ')' as telp,disp,showC,rg,tglplan,tglinput,'NON ROUTE' as keterangan,kdcust_iap from (" & sqlNR & ") x"
      
sql3 = "select kdcustomer_iap, SUM(qty_gln1) as qty_gln1,SUM(qty_gln2) as qty_gln2,SUM(qty_gln3) as qty_gln3,SUM(qty_sps1) as qty_sps1,SUM(qty_sps2) as qty_sps2,SUM(qty_sps3) as qty_sps3 from ( " & vbCrLf & _
            "select kdcustomer_IAP,qty_gln as qty_gln1,qty_cup + qty_btl as qty_SPS1,0 as qty_gln2,0 as qty_sps2,0 as qty_gln3,0 as qty_sps3 from omset_iap..omset where bulan=" & cmbbln.Text & " and tahun=" & cmbtahun.Text & " Union all" & vbCrLf & _
            "select kdcustomer_IAP,0 as qty_gln1,0 as qty_sps1,qty_gln as qty_gln2,qty_cup + qty_btl as qty_SPS2,0 as qty_gln3,0 as qty_sps3 from omset_iap..omset where bulan=" & cmbbln2.Text & " and tahun=" & cmbtahun2.Text & " Union all" & vbCrLf & _
            "select kdcustomer_IAP,0 as qty_gln1,0 as qty_sps1,0 as qty_gln2,0 as qty_sps2,qty_gln as qty_gln3,qty_cup + qty_btl as qty_SPS3 from omset_iap..omset where bulan=" & cmbbln3.Text & " and tahun=" & cmbtahun3.Text & " " & vbCrLf & _
       ") x group by kdcustomer_IAP"

sql4 = "select kdcustomer,jmlpiutang,'X' as ket_sewa from piutangsewa where bln=" & cmbbln3.Text & " and tahun=" & cmbtahun3.Text & " "

sql = "select a.*,isnull(b.qty_gln1,0) as qty_GLN1,isnull(b.qty_SPS1,0) qty_SPS1,isnull(b.qty_gln2,0) as qty_GLN2,isnull(b.qty_SPS2,0) qty_SPS2,isnull(b.qty_gln3,0) as qty_GLN3,isnull(b.qty_SPS3,0) qty_SPS3,isnull(c.ket_sewa,'') as ket_sewa from (" & sql2 & ") a left join (" & sql3 & ") b on a.kdcust_iap = b.kdcustomer_iap left join (" & sql4 & ") c on a.kdcustomer=c.kdcustomer where " & kata & " order by a.kode,a.tglplan,a.tglinput"



With AR_FRM_KUNJUNGAN2.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_FRM_KUNJUNGAN2
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldtelp.DataField = "telp"
.flddisp.DataField = "disp"
.fldSH.DataField = "showC"
.fldRG.DataField = "rg"
.fldketerangan.DataField = "keterangan"
.fldgln1.DataField = "qty_gln1"
.fldgln2.DataField = "qty_gln2"
.fldgln3.DataField = "qty_gln3"
.fldsps1.DataField = "qty_sps1"
.fldsps2.DataField = "qty_sps2"
.fldsps3.DataField = "qty_sps3"
.fldket_sewa.DataField = "ket_sewa"


.lblcetak = Format(Now, "dd/MM/yyyy  HH:mm:ss")
.lblnmteknisi = "( " & lblnmteknisi & " )"
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2

.lblbln1 = cmbbln.Text
.lblbln2 = cmbbln2.Text
.lblbln3 = cmbbln3.Text
.lblthn1 = cmbtahun.Text
.lblthn2 = cmbtahun2.Text
.lblthn3 = cmbtahun3.Text
'.fldnmcustomer = lblnmcustomer
'.fldalamat = lblalamat

'
'
Set Me.ARV1.ReportSource = AR_FRM_KUNJUNGAN2
End With

'
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub







Private Sub cmdfs_Click()
AR_FRM_KUNJUNGAN2.Show vbModal
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
Call Cetak

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

CMbkolom.AddItem "0"
CMbkolom.AddItem "1"
CMbkolom.AddItem "2"
CMbkolom.AddItem "3"
CMbkolom.ListIndex = 1

cmbtahun.AddItem Year(Date) - 3
cmbtahun.AddItem Year(Date) - 2
cmbtahun.AddItem Year(Date) - 1
cmbtahun.AddItem Year(Date)
cmbtahun.AddItem Year(Date) + 1
cmbtahun.AddItem Year(Date) + 2
cmbtahun.AddItem Year(Date) + 3


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


cmbtahun2.AddItem Year(Date) - 3
cmbtahun2.AddItem Year(Date) - 2
cmbtahun2.AddItem Year(Date) - 1
cmbtahun2.AddItem Year(Date)
cmbtahun2.AddItem Year(Date) + 1
cmbtahun2.AddItem Year(Date) + 2
cmbtahun2.AddItem Year(Date) + 3
cmbtahun2.ListIndex = 3

cmbbln2.AddItem "1"
cmbbln2.AddItem "2"
cmbbln2.AddItem "3"
cmbbln2.AddItem "4"
cmbbln2.AddItem "5"
cmbbln2.AddItem "6"
cmbbln2.AddItem "7"
cmbbln2.AddItem "8"
cmbbln2.AddItem "9"
cmbbln2.AddItem "10"
cmbbln2.AddItem "11"
cmbbln2.AddItem "12"


cmbtahun3.AddItem Year(Date) - 3
cmbtahun3.AddItem Year(Date) - 2
cmbtahun3.AddItem Year(Date) - 1
cmbtahun3.AddItem Year(Date)
cmbtahun3.AddItem Year(Date) + 1
cmbtahun3.AddItem Year(Date) + 2
cmbtahun3.AddItem Year(Date) + 3


cmbbln3.AddItem "1"
cmbbln3.AddItem "2"
cmbbln3.AddItem "3"
cmbbln3.AddItem "4"
cmbbln3.AddItem "5"
cmbbln3.AddItem "6"
cmbbln3.AddItem "7"
cmbbln3.AddItem "8"
cmbbln3.AddItem "9"
cmbbln3.AddItem "10"
cmbbln3.AddItem "11"
cmbbln3.AddItem "12"

If Month(Date) > 1 Then
cmbbln3.ListIndex = CLng(Month(Date)) - 2
cmbtahun3.ListIndex = 3
ElseIf Month(Date) = 1 Then
cmbbln3.ListIndex = 11
cmbtahun3.ListIndex = 3
End If



If Month(Date) > 2 Then
cmbbln2.ListIndex = CLng(Month(Date)) - 3
cmbtahun2.ListIndex = 3
ElseIf Month(Date) = 2 Then
cmbbln2.ListIndex = 11
cmbtahun2.ListIndex = 2
ElseIf Month(Date) = 1 Then
cmbbln2.ListIndex = 10
cmbtahun2.ListIndex = 2
End If


If Month(Date) > 3 Then
cmbbln.ListIndex = CLng(Month(Date)) - 4
cmbtahun.ListIndex = 3
ElseIf Month(Date) = 3 Then
cmbbln.ListIndex = 11
cmbtahun.ListIndex = 2
ElseIf Month(Date) = 2 Then
cmbbln.ListIndex = 10
cmbtahun.ListIndex = 2
ElseIf Month(Date) = 1 Then
cmbbln.ListIndex = 9
cmbtahun.ListIndex = 2
End If



txttgl1 = Date
txttgl2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub


Private Sub cmdBR_Click()
fixrute_Br1.lblfrm = "FRM_KUNJUNGAN"
fixrute_Br1.Show vbModal
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









