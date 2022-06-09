VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form PObeli_d 
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnoEASAP 
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
      Left            =   8235
      TabIndex        =   3
      Top             =   1530
      Width           =   1905
   End
   Begin VB.Timer TimerNO 
      Left            =   1755
      Top             =   720
   End
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
   End
   Begin VB.TextBox txttglPO 
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
      Left            =   4050
      TabIndex        =   0
      Top             =   1170
      Width           =   1590
   End
   Begin VB.TextBox txtketerangan 
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
      TabIndex        =   2
      Top             =   1530
      Width           =   5730
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   11
      Top             =   720
      Width           =   12255
      _Version        =   524288
      _ExtentX        =   21616
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   12015
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   1125
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
      Picture         =   "PObeli_d.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   915
      Left            =   12825
      TabIndex        =   4
      ToolTipText     =   "Simpan"
      Top             =   1215
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "PObeli_d.frx":2832
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5235
      Left            =   180
      TabIndex        =   5
      Top             =   2610
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   9234
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   14
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   180
      TabIndex        =   12
      Top             =   2205
      Width           =   12300
      _Version        =   524288
      _ExtentX        =   21696
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   0
      Left            =   12870
      TabIndex        =   6
      ToolTipText     =   "Tambah"
      Top             =   2655
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16744576
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PObeli_d.frx":529F
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   12870
      TabIndex        =   7
      ToolTipText     =   "Ubah"
      Top             =   3600
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PObeli_d.frx":7F13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   12870
      TabIndex        =   8
      ToolTipText     =   "Hapus"
      Top             =   4545
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PObeli_d.frx":B110
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   12870
      TabIndex        =   9
      ToolTipText     =   "Refresh"
      Top             =   5490
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PObeli_d.frx":E1A9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   12870
      TabIndex        =   10
      ToolTipText     =   "Cetak"
      Top             =   6435
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PObeli_d.frx":11325
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   13
      Top             =   8280
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
      Picture         =   "PObeli_d.frx":14D82
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NO. EASAP"
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
      Left            =   7335
      TabIndex        =   24
      Top             =   1575
      Width           =   1320
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   6300
      TabIndex        =   23
      Top             =   8820
      Width           =   1095
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   3780
      TabIndex        =   22
      Top             =   8955
      Width           =   1545
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE :"
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
      Left            =   270
      TabIndex        =   21
      Top             =   1215
      Width           =   645
   End
   Begin VB.Label txtkdPO 
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
      Left            =   990
      TabIndex        =   20
      Top             =   1170
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL PO :"
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
      Left            =   3330
      TabIndex        =   19
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Pembelian Barang"
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
      Index           =   1
      Left            =   990
      TabIndex        =   18
      Top             =   45
      Width           =   6000
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
      Left            =   6795
      TabIndex        =   17
      Top             =   1170
      Width           =   1140
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
      Left            =   5940
      TabIndex        =   16
      Top             =   1215
      Width           =   825
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
      Left            =   7965
      TabIndex        =   15
      Top             =   1170
      Width           =   4065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN :"
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
      Left            =   270
      TabIndex        =   14
      Top             =   1575
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "PObeli_d.frx":1B5E4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14100
   End
End
Attribute VB_Name = "PObeli_d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK, rsT As ADODB.Recordset
Dim a As Integer
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim sqlA As String
Dim color As Long, flag As Byte
Dim rscek As ADODB.Recordset



Private Sub cek_dalem()
sqlcek = "select * from PObeli_D where kdPObeli='" & txtkdPO & "'"
Set rscek = con.Execute(sqlcek)
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


Private Sub Cetak()

Unload AR_PObeli

sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan,b.kdkategori from pobeli_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdpobeli='" & txtkdPO & "' order by a.kdbarang"

Set rsX = con.Execute(sqlX)

With AR_PObeli.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_PObeli
.fldunit.DataField = "unit"
.fldnmbarang.DataField = "nmbarang"
.fldsatuan.DataField = "satuan"
.fldketerangan.DataField = "keterangan"

If CLng(rsX!kdkategori) >= 4 Then
.fldkdbarang.DataField = "kdbarang"
Else
.fldkdbarang = ""
End If

.lblnoPO = txtkdPO
.lblnmgudang = lblnmgudang
.lbltglPO = Format(txttglPO, "dd/MM/yyyy")
.lblnoEASAP = UCase(txtnoEASAP)

If txtketerangan = "" Then
.lblNB = ""
Else
.lblNB = "NB : " & txtketerangan
End If

.lbljudul2.Visible = False
.lblkategori.Visible = False


AR_PObeli.Show vbModal

End With

End Sub


Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub

Private Sub tbl()
If rs.RecordCount = 0 Then
    cmdT(1).Enabled = False
    cmdT(2).Enabled = False
    datagrid1.Enabled = False

Else
    cmdT(1).Enabled = True
    cmdT(2).Enabled = True
    datagrid1.Enabled = True
End If
End Sub


Private Sub LG()
On Error GoTo hell

With datagrid1.Columns(0)
.Caption = "KODE"
.Width = 115
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "BARANG"
.Width = 250
End With

With datagrid1.Columns(2)
.Caption = "UNIT"
.Width = 50
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With

With datagrid1.Columns(3)
.Caption = "BELI"
.Width = 50
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With

With datagrid1.Columns(4)
.Caption = "SISA"
.Width = 50
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With


With datagrid1.Columns(5)
.Caption = "SATUAN"
.Width = 70
.Alignment = dbgCenter

End With


With datagrid1.Columns(6)
.Caption = "KETERANGAN"
.Width = 180
End With

With datagrid1.Columns(7)
.Caption = "KDPOBELI_D"
.Width = 0
End With



Call tbl

Exit Sub
hell:
End Sub


Private Sub all()


sqlA = "select a.kdbarang,sum(a.unit) as Ubeli,b.kdpo from beli_d a left join  " & vbCrLf & _
      "beli b  on a.kdbeli=b.kdbeli where b.kdpo='" & txtkdPO & "' group by a.kdbarang,b.kdpo"


sql1 = "select a.kdbarang,b.nmbarang,a.unit,isnull(sum(c.Ubeli),0) as Ubeli,b.satuan,a.keterangan,a.kdpobeli_d from pobeli_d a left join barang b " & vbCrLf & _
      "on a.kdbarang=b.kdbarang left join (" & sqlA & ") c on a.kdPObeli=c.kdPO and a.kdbarang=c.kdbarang where a.kdpobeli='" & txtkdPO & "' group by a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan,a.kdpobeli_d"
      
      
      
sql = "select kdbarang,nmbarang,unit,Ubeli,unit-Ubeli as sisa,satuan,keterangan,kdpobeli_d from (" & sql1 & ") a where unit-Ubeli <>0 order by kdbarang"

      
Set rs = con.Execute(sql)

Set datagrid1.DataSource = rs


Call LG
End Sub



Private Sub tbh()
Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat diUpdate, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else
    If cmdBR.Enabled = False Then
    PObeli_DTU.LBLKODE = 1
          
    PObeli_DTU.Show vbModal
    
    Else
    MsgBox "Kepala data belum disimpan !", vbCritical, "INfo !!"
    End If
End If
End Sub


Private Sub ubh()
Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat diUpdate, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else


    If rs!Ubeli <> 0 Then
        MsgBox "Data Tidak Dapat diubah karena sudah ada Pembelian !", vbCritical, "Error !"
        Exit Sub
    Else
        PObeli_DTU.LBLKODE = 2
        
        
        lblpos = rs.AbsolutePosition
        kode = 2
        
        PObeli_DTU.lblkdPObeli_d = rs!kdPObeli_d
        
        PObeli_DTU.lblkdbarang = rs!kdbarang
        PObeli_DTU.lblnmbarang = rs!nmbarang
        PObeli_DTU.lblsatuan = rs!satuan
        PObeli_DTU.txtunit = FormatNumber(rs!unit, 0)
        PObeli_DTU.lblunit_awal = FormatNumber(rs!unit, 0)
        PObeli_DTU.txtketerangan = rs!keterangan
        PObeli_DTU.cmdBR.Enabled = False
        
          
        PObeli_DTU.Show vbModal
    End If
    
End If
End Sub


Private Sub hps()
On Error GoTo hell
Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Data Tidak dapat diUpdate, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    Exit Sub
Else


    If rs!Ubeli <> 0 Then
        MsgBox "Data Tidak Dapat dihapus karena sudah ada Pembelian !", vbCritical, "Error !"
        Exit Sub
    Else
    
        kode = 2
        Call max
        
        
        ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
        If ms = vbYes Then
            sql = "delete from PObeli_d where kdpobeli_d ='" & rs!kdPObeli_d & "'"
            con.Execute (sql)
            TimerALL.Interval = 10
            PObeli.TimerALL.Interval = 10
        End If
    
    End If
    
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub









Private Sub nomer()
On Error GoTo hell

If LBLKODE = 1 Then
    sql = "select isnull(max(right(kdpobeli,4)),0) as xx from PObeli where Month(tglPObeli)='" & Month(txttglPO) & "'  and year(tglPObeli)='" & Year(txttglPO) & "' and kdgudang= '" & lblkdgudang & "'"
    Set rs = con.Execute(sql)
    
    a = CCur(rs!xx) + 1
    
    If a > 0 Then
    
        Select Case Len(CStr(a))
                Case 1
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "000" & a
                Case 2
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "00" & a
                Case 3
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "0" & a
                Case 4
                    txtkdPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & a
        End Select
    
    Else
        txtnoPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "0001"
    
    End If

End If

Exit Sub
hell:
txtnoPO = lblkdgudang & "/A/" & Format(txttglPO, "MMyy") & "/" & "0001"
End Sub




Private Sub cmdBR_Click()
Gudang_BR.LBLKODE = "PObeli_D"
Gudang_BR.Show vbModal

End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
Call ubh
ElseIf Index = 2 Then
Call hps
ElseIf Index = 3 Then
Call all
ElseIf Index = 4 Then
Call Cetak
End If

End Sub

Private Sub cmdT_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
 TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
End If
End Sub


Private Sub cmdsimpan_Click()

Call Cek_tglOD
If CDate(txttglPO) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Hanya Meng-Update No EASAP saja ya Gaes, Data Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Info !"
   
    
    sql = "Update PObeli set noEASAP='" & UCase(txtnoEASAP) & "' where kdpobeli='" & txtkdPO & "'"
    con.Execute (sql)
    
    sql = "Update beli set noEASAP='" & UCase(txtnoEASAP) & "' where kdPO='" & txtkdPO & "'"
    con.Execute (sql)
    
    cmdsimpan.Enabled = False
    txtketerangan.Enabled = False
    txtnoEASAP.Enabled = False
    
    PObeli.TimerALL.Interval = 10
    
    Exit Sub
Else

    If txtkdPO = "" Or lblkdgudang = "" Then
    MsgBox "Data Belum Lengkap !", vbCritical, "Error !"
    Exit Sub
    Else
    
        If LBLKODE = 1 Then
        Call nomer
        
        sql = "insert into PObeli values ('" & txtkdPO & "','" & Format(txttglPO, "yyyy-MM-dd") & "','" & lblkdgudang & "','" & UCase(txtketerangan) & "','" & UCase(txtnoEASAP) & "')"
        con.Execute (sql)
        
        txttglPO.Enabled = False
        cmdBR.Enabled = False
        txtketerangan.Enabled = False
        txtnoEASAP.Enabled = False
        cmdsimpan.Enabled = False
        cmdT(0).SetFocus
        
        
        
        ElseIf LBLKODE = 2 Then
        sql = "Update PObeli set keterangan='" & UCase(txtketerangan) & "',noEASAP='" & UCase(txtnoEASAP) & "' where kdpobeli='" & txtkdPO & "'"
        con.Execute (sql)
        
        sql = "Update beli set noEASAP='" & UCase(txtnoEASAP) & "' where kdPO='" & txtkdPO & "'"
        con.Execute (sql)
        
        
        txttglPO.Enabled = False
        cmdBR.Enabled = False
        txtketerangan.Enabled = False
        txtnoEASAP.Enabled = False
        cmdsimpan.Enabled = False
        cmdT(0).SetFocus
    
        
        MsgBox "Header PO berhasil di Ubah ", vbInformation, "Info !"
        End If
     
    End If
     
        PObeli.TimerALL.Interval = 10
        PObeli_d.TimerALL.Interval = 10


End If
End Sub




Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub datagrid1_DblClick()
Call ubh
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyLeft Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyRight Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
 Call Cetak
 
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0



txttglPO = Date
txttglPO.Enabled = True




TimerALL.Interval = 10
TimerNO.Interval = 10


Call nul(lblkdgudang)

Call nul(lblnmgudang)



End Sub


Private Sub Form_Unload(Cancel As Integer)
Call cek_dalem

If txttglPO.Enabled = False And rscek.RecordCount = 0 Then
 ms = MsgBox("Tidak Ada Detail PO, apa anda ingin membatalkan Header PO ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = " delete from PObeli where kdPObeli='" & txtkdPO & "' "
        con.Execute (sql)

        PObeli.TimerALL.Interval = 10

        Unload Me

    Else
        Cancel = 1
    End If
End If

End Sub

Private Sub lblkdgudang_Change()
Call nul(lblkdgudang)
Call nomer
End Sub

Private Sub lblnmgudang_Change()
Call nul(lblnmgudang)
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If rs.RecordCount = 0 Then
cmdT(1).SetFocus
Else
datagrid1.SetFocus
End If

If kode = 2 Then
rs.AbsolutePosition = lblpos
End If

 

TimerALL.Interval = 0

End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If


TimerNO.Interval = 0
End Sub



Private Sub txtketerangan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtketerangan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtketerangan_LostFocus()
txtketerangan = UCase(txtketerangan)
End Sub

Private Sub txtnoeasap_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoeasap_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnoeasap_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtnoeasap_LostFocus()
txtnoEASAP = UCase(txtnoEASAP)
End Sub

Private Sub txttglPO_Change()
Call nul(txttglPO)
Call nomer

End Sub

Private Sub txttglPO_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglPO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txttglPO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttglPO_LostFocus()
On Error GoTo hell

txttglPO = FormatDateTime(txttglPO, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglPO.SetFocus

End Sub


