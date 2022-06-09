VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form TeknisiLuar 
   BorderStyle     =   0  'None
   ClientHeight    =   10305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10305
   ScaleWidth      =   19800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Opt3 
      BackColor       =   &H00000000&
      Caption         =   "SELESAI"
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
      Left            =   3690
      TabIndex        =   12
      Top             =   720
      Width           =   1185
   End
   Begin VB.OptionButton OPT1 
      BackColor       =   &H00000000&
      Caption         =   "NOT PLANNING"
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
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   1680
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00000000&
      Caption         =   "ON PLANNING"
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
      Left            =   2025
      TabIndex        =   11
      Top             =   720
      Width           =   1680
   End
   Begin VB.ComboBox CMBCARI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   9540
      Width           =   1860
   End
   Begin VB.TextBox TXTCARI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   9
      Top             =   9540
      Width           =   2850
   End
   Begin VB.Timer TimerAll 
      Left            =   5625
      Top             =   4815
   End
   Begin VB.Timer TimerG 
      Left            =   6165
      Top             =   4815
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   13
      Top             =   675
      Width           =   18330
      _Version        =   524288
      _ExtentX        =   32332
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
      Left            =   18720
      TabIndex        =   1
      ToolTipText     =   "Tambah"
      Top             =   1350
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
      Picture         =   "TeknisiLuar.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   18720
      TabIndex        =   2
      ToolTipText     =   "Ubah"
      Top             =   2295
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
      Picture         =   "TeknisiLuar.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   18720
      TabIndex        =   3
      ToolTipText     =   "Hapus"
      Top             =   3240
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
      Picture         =   "TeknisiLuar.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   18720
      TabIndex        =   4
      ToolTipText     =   "Refresh"
      Top             =   4185
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
      Picture         =   "TeknisiLuar.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   18720
      TabIndex        =   5
      ToolTipText     =   "Cari Data"
      Top             =   5130
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
      Picture         =   "TeknisiLuar.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7800
      Left            =   270
      TabIndex        =   0
      Top             =   1080
      Width           =   18285
      _cx             =   32253
      _cy             =   13758
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648384
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"TeknisiLuar.frx":EFAC
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   4
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   5
      Left            =   18720
      TabIndex        =   6
      ToolTipText     =   "Cetak"
      Top             =   6075
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
      Picture         =   "TeknisiLuar.frx":F21F
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   6
      Left            =   18720
      TabIndex        =   7
      ToolTipText     =   "Cek Omset"
      Top             =   7020
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
      Picture         =   "TeknisiLuar.frx":12C7C
      ButtonStyle     =   4
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   18720
      Picture         =   "TeknisiLuar.frx":172F3
      Stretch         =   -1  'True
      Top             =   405
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6300
      Picture         =   "TeknisiLuar.frx":176B3
      Stretch         =   -1  'True
      Top             =   9540
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Pencarian"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1530
      TabIndex        =   17
      Top             =   9180
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1305
      Top             =   9135
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Teknisi Luar"
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
      Top             =   0
      Width           =   5685
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   10035
      Picture         =   "TeknisiLuar.frx":24563
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   555
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA TIDAK ADA"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9225
      TabIndex        =   15
      Top             =   9765
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   765
      TabIndex        =   14
      Top             =   9855
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   10185
      Left            =   0
      Picture         =   "TeknisiLuar.frx":2ADB5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19770
   End
End
Attribute VB_Name = "TeknisiLuar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori, sqlcek As String
Dim kode As Integer
Dim rsmax As ADODB.Recordset
Dim rscek As ADODB.Recordset
Dim color As Long, flag As Byte
Dim sqlN As String
Dim rsC As ADODB.Recordset
Dim rsN As ADODB.Recordset

Private Sub cek_dalam()
sqlcek = "select * from TeknisiLuar_D where kdTD='" & rs!kdTD & "'"
Set rscek = con.Execute(sqlcek)
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub



'untuk set cursor pada saat dihapus
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
    img1.Visible = True
    lbl1.Visible = True
Else
    cmdT(1).Enabled = True
    cmdT(2).Enabled = True
    datagrid1.Enabled = True
    img1.Visible = False
    lbl1.Visible = False
End If
End Sub


Private Sub LG()
On Error GoTo hell
Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()
TeknisiLuar_D.LBLKODE = 1
TeknisiLuar_D.TimerCHKrencana.Interval = 10
TeknisiLuar_D.Show vbModal
End Sub

Private Sub ubh()
TeknisiLuar_D.LBLKODE = 2
lblpos = rs.AbsolutePosition
kode = 2

TeknisiLuar_D.txtkdTL = rs!kdTL

TeknisiLuar_D.txttglkomplain = rs!tglkomplain
TeknisiLuar_D.CHKRencana = rs!rencana


TeknisiLuar_D.TimerCHKrencana.Interval = 10
TeknisiLuar_D.TimerCHKTL.Interval = 10

TeknisiLuar_D.chkTL = rs!TL
TeknisiLuar_D.txttglTL = rs!tglTL

TeknisiLuar_D.txttglrencana = rs!tglrencana

TeknisiLuar_D.lblkdcustomer = rs!kdcustomer
TeknisiLuar_D.lblnmcustomer = rs!nmcustomer
TeknisiLuar_D.lblalamat = rs!alamat

TeknisiLuar_D.txtkerusakan = rs!kerusakan
TeknisiLuar_D.lblkd1 = rs!kd1
TeknisiLuar_D.lblkdbarang = rs!kdbarang
TeknisiLuar_D.lblnmkategori = rs!nmkategori
TeknisiLuar_D.lblkdsap = rs!kdSAP
TeknisiLuar_D.lblkdteknisi = rs!kdteknisi
TeknisiLuar_D.lblnmteknisi = rs!nmteknisi
TeknisiLuar_D.txtPIC_OTL = rs!pic_otl
TeknisiLuar_D.CMbTindakan.Text = rs!tindakan

TeknisiLuar_D.txtjam_datang = Left(rs!jam_datang, 5)
TeknisiLuar_D.txtjam_selesai = Left(rs!jam_selesai, 5)


TeknisiLuar_D.txttglkomplain.Enabled = False
TeknisiLuar_D.Show vbModal
End Sub

Private Sub hps()
On Error GoTo hell

    kode = 3
    Call max
        ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
        If ms = vbYes Then
            sql = "delete from TeknisiLuar_d where kdTL='" & rs!kdTL & "' "
            con.Execute (sql)
            
            sql = "delete from TeknisiLuar where kdTL='" & rs!kdTL & "' "
            con.Execute (sql)
            
            TimerALL.Interval = 10
        Else
            Exit Sub
        End If



Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"
End Sub


Private Sub all()

sql1 = "select a.kdTL,a.kdcustomer,e.nmcustomer,e.alamat,a.tglkomplain,a.rencana,a.tglrencana,a.TL,a.tglTL,a.kdbarang,b.kd1,b.kdsap,c.nmkategori,b.merk,a.kerusakan,a.kdteknisi,d.nmteknisi,a.jam_datang,a.jam_selesai,a.tindakan,a.pic_otl from teknisiluar a left join barang b on a.kdbarang=b.kdbarang left join kategoribrg c on b.kdkategori= c.kdkategori " & vbCrLf & _
      "left join teknisi d on a.kdteknisi=d.kdteknisi left join customer e on a.kdcustomer=e.kdcustomer "

If OPT1.Value = True Then

    If TXTCARI = "" Then
    sql = "select * from (" & sql1 & ") x where rencana=0 order by tglkomplain desc"
          
    Else
        If CMBCARI.ListIndex < 9 Then
        sql = "select * from (" & sql1 & ") x where rencana=0 and  " & kategori & " like '%" & TXTCARI & "%' order by tglkomplain desc"
        Else
        sql = "select * from (" & sql1 & ") x where rencana=0 and  " & kategori & " = '" & Format(TXTCARI, "yyyy/MM/dd") & "' order by tglkomplain desc"
        End If
    End If
    
ElseIf Opt2.Value = True Then

    If TXTCARI = "" Then
    sql = "select * from (" & sql1 & ") x where rencana=1 and TL=0 order by tglrencana desc"
          
    Else
        If CMBCARI.ListIndex < 9 Then
        sql = "select * from (" & sql1 & ") x where rencana=1 and TL=0 and " & kategori & " like '%" & TXTCARI & "%' order by tglrencana desc"
        Else
        sql = "select * from (" & sql1 & ") x where rencana=1 and TL=0 and  " & kategori & " = '" & Format(TXTCARI, "yyyy/MM/dd") & "' order by tglrencana desc"
        End If
    End If
    
ElseIf Opt3.Value = True Then

    If TXTCARI = "" Then
    sql = "select * from (" & sql1 & ") x where TL=1 order by tglTL desc"
          
    Else
        If CMBCARI.ListIndex < 9 Then
        sql = "select * from (" & sql1 & ") x where TL=1 and " & kategori & " like '%" & TXTCARI & "%' order by tglTL desc"
        Else
        sql = "select * from (" & sql1 & ") x where TL=1 and  " & kategori & " = '" & Format(TXTCARI, "yyyy/MM/dd") & "' order by tglTL desc"
        End If
    End If


End If

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

Call LG
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "kdTL"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "KDBARANG"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "nmkategori"
ElseIf CMBCARI.ListIndex = 3 Then
kategori = "kdcustomer"
ElseIf CMBCARI.ListIndex = 4 Then
kategori = "nmcustomer"
ElseIf CMBCARI.ListIndex = 5 Then
kategori = "alamat"
ElseIf CMBCARI.ListIndex = 6 Then
kategori = "nmteknisi"
ElseIf CMBCARI.ListIndex = 7 Then
kategori = "kerusakan"
ElseIf CMBCARI.ListIndex = 8 Then
kategori = "tindakan"
ElseIf CMBCARI.ListIndex = 9 Then
kategori = "tglkomplain"
ElseIf CMBCARI.ListIndex = 10 Then
kategori = "tglrencana"
ElseIf CMBCARI.ListIndex = 11 Then
kategori = "tglselesai"

End If

TimerALL.Interval = 10
End Sub

Private Sub CMBCARI_KeyPress(KeyAscii As Integer)
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
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
     If rs.RecordCount <> 0 Then
     Call ubh
     End If
ElseIf Index = 2 Then
     If rs.RecordCount <> 0 Then
     Call hps
     End If
ElseIf Index = 3 Then
TXTCARI = ""
Call all
ElseIf Index = 4 Then
TXTCARI = ""
    If TXTCARI.Enabled = True Then
    Me.Height = Me.Height - 1170

    TXTCARI.Enabled = False
    CMBCARI.Enabled = False
    Else
    Me.Height = Me.Height + 1170

    TXTCARI.Enabled = True
    CMBCARI.Enabled = True
    End If
ElseIf Index = 5 Then
TeknisiLuar_list.Show vbModal

ElseIf Index = 6 Then
sqlC = "select a.kdcustomer,a.kdsp + '/' + a.kdcustomer_IAP as kdcust_IAP,isnull(b.nmcustomer_iap,'-') as nmcustomer_IAP,isnull(alamat_iap,'-') as alamat_iap,isnull(c.nmsp,'-') as nmsp from customer a left join customer_IAP b " & vbCrLf & _
       "on a.kdsp + '/' + a.kdcustomer_iap = b.pk_cust_IAP left join sp_iap c on a.kdsp=c.kdsp where a.kdcustomer='" & rs!kdcustomer & "'"
Set rsC = con.Execute(sqlC)

LIST_Omset_IAP.lblkdcustomer_IAP = rsC!kdcust_IAP
LIST_Omset_IAP.lblnmcustomer_IAP = rsC!nmcustomer_IAP
LIST_Omset_IAP.lblalamat_IAP = rsC!alamat_IAP
LIST_Omset_IAP.lblnmsp = rsC!nmsp
LIST_Omset_IAP.lblkdcustomer = rs!kdcustomer
LIST_Omset_IAP.Show vbModal

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
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub

Private Sub datagrid1_Click()
TimerG.Interval = 10
End Sub

Private Sub datagrid1_DblClick()
 If rs.RecordCount <> 0 Then
 Call ubh
 End If

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
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()

GradientForm Me, 0

Me.Height = Me.Height - 1170

OPT1.Value = True

CMBCARI.AddItem "KODE"
CMBCARI.AddItem "KD BARANG"
CMBCARI.AddItem "JNS UNIT"
CMBCARI.AddItem "KD CUSTOMER"
CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "ALAMAT"
CMBCARI.AddItem "TEKNISI"
CMBCARI.AddItem "KERUSAKAN"
CMBCARI.AddItem "TINDAKAN"
CMBCARI.AddItem "TGL KOMPLAIN"
CMBCARI.AddItem "TGL RENCANA"
CMBCARI.AddItem "TGL SELESAI"

CMBCARI.ListIndex = 0



TimerALL.Interval = 10
End Sub

Private Sub OPT1_Click()
cmdT(5).Enabled = False
TimerALL.Interval = 10
End Sub

Private Sub Opt2_Click()
cmdT(5).Enabled = True
TimerALL.Interval = 10
End Sub

Private Sub Opt3_Click()
cmdT(5).Enabled = False
TimerALL.Interval = 10
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerALL.Interval = 0

End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerALL.Interval = 10
End Sub

Private Sub TXTCARI_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTCARI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    TimerG.Interval = 10
    Else
    SendKeys vbTab
    End If
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub








