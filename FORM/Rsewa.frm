VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Rsewa 
   BorderStyle     =   0  'None
   ClientHeight    =   10170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   19530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtR 
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
      Left            =   16740
      TabIndex        =   8
      Text            =   "20"
      Top             =   315
      Width           =   735
   End
   Begin VB.CheckBox ChkR 
      BackColor       =   &H00000000&
      Caption         =   "TAMPILKAN :"
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
      Left            =   15165
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   315
      Value           =   1  'Checked
      Width           =   1545
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
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   9495
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
      Left            =   3555
      TabIndex        =   6
      Top             =   9495
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
      Left            =   315
      TabIndex        =   9
      Top             =   675
      Width           =   18015
      _Version        =   524288
      _ExtentX        =   31776
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
      Left            =   18405
      TabIndex        =   0
      ToolTipText     =   "Tambah"
      Top             =   1260
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
      Picture         =   "Rsewa.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   18405
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   2205
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
      Picture         =   "Rsewa.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   18405
      TabIndex        =   2
      ToolTipText     =   "Hapus"
      Top             =   3150
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
      Picture         =   "Rsewa.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   18405
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   4095
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
      Picture         =   "Rsewa.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   18405
      TabIndex        =   4
      ToolTipText     =   "Cari Data"
      Top             =   5040
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
      Picture         =   "Rsewa.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7890
      Left            =   225
      TabIndex        =   14
      Top             =   1080
      Width           =   18060
      _cx             =   31856
      _cy             =   13917
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
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Rsewa.frx":EFAC
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
   Begin Threed.SSOption Opt1 
      Height          =   330
      Left            =   360
      TabIndex        =   16
      Top             =   720
      Width           =   1365
      _ExtentX        =   2408
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
      Caption         =   "On Progress"
   End
   Begin Threed.SSOption Opt2 
      Height          =   330
      Left            =   1800
      TabIndex        =   17
      Top             =   720
      Width           =   1635
      _ExtentX        =   2884
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
      Caption         =   "Sudah di Retur"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RECORD"
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
      Left            =   17505
      TabIndex        =   15
      Top             =   360
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   18405
      Picture         =   "Rsewa.frx":F0DF
      Stretch         =   -1  'True
      Top             =   315
      Width           =   330
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6480
      Picture         =   "Rsewa.frx":F49F
      Stretch         =   -1  'True
      Top             =   9450
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
      Left            =   1665
      TabIndex        =   13
      Top             =   9135
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1485
      Top             =   9090
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Retur Sewa"
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
      TabIndex        =   12
      Top             =   0
      Width           =   2985
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   10440
      Picture         =   "Rsewa.frx":1C34F
      Stretch         =   -1  'True
      Top             =   9090
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
      Left            =   9630
      TabIndex        =   11
      Top             =   9585
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   9180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   0
      Picture         =   "Rsewa.frx":22BA1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19455
   End
End
Attribute VB_Name = "Rsewa"
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

Private Sub cek_dalem()
sqlcek = "select * from Rsewa_D where kdRsewa='" & rs!kdRSewa & "'"
Set rscek = con.Execute(sqlcek)
End Sub

Private Sub ChkR_Click()
TimerAll.Interval = 10

If ChkR.Value = 0 Then
txtR.Enabled = False
Else
txtR.Enabled = True
End If

End Sub

Private Sub ChkR_KeyPress(KeyAscii As Integer)
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

With datagrid1.Columns(0)
.Width = 120
.Caption = "KODE"
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "TANGGAL"
.Width = 80
.Alignment = dbgCenter
End With

With datagrid1.Columns(2)
.Caption = "kdgudang"
.Width = 0
.Alignment = dbgCenter
End With

With datagrid1.Columns(3)
.Caption = "MSK GUDANG"
.Width = 180
End With

With datagrid1.Columns(4)
.Caption = "kdcustomer"
.Width = 0
.Alignment = dbgCenter
End With

With datagrid1.Columns(5)
.Caption = "CUSTOMER"
.Width = 200
End With

With datagrid1.Columns(6)
.Caption = "ALAMAT"
.Width = 325
End With


With datagrid1.Columns(7)
.Caption = "KETERANGAN"
.Width = 250
End With



Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()

RSewa_d.lblkode = 1
RSewa_d.TimerCHKrtr.Interval = 10
RSewa_d.Show vbModal
End Sub

Private Sub ubh()

RSewa_d.lblkode = 2
lblpos = rs.AbsolutePosition
kode = 2

RSewa_d.lblkdgudang = rs!kdgudang
RSewa_d.lblnmgudang = rs!nmgudang
RSewa_d.lblkdcustomer = rs!kdcustomer
RSewa_d.lblnmcustomer = rs!nmcustomer
RSewa_d.lblalamat = rs!alamat

RSewa_d.lblKDRsewa = rs!kdRSewa

RSewa_d.txttglpengajuan = rs!tglpengajuan
RSewa_d.CHKRtr.Value = rs!rtr
RSewa_d.TimerCHKrtr.Interval = 10

RSewa_d.txttglRSewa = rs!tglRsewa

If rs!keterangan = "RETUR" Then
RSewa_d.CMBket.ListIndex = 0
ElseIf rs!keterangan = "PENGGANTIAN" Then
RSewa_d.CMBket.ListIndex = 1
Else
RSewa_d.CMBket.ListIndex = 2
End If

RSewa_d.txtketerangan = rs!keterangan

RSewa_d.txttglpengajuan.Enabled = False
RSewa_d.cmdBR.Enabled = False
RSewa_d.CmdBR1.Enabled = False





RSewa_d.Show vbModal
End Sub

Private Sub hps()
On Error GoTo hell
kode = 3
Call max


    Call cek_dalem
    If rscek.RecordCount <> 0 Then
        MsgBox "Data Tidak dapat dihapus, karena Detail retur masih ada", vbCritical, "Error !"
        Exit Sub

    Else
        ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
        If ms = vbYes Then
            sql = "delete from Rsewa where kdRsewa='" & rs!kdRSewa & "' "
            con.Execute (sql)
            
            TimerAll.Interval = 10
        Else
            Exit Sub
        End If
    End If

Exit Sub
hell:
MsgBox err.Description
End Sub


Private Sub all()

If Opt1.Value = False Then
    If ChkR.Value = 0 Then
        If TXTCARI = "" Then
        sql = "select a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
              "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=1 order by year(a.tglRsewa) desc,a.tglRsewa desc,a.kdRsewa"
        Else
            If CMBCARI.ListIndex < 4 Then
            sql = "select a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=1 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglRsewa) desc,a.tglRsewa desc,a.kdRsewa"
        
            Else
            
            sql = "select a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=1 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglRsewa) desc,a.tglRsewa desc,a.kdRsewa"
        
            End If
        End If
    
    Else
    
        If TXTCARI = "" Then
        sql = "select TOP " & CCur(txtR) & " a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
              "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=1 order by year(a.tglRsewa) desc,a.tglRsewa desc,a.kdRsewa"
        Else
            If CMBCARI.ListIndex < 4 Then
            sql = "select TOP " & CCur(txtR) & " a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=1 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglRsewa) desc,a.tglRsewa desc,a.kdRsewa"
        
            Else
            
            sql = "select TOP " & CCur(txtR) & " a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=1 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglRsewa) desc,a.tglRsewa desc,a.kdRsewa"
        
            End If
        End If
    
    End If

Else
    
    If ChkR.Value = 0 Then
        If TXTCARI = "" Then
        sql = "select a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
              "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=0 order by year(a.tglpengajuan) desc,a.tglpengajuan desc,a.kdRsewa"
        Else
            If CMBCARI.ListIndex < 4 Then
            sql = "select a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=0 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglpengajuan) desc,a.tglpengajuan desc,a.kdRsewa"
        
            Else
            
            sql = "select a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=0 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglpengajuan) desc,a.tglpengajuan desc,a.kdRsewa"
        
            End If
        End If
    
    Else
    
        If TXTCARI = "" Then
        sql = "select TOP " & CCur(txtR) & " a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
              "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=0 order by year(a.tglpengajuan) desc,a.tglpengajuan desc,a.kdRsewa"
        Else
            If CMBCARI.ListIndex < 4 Then
            sql = "select TOP " & CCur(txtR) & " a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=0 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglpengajuan) desc,a.tglpengajuan desc,a.kdRsewa"
        
            Else
            
            sql = "select TOP " & CCur(txtR) & " a.kdRsewa,a.tglpengajuan,a.tglRsewa,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.keterangan,a.rtr from Rsewa a left join gudang b on a.kdgudang=b.kdgudang " & vbCrLf & _
                  "left join customer c on a.kdcustomer=c.kdcustomer where a.rtr=0 and " & kategori & " like '%" & TXTCARI & "%' order by year(a.tglpengajuan) desc,a.tglpengajuan desc,a.kdRsewa"
        
            End If
        End If
    
    End If

End If

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

Call LG
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "a.kdRsewa"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "b.nmgudang"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "c.nmcustomer"
ElseIf CMBCARI.ListIndex = 3 Then
kategori = "a.keterangan"
ElseIf CMBCARI.ListIndex = 4 Then
kategori = "a.tglRsewa"
End If

TimerAll.Interval = 10
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
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()

GradientForm Me, 0

Me.Top = Screen.Height / 3

Me.Height = Me.Height - 1170


CMBCARI.AddItem "KODE"
CMBCARI.AddItem "GUDANG"
CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "KETERANGAN"
CMBCARI.AddItem "TANGGAL"
CMBCARI.ListIndex = 0

Opt1.Value = True

TimerAll.Interval = 10
End Sub

Private Sub Opt1_Click(Value As Integer)
TimerAll.Interval = 10
End Sub

Private Sub Opt2_Click(Value As Integer)
TimerAll.Interval = 10
End Sub

Private Sub TimerAll_Timer()

On Error Resume Next
Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerAll.Interval = 0


End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerAll.Interval = 10
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




Private Sub txtR_Change()
Call nul(txtR)
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerAll.Interval = 10
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii <> vbKeyBack Then

    cekTBL = InStr("1234567890.,", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txtR_LostFocus()
On Error GoTo hell

txtR = FormatNumber(txtR, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtR.SetFocus

End Sub

