VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Planning_kirim 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10260
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chKP 
      BackColor       =   &H00000000&
      Caption         =   "Tampilkan Hanya yg Belum Ter Eksekusi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   4140
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   720
      Value           =   1  'Checked
      Width           =   3885
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
      Height          =   285
      Left            =   15570
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   720
      Value           =   1  'Checked
      Width           =   1545
   End
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
      Left            =   17145
      TabIndex        =   14
      Text            =   "50"
      Top             =   720
      Width           =   735
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
      TabIndex        =   7
      Top             =   9675
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
      Left            =   3510
      TabIndex        =   8
      Top             =   9675
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
      TabIndex        =   15
      Top             =   675
      Width           =   18960
      _Version        =   524288
      _ExtentX        =   33443
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   870
      Index           =   0
      Left            =   19440
      TabIndex        =   16
      ToolTipText     =   "Tambah"
      Top             =   7470
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
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
      Picture         =   "Planning_kirim.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   1
      Left            =   19395
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   1125
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   2
      Left            =   19395
      TabIndex        =   2
      ToolTipText     =   "Hapus"
      Top             =   1980
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   3
      Left            =   19395
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   2835
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   4
      Left            =   19395
      TabIndex        =   4
      ToolTipText     =   "Cari Data"
      Top             =   3690
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7890
      Left            =   225
      TabIndex        =   0
      Top             =   1035
      Width           =   19005
      _cx             =   33523
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Planning_kirim.frx":EFAC
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
      Left            =   405
      TabIndex        =   9
      Top             =   720
      Width           =   1770
      _ExtentX        =   3122
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
      Caption         =   "OTS Pengiriman"
   End
   Begin Threed.SSOption Opt2 
      Height          =   330
      Left            =   2250
      TabIndex        =   10
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
      Caption         =   "Planning Kirim"
   End
   Begin MSComCtl2.DTPicker DTPCari 
      Height          =   330
      Left            =   12060
      TabIndex        =   12
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16761024
      CheckBox        =   -1  'True
      CustomFormat    =   "dd / MM / yyyy"
      Format          =   90898433
      CurrentDate     =   43923
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   6
      Left            =   19395
      TabIndex        =   5
      ToolTipText     =   "Cek Omset"
      Top             =   5400
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":F19B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   5
      Left            =   19395
      TabIndex        =   23
      ToolTipText     =   "Cetak"
      Top             =   4545
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":13812
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   7
      Left            =   19395
      TabIndex        =   6
      ToolTipText     =   "Detail Item"
      Top             =   6255
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1455
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
      Picture         =   "Planning_kirim.frx":1726F
      ButtonStyle     =   4
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
      Height          =   195
      Left            =   17910
      TabIndex        =   22
      Top             =   765
      Width           =   1185
   End
   Begin VB.Label lblket_tgl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Tanggal :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   9720
      TabIndex        =   21
      Top             =   720
      Width           =   2355
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   19395
      Picture         =   "Planning_kirim.frx":1BE01
      Stretch         =   -1  'True
      Top             =   405
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6390
      Picture         =   "Planning_kirim.frx":1C1C1
      Stretch         =   -1  'True
      Top             =   9630
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
      Left            =   1620
      TabIndex        =   20
      Top             =   9315
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1395
      Top             =   9270
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planning Kiriman"
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
      Left            =   1035
      TabIndex        =   19
      Top             =   45
      Width           =   4560
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   7785
      Picture         =   "Planning_kirim.frx":29071
      Stretch         =   -1  'True
      Top             =   9315
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
      Left            =   6975
      TabIndex        =   18
      Top             =   9810
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   540
      TabIndex        =   17
      Top             =   9405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   10230
      Left            =   45
      Picture         =   "Planning_kirim.frx":2F8C3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20400
   End
End
Attribute VB_Name = "Planning_kirim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim kode As Integer
Dim rsmax As ADODB.Recordset
Dim i As Integer
Dim rsC As ADODB.Recordset

Dim color As Long, flag As Byte




Private Sub chKP_Click()
TimerALL.Interval = 10
End Sub

Private Sub ChkR_Click()
TimerALL.Interval = 10

If ChkR.Value = 0 Then
txtR.Enabled = False
Else
txtR.Enabled = True
End If

End Sub

Private Sub DTPCari_Change()
TimerALL.Interval = 10
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
   
    cmdT(6).Enabled = False
    cmdT(7).Enabled = False
    datagrid1.Enabled = False
    img1.Visible = True
    lbl1.Visible = True
Else
   
    cmdT(6).Enabled = True
    cmdT(7).Enabled = True
    datagrid1.Enabled = True
    img1.Visible = False
    lbl1.Visible = False
End If

If rs.RecordCount = 0 And Opt2.Value = True Then
cmdT(2).Enabled = False
ElseIf rs.RecordCount <> 0 And Opt2.Value = True Then
cmdT(2).Enabled = True
End If
End Sub


Private Sub LG()
On Error GoTo hell


Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()
End Sub

Private Sub ubh()
On Error Resume Next

If Opt1.Value = True Then
Planning_kirim_TU.LBLKODE = 1
lblpos = rs.AbsolutePosition
kode = 2
Planning_kirim_TU.txttglPK = Date

Else
Planning_kirim_TU.LBLKODE = 2
lblpos = rs.AbsolutePosition
kode = 2
Planning_kirim_TU.txttglPK = rs!tglPK
Planning_kirim_TU.lblkdteknisi = rs!kdteknisi
Planning_kirim_TU.lblnmteknisi = rs!nmteknisi
Planning_kirim_TU.txturaian = rs!uraian
End If

Planning_kirim_TU.lblkdcustomer = rs!kdcustomer
Planning_kirim_TU.lblnmcustomer = rs!nmcustomer
Planning_kirim_TU.lblalamat = rs!alamat
Planning_kirim_TU.lblkdPK = rs!kode
Planning_kirim_TU.lblnmkategori = rs!nmkategori
Planning_kirim_TU.lblketerangan = rs!keterangan
Planning_kirim_TU.lbljmlunit = CInt(rs!qty_disp) + CInt(rs!qty_SH) + CInt(rs!qty_lain)

Planning_kirim_TU.Show vbModal

End Sub

Private Sub hps()
On Error GoTo hell
kode = 3
Call max
    ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = "delete from planning_kirim  where kdPK='" & rs!kode & "' "
        con.Execute (sql)
        
        TimerALL.Interval = 10
    Else
        Exit Sub
    End If


Exit Sub
hell:
MsgBox err.Description
End Sub


Private Sub all()
MousePointer = vbHourglass

If Opt1.Value = True Then

    If TXTCARI = "" Then
    sql1 = "select '' as kdteknisi,'' as nmteknisi, '' as tglPK,*,'' as uraian,'' as ok from V_tanggungan_kirim where kode not in (select kdPK from planning_kirim) "
    Else
    sql1 = "select '' as kdteknisi,'' as nmteknisi, '' as tglPK,*,'' as uraian,'' as ok from V_tanggungan_kirim where kode not in (select kdPK from planning_kirim) and " & kategori & " like '%" & TXTCARI & "%' "
    End If
    
    If ChkR.Value = 0 Then
        If IsNull(DTPCari.Value) Then
        sql = "select * from (" & sql1 & ") x order by tglpengajuan desc"
        Else
        sql = "select * from (" & sql1 & ") x where tglpengajuan ='" & Format(DTPCari, "yyyy/MM/dd") & "' order by tglpengajuan desc"
        End If
    Else
        If IsNull(DTPCari.Value) Then
        sql = "select top " & CLng(txtR) & " * from (" & sql1 & ") x order by tglpengajuan desc"
        Else
        sql = "select top " & CLng(txtR) & " * from (" & sql1 & ") x where tglpengajuan ='" & Format(DTPCari, "yyyy/MM/dd") & "' order by tglpengajuan desc"
        End If
    End If
    
Else
    If chKP.Value = 0 Then
        If TXTCARI = "" Then
        sql1 = "select a.kdteknisi,c.nmteknisi, a.tglPK,a.KDpk AS kode,b.tglpengajuan,b.nmkategori,b.keterangan,b.kdcustomer,b.nmcustomer,b.alamat,b.nmareaC,b.qty_DISP,b.qty_SH,b.qty_lain,a.Uraian,d.ok from planning_kirim a left join V_ALL_PO_RETUR b on a.kdPK=b.kode left join teknisi c on a.kdteknisi=c.kdteknisi left join V_PK_OK d on a.kdPK=d.kode"
        Else
        sql1 = "select a.kdteknisi,c.nmteknisi, a.tglPK,a.KDpk AS kode,b.tglpengajuan,b.nmkategori,b.keterangan,b.kdcustomer,b.nmcustomer,b.alamat,b.nmareaC,b.qty_DISP,b.qty_SH,b.qty_lain,a.Uraian,d.ok from planning_kirim a left join V_ALL_PO_RETUR b on a.kdPK=b.kode left join teknisi c on a.kdteknisi=c.kdteknisi left join V_PK_OK d on a.kdPK=d.kode where " & kategori & " like '%" & TXTCARI & "%' "
        End If
        
        If ChkR.Value = 0 Then
            If IsNull(DTPCari.Value) Then
            sql = "select * from (" & sql1 & ") x order by tglPK desc,nmteknisi"
            Else
            sql = "select * from (" & sql1 & ") x where tglPK ='" & Format(DTPCari, "yyyy/MM/dd") & "' order by tglPK desc,nmteknisi"
            End If
        Else
            If IsNull(DTPCari.Value) Then
            sql = "select top " & CLng(txtR) & " * from (" & sql1 & ") x order by tglPK desc,nmteknisi"
            Else
            sql = "select top " & CLng(txtR) & " * from (" & sql1 & ") x where tglPK ='" & Format(DTPCari, "yyyy/MM/dd") & "' order by tglPK desc,nmteknisi"
            End If
        End If
    Else
        If TXTCARI = "" Then
        sql1 = "select a.kdteknisi,c.nmteknisi, a.tglPK,a.KDpk AS kode,b.tglpengajuan,b.nmkategori,b.keterangan,b.kdcustomer,b.nmcustomer,b.alamat,b.nmareaC,b.qty_DISP,b.qty_SH,b.qty_lain,a.Uraian,d.ok from planning_kirim a left join V_ALL_PO_RETUR b on a.kdPK=b.kode left join teknisi c on a.kdteknisi=c.kdteknisi left join V_PK_OK d on a.kdPK=d.kode where d.ok is null"
        Else
        sql1 = "select a.kdteknisi,c.nmteknisi, a.tglPK,a.KDpk AS kode,b.tglpengajuan,b.nmkategori,b.keterangan,b.kdcustomer,b.nmcustomer,b.alamat,b.nmareaC,b.qty_DISP,b.qty_SH,b.qty_lain,a.Uraian,d.ok from planning_kirim a left join V_ALL_PO_RETUR b on a.kdPK=b.kode left join teknisi c on a.kdteknisi=c.kdteknisi left join V_PK_OK d on a.kdPK=d.kode where " & kategori & " like '%" & TXTCARI & "%' and d.ok is null"
        End If
        
        If ChkR.Value = 0 Then
            If IsNull(DTPCari.Value) Then
            sql = "select * from (" & sql1 & ") x order by tglPK desc,nmteknisi"
            Else
            sql = "select * from (" & sql1 & ") x where tglPK ='" & Format(DTPCari, "yyyy/MM/dd") & "' and ok isnull order by tglPK desc,nmteknisi"
            End If
        Else
            If IsNull(DTPCari.Value) Then
            sql = "select top " & CLng(txtR) & " * from (" & sql1 & ") x order by tglPK desc,nmteknisi"
            Else
            sql = "select top " & CLng(txtR) & " * from (" & sql1 & ") x where tglPK ='" & Format(DTPCari, "yyyy/MM/dd") & "' order by tglPK desc,nmteknisi"
            End If
        End If
     End If
End If



Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

For i = 1 To (datagrid1.Rows - 1)
For j = 1 To (datagrid1.Cols - 1)

datagrid1.TextMatrix(i, 0) = i

If datagrid1.TextMatrix(i, 16) = "x" Then
datagrid1.Cell(flexcpBackColor, i, j) = &HC0FFC0
End If

Next
Next



If Opt1.Value = True Then

datagrid1.ColHidden(2) = True
datagrid1.ColHidden(3) = True
datagrid1.ColHidden(15) = True
Else

datagrid1.ColHidden(2) = False
datagrid1.ColHidden(3) = False
datagrid1.ColHidden(15) = False
End If


Call LG

MousePointer = vbDefault
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "kode"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "nmkategori"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "keterangan"
ElseIf CMBCARI.ListIndex = 3 Then
kategori = "kdcustomer"
ElseIf CMBCARI.ListIndex = 4 Then
kategori = "nmcustomer"
ElseIf CMBCARI.ListIndex = 5 Then
kategori = "alamat"
ElseIf CMBCARI.ListIndex = 6 Then
kategori = "nmareaC"
ElseIf CMBCARI.ListIndex = 7 Then
kategori = "nmteknisi"

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
    If Opt1.Value = True Then
    List_planning_kirim.lbljudul = "OUTSTADING PENGIRIMAN"
    List_planning_kirim.lbltgl = "TGL PENGAJUAN :"
    List_planning_kirim.cmdBR1.Enabled = False
    Else
    List_planning_kirim.lbljudul = "RENCANA PENGIRIMAN"
    List_planning_kirim.lbltgl = "TGL RENCANA KIRIM :"
    List_planning_kirim.cmdBR1.Enabled = True
    End If
    List_planning_kirim.Show vbModal
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
ElseIf Index = 7 Then
Item_PK.lblkdPK = rs!kode
Item_PK.Show vbModal

    
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
ElseIf KeyAscii = Asc("i") Or KeyAscii = Asc("I") Then
 Item_PK.lblkdPK = rs!kode
 Item_PK.Show vbModal

ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") And cmdT(2).Enabled = True Then
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

Me.Height = Me.Height - 1170


CMBCARI.AddItem "KODE"
CMBCARI.AddItem "KATEGORI"
CMBCARI.AddItem "KETERANGAN"
CMBCARI.AddItem "KD CUST"
CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "ALAMAT"
CMBCARI.AddItem "AREA"
CMBCARI.AddItem "SOPIR"

CMBCARI.ListIndex = 0

Opt1.Value = True

DTPCari.Value = Date
DTPCari.Value = Null


TimerALL.Interval = 10
End Sub

Private Sub OPT1_Click(Value As Integer)
TimerALL.Interval = 10

chKP.Visible = False
cmdT(2).Enabled = False
lblket_tgl = "Tgl Pengajuan :"
End Sub

Private Sub Opt2_Click(Value As Integer)
TimerALL.Interval = 10
chKP.Visible = True
chKP.Value = 1
cmdT(2).Enabled = True
lblket_tgl = "Tgl Rencana Krm :"
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
TimerALL.Interval = 10
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


