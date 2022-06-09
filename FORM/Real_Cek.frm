VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Real_Cek 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkclose 
      BackColor       =   &H00000000&
      Caption         =   "Tampilkan Semua"
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
      Left            =   6795
      TabIndex        =   7
      Top             =   225
      Width           =   1950
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
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   8055
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
      Left            =   2880
      TabIndex        =   6
      Top             =   8055
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
      TabIndex        =   8
      Top             =   675
      Width           =   8475
      _Version        =   524288
      _ExtentX        =   14949
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
      Left            =   8910
      TabIndex        =   1
      ToolTipText     =   "Tambah"
      Top             =   1125
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
      Picture         =   "Real_Cek.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   8910
      TabIndex        =   2
      ToolTipText     =   "Ubah"
      Top             =   2070
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
      Picture         =   "Real_Cek.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   8640
      TabIndex        =   9
      ToolTipText     =   "Hapus"
      Top             =   6840
      Visible         =   0   'False
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
      Picture         =   "Real_Cek.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   8910
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   3015
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
      Picture         =   "Real_Cek.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   8910
      TabIndex        =   4
      ToolTipText     =   "Cari Data"
      Top             =   3960
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
      Picture         =   "Real_Cek.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   6675
      Left            =   135
      TabIndex        =   0
      Top             =   765
      Width           =   8610
      _cx             =   15187
      _cy             =   11774
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
      BackColorAlternate=   14737632
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Real_Cek.frx":EFAC
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
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
      WordWrap        =   0   'False
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
   Begin VB.Image Image2 
      Height          =   540
      Left            =   8910
      Picture         =   "Real_Cek.frx":F068
      Stretch         =   -1  'True
      Top             =   180
      Width           =   510
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5760
      Picture         =   "Real_Cek.frx":F428
      Stretch         =   -1  'True
      Top             =   8010
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
      Left            =   990
      TabIndex        =   13
      Top             =   7695
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   765
      Top             =   7650
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Realisasi Route Plan"
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
      Left            =   540
      TabIndex        =   12
      Top             =   0
      Width           =   5685
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   7290
      Picture         =   "Real_Cek.frx":1C2D8
      Stretch         =   -1  'True
      Top             =   7695
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
      Left            =   6480
      TabIndex        =   11
      Top             =   8190
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   810
      TabIndex        =   10
      Top             =   9225
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "Real_Cek.frx":22B2A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9915
   End
End
Attribute VB_Name = "Real_Cek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori, sqlcek As String
Dim KODE As Integer
Dim rsmax As ADODB.Recordset
Dim rscek As ADODB.Recordset
Dim color As Long, flag As Byte

Private Sub chkclose_Click()
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
Real_cek_TU.lblfrm = "REAL_CEK"
Real_cek_TU.Show vbModal
End Sub

Private Sub ubh()
Real_cek_TU.LBLKODE = 2
lblpos = rs.AbsolutePosition
KODE = 2

Real_cek_TU.lblfrm = "REAL_CEK"
Real_cek_TU.txtperiode = rs!nmrute
Real_cek_TU.lblkdteknisi = rs!kdteknisi
Real_cek_TU.lblnmteknisi = rs!nmteknisi



Real_cek_TU.Show vbModal
End Sub

Private Sub hps()
'On Error GoTo hell
'KODE = 3
'Call max
'
'
'    Call cek_dalem
'    If rscek.RecordCount <> 0 Then
'        MsgBox "Data Tidak dapat dihapus, karena Detail PO masih ada", vbCritical, "Error !"
'        Exit Sub
'
'    Else
'        ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
'        If ms = vbYes Then
'            sql = "delete from PO where kdpo='" & rs!kdPO & "' "
'            con.Execute (sql)
'
'            TimerAll.Interval = 10
'        Else
'            Exit Sub
'        End If
'    End If
'
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !!"
End Sub


Private Sub all()

MousePointer = vbHourglass

If txtcari = "" Then
sql1 = "select a.nmrute,a.kdteknisi,b.nmteknisi,isnull(c.data_pertgl,'1900/01/01') as data_pertgl,isnull(d.kdclose,'') as kdclose from real_cek a left join teknisi b on a.kdteknisi=b.kdteknisi left join route_plan c on a.kdteknisi=c.kdteknisi and a.nmrute=c.nmrute left join close_route_plan d on a.nmrute=d.nmrute and a.kdteknisi=d.kdteknisi group by a.nmrute,a.kdteknisi,b.nmteknisi,c.data_pertgl,d.kdclose"
Else
sql1 = "select a.nmrute,a.kdteknisi,b.nmteknisi,isnull(c.data_pertgl,'1900/01/01') as data_pertgl,isnull(d.kdclose,'') as kdclose from real_cek a left join teknisi b on a.kdteknisi=b.kdteknisi left join route_plan c on a.kdteknisi=c.kdteknisi and a.nmrute=c.nmrute left join close_route_plan d on a.nmrute=d.nmrute and a.kdteknisi=d.kdteknisi where " & kategori & " like '%" & txtcari & "%' group by a.nmrute,a.kdteknisi,b.nmteknisi,c.data_pertgl,d.kdclose"
End If

If chkclose.Value = 0 Then
sql = "select nmrute,kdteknisi,nmteknisi,data_pertgl,(case when kdclose='' then 'OPEN' else 'CLOSE' end) as Status from (" & sql1 & ") a  where kdclose='' order by data_pertgl desc,nmteknisi"
Else
sql = "select nmrute,kdteknisi,nmteknisi,data_pertgl,(case when kdclose='' then 'OPEN' else 'CLOSE' end) as Status from (" & sql1 & ") a  order by data_pertgl desc,nmteknisi"
End If


Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

MousePointer = vbDefault
Call LG
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "a.nmrute"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "b.nmteknisi"
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
txtcari = ""
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
txtcari = ""
Call all
ElseIf Index = 4 Then
txtcari = ""
    If txtcari.Enabled = True Then
    Me.Height = Me.Height - 1170

    txtcari.Enabled = False
    CMBCARI.Enabled = False
    Else
    Me.Height = Me.Height + 1170

    txtcari.Enabled = True
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
 txtcari = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
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
txtcari = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
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


CMBCARI.AddItem "RUTE"
CMBCARI.AddItem "TEKNISI"
CMBCARI.ListIndex = 0

If UTAMA.lblstatus = 0 Then
chkclose.Enabled = False
Else
chkclose.Enabled = True
End If



TimerALL.Interval = 10
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If KODE = 2 Or KODE = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerALL.Interval = 0

MousePointer = vbDefault

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








