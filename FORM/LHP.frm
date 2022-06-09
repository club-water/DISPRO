VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form LHP 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15555
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   15555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6390
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
      Left            =   2520
      TabIndex        =   8
      Top             =   6390
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
      TabIndex        =   9
      Top             =   675
      Width           =   14145
      _Version        =   524288
      _ExtentX        =   24950
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   0
      Left            =   14625
      TabIndex        =   0
      ToolTipText     =   "Tambah"
      Top             =   990
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "LHP.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   1
      Left            =   14625
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   1800
      Width           =   780
      _ExtentX        =   1376
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "LHP.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   2
      Left            =   4635
      TabIndex        =   10
      ToolTipText     =   "Hapus"
      Top             =   0
      Visible         =   0   'False
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
      Picture         =   "LHP.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   3
      Left            =   14625
      TabIndex        =   2
      ToolTipText     =   "Refresh"
      Top             =   2610
      Width           =   780
      _ExtentX        =   1376
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "LHP.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   4
      Left            =   14625
      TabIndex        =   3
      ToolTipText     =   "Cari Data"
      Top             =   3420
      Width           =   780
      _ExtentX        =   1376
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "LHP.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   4650
      Left            =   135
      TabIndex        =   4
      Top             =   855
      Width           =   14325
      _cx             =   25268
      _cy             =   8202
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
      BackColorAlternate=   16777088
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"LHP.frx":EFAC
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
   Begin Threed.SSOption Oblunas 
      Height          =   330
      Left            =   135
      TabIndex        =   5
      Top             =   5535
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "Belum Clear"
   End
   Begin Threed.SSOption Olunas 
      Height          =   330
      Left            =   1575
      TabIndex        =   6
      Top             =   5535
      Width           =   780
      _ExtentX        =   1376
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
      Caption         =   "Clear"
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   14625
      Picture         =   "LHP.frx":F134
      Stretch         =   -1  'True
      Top             =   270
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5400
      Picture         =   "LHP.frx":F4F4
      Stretch         =   -1  'True
      Top             =   6345
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
      Left            =   630
      TabIndex        =   14
      Top             =   6030
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   405
      Top             =   5985
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LHP "
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
      Left            =   900
      TabIndex        =   13
      Top             =   0
      Width           =   4560
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   6795
      Picture         =   "LHP.frx":1C3A4
      Stretch         =   -1  'True
      Top             =   6030
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
      Left            =   5985
      TabIndex        =   12
      Top             =   6525
      Width           =   2220
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   7875
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   7035
      Left            =   0
      Picture         =   "LHP.frx":22BF6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15450
   End
End
Attribute VB_Name = "LHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim KODE As Integer
Dim rsmax As ADODB.Recordset
Dim kata_clr As String

Dim color As Long, flag As Byte

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
LHP_D.LBLKODE = 1
LHP_D.Show vbModal
End Sub

Private Sub ubh()
On Error Resume Next

LHP_D.LBLKODE = 2
lblpos = rs.AbsolutePosition
KODE = 2

LHP_D.txttglLHP = rs!tglLHP
LHP_D.txttglCLR = rs!tglCLR
LHP_D.lblkdkolektor = rs!kdkolektor
LHP_D.lblnmkolektor = rs!nmkolektor
LHP_D.ChKCLEAR.Value = rs!clr

LHP_D.Show vbModal

End Sub

Private Sub hps()
End Sub


Private Sub all()
MousePointer = vbHourglass

If Oblunas.Value = True Then
kata_clr = "a.clr =0"
Else
kata_clr = "a.clr =1"
End If

If txtcari = "" Then
sql1 = "select a.tgllhp,(CASE WHEN DATENAME(dw, a.tglLHP)='Sunday' then 'MINGGU'" & vbCrLf & _
      "WHEN DATENAME(dw, a.tglLHP)='Monday' THEN 'SENIN' WHEN DATENAME(dw, a.tglLHP)='Tuesday' THEN 'SELASA' WHEN DATENAME(dw, a.tglLHP)='Wednesday' THEN 'RABU'" & vbCrLf & _
      "WHEN DATENAME(dw, a.tglLHP)='Thursday' THEN 'KAMIS' WHEN DATENAME(dw, a.tglLHP)='Friday' THEN 'JUMAT' ELSE 'SABTU' END ) as hari,a.kdkolektor,c.nmkolektor" & vbCrLf & _
      ",SUM(b.jmlpiutang) as jmlLHP,a.clr,N_CLR = case when A.clr=1 then 'X' else '' end,a.tglCLR,a.status  from Lhp a left join PiutangSewa b on a.kdpiutang=b.kdpiutang left join kolektor c on a.kdkolektor=c.kdkolektor  where " & kata_clr & " group by a.tgllhp,a.kdkolektor,c.nmkolektor,a.clr,a.tglCLR,a.status "
Else
    If CMBCARI.ListIndex = 0 Then
    sql1 = "select a.tgllhp,(CASE WHEN DATENAME(dw, a.tglLHP)='Sunday' then 'MINGGU'" & vbCrLf & _
          "WHEN DATENAME(dw, a.tglLHP)='Monday' THEN 'SENIN' WHEN DATENAME(dw, a.tglLHP)='Tuesday' THEN 'SELASA' WHEN DATENAME(dw, a.tglLHP)='Wednesday' THEN 'RABU'" & vbCrLf & _
          "WHEN DATENAME(dw, a.tglLHP)='Thursday' THEN 'KAMIS' WHEN DATENAME(dw, a.tglLHP)='Friday' THEN 'JUMAT' ELSE 'SABTU' END ) as hari,a.kdkolektor,c.nmkolektor" & vbCrLf & _
          ",SUM(b.jmlpiutang) as jmlLHP,a.clr,N_CLR = case when A.clr=1 then 'X' else '' end,a.tglCLR,a.status  from Lhp a left join PiutangSewa b on a.kdpiutang=b.kdpiutang left join kolektor c on a.kdkolektor=c.kdkolektor where " & kata_clr & " and " & kategori & " like '%" & txtcari & "%' group by a.tgllhp,a.kdkolektor,c.nmkolektor,a.clr,a.tglCLR,a.status"

    
    Else
    sql1 = "select a.tgllhp,(CASE WHEN DATENAME(dw, a.tglLHP)='Sunday' then 'MINGGU'" & vbCrLf & _
          "WHEN DATENAME(dw, a.tglLHP)='Monday' THEN 'SENIN' WHEN DATENAME(dw, a.tglLHP)='Tuesday' THEN 'SELASA' WHEN DATENAME(dw, a.tglLHP)='Wednesday' THEN 'RABU'" & vbCrLf & _
          "WHEN DATENAME(dw, a.tglLHP)='Thursday' THEN 'KAMIS' WHEN DATENAME(dw, a.tglLHP)='Friday' THEN 'JUMAT' ELSE 'SABTU' END ) as hari,a.kdkolektor,c.nmkolektor" & vbCrLf & _
          ",SUM(b.jmlpiutang) as jmlLHP,a.clr,N_CLR = case when A.clr=1 then 'X' else '' end,a.tglCLR,a.status  from Lhp a left join PiutangSewa b on a.kdpiutang=b.kdpiutang left join kolektor c on a.kdkolektor=c.kdkolektor where " & kata_clr & " and  a.tgllhp = '" & Format(txtcari, "yyyy/MM/dd") & "' group by a.tgllhp,a.kdkolektor,c.nmkolektor,a.clr,a.tglCLR,a.status"
    
    End If
    
End If

sql = "select tglLHP,hari,kdkolektor,nmkolektor," & vbCrLf & _
      "sum(case status when 'TERTAGIH' then jmlLHP else 0 end) as S1," & vbCrLf & _
      "sum(case status when 'TDK TERTAGIH' then jmlLHP else 0 end) as S2," & vbCrLf & _
      "sum(case status when 'TANDA TERIMA' then jmlLHP else 0 end) as S3," & vbCrLf & _
      "sum(jmlLHP) as jmlLHP,clr,N_clr,tglCLR from (" & sql1 & ") x group by tglLHP,hari,kdkolektor,nmkolektor,clr,N_clr,tglCLR order by tgllhp desc"
      


Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

datagrid1.MergeCells = flexMergeRestrictAll
datagrid1.MergeRow(0) = True


Call LG

MousePointer = vbDefault
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "c.nmkolektor"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "a.tglLHP"
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

Private Sub DataGrid1_DblClick()
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


Private Sub Form_Load()

GradientForm Me, 0

Me.Height = Me.Height - 1170


Oblunas.Value = True

CMBCARI.AddItem "KOLEKTOR"
CMBCARI.AddItem "TGL LHP"


CMBCARI.ListIndex = 0



TimerALL.Interval = 10
End Sub

Private Sub Oblunas_Click(Value As Integer)
cmdT(1).Enabled = True
TimerALL.Interval = 10
End Sub

Private Sub Oblunas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Olunas_Click(Value As Integer)
cmdT(1).Enabled = False
TimerALL.Interval = 10
End Sub

Private Sub Olunas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If KODE = 2 Or KODE = 3 Then
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








