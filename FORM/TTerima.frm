VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form TTerima 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7065
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   225
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
      Left            =   8640
      TabIndex        =   8
      Text            =   "30"
      Top             =   225
      Width           =   735
   End
   Begin VB.Timer TimerG 
      Left            =   6165
      Top             =   4815
   End
   Begin VB.Timer TimerAll 
      Left            =   5625
      Top             =   4815
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
      TabIndex        =   6
      Top             =   6390
      Width           =   2850
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
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6390
      Width           =   1860
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   9
      Top             =   675
      Width           =   10275
      _Version        =   524288
      _ExtentX        =   18124
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
      Left            =   10890
      TabIndex        =   0
      ToolTipText     =   "Tambah"
      Top             =   1035
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
      Picture         =   "TTerima.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   1
      Left            =   10890
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   1935
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
      Picture         =   "TTerima.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   2
      Left            =   9675
      TabIndex        =   10
      ToolTipText     =   "Hapus"
      Top             =   5850
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
      Picture         =   "TTerima.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   3
      Left            =   10890
      TabIndex        =   2
      ToolTipText     =   "Refresh"
      Top             =   2790
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
      Picture         =   "TTerima.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   4
      Left            =   10890
      TabIndex        =   3
      ToolTipText     =   "Cari Data"
      Top             =   3645
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
      Picture         =   "TTerima.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   5100
      Left            =   90
      TabIndex        =   4
      Top             =   765
      Width           =   10500
      _cx             =   18521
      _cy             =   8996
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
      AllowUserResizing=   3
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
      FormatString    =   $"TTerima.frx":EFAC
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
      Left            =   9405
      TabIndex        =   15
      Top             =   270
      Width           =   1185
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   1440
      TabIndex        =   14
      Top             =   7875
      Width           =   735
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
      TabIndex        =   13
      Top             =   6525
      Width           =   2220
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   6795
      Picture         =   "TTerima.frx":F05E
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanda Terima"
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
      Left            =   630
      TabIndex        =   12
      Top             =   0
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
      TabIndex        =   11
      Top             =   6030
      Width           =   4560
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5400
      Picture         =   "TTerima.frx":158B0
      Stretch         =   -1  'True
      Top             =   6345
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   10800
      Picture         =   "TTerima.frx":22760
      Stretch         =   -1  'True
      Top             =   225
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   7035
      Left            =   0
      Picture         =   "TTerima.frx":22B20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "TTerima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim KODE As Integer
Dim rsmax As ADODB.Recordset

Dim color As Long, flag As Byte

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
.Width = 80
.Caption = "TGL TT"
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "kdcustomer"
.Width = 0
End With

With datagrid1.Columns(2)
.Caption = "CUSTOMER"
.Width = 200
End With

With datagrid1.Columns(3)
.Caption = "ALAMAT"
.Width = 310
End With

With datagrid1.Columns(4)
.Caption = "JML KWT"
.Width = 60
.Alignment = dbgRight
End With


Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()
TTerima_D.LBLKODE = 1
TTerima_D.Show vbModal
End Sub

Private Sub ubh()
TTerima_D.LBLKODE = 2
lblpos = rs.AbsolutePosition
KODE = 2

TTerima_D.txttglTT = rs!tglTT
TTerima_D.lblkdcustomer = rs!kdcustomer
TTerima_D.lblnmcustomer = rs!nmcustomer
TTerima_D.lblalamat = rs!alamat

TTerima_D.Show vbModal

End Sub

Private Sub hps()
'On Error GoTo hell
'kode = 3
'Call max
'    ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
'    If ms = vbYes Then
'        sql = "delete from Teknisi  where kdteknisi='" & rs!kdteknisi & "' "
'        con.Execute (sql)
'
'        TimerAll.Interval = 10
'    Else
'        Exit Sub
'    End If
'
'
'Exit Sub
'hell:
'MsgBox err.Description
End Sub


Private Sub all()

MousePointer = vbHourglass

If ChkR.Value = 1 Then


    If TXTCARI = "" Then
    sql = "select TOP " & CLng(txtR) & " a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat ,count(a.kdpiutang) as jmlkw  from tanda_terima a left join piutangsewa b on a.kdpiutang=b.kdpiutang" & vbCrLf & _
        "left join customer c on b.kdcustomer=c.kdcustomer group by a.tglTT ,b.kdcustomer,c.nmcustomer,c.alamat order by a.tglTT DESC,c.nmcustomer"
    Else
        If CMBCARI.ListIndex < 2 Then
        sql = "select TOP " & CLng(txtR) & " a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat ,count(a.kdpiutang) as jmlkw  from tanda_terima a left join piutangsewa b on a.kdpiutang=b.kdpiutang" & vbCrLf & _
            "left join customer c on b.kdcustomer=c.kdcustomer where " & kategori & " like '%" & TXTCARI & "%' group by a.tglTT ,b.kdcustomer,c.nmcustomer,c.alamat order by a.tglTT DESC,c.nmcustomer"
        Else
        sql = "select TOP " & CLng(txtR) & " a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat ,count(a.kdpiutang) as jmlkw  from tanda_terima a left join piutangsewa b on a.kdpiutang=b.kdpiutang" & vbCrLf & _
            "left join customer c on b.kdcustomer=c.kdcustomer where a.tglTT = '" & Format(TXTCARI, "yyyy/MM/dd") & "' group by a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat order by a.tglTT DESC,c.nmcustomer"
        End If
    End If
    
Else

    If TXTCARI = "" Then
    sql = "select a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat ,count(a.kdpiutang) as jmlkw  from tanda_terima a left join piutangsewa b on a.kdpiutang=b.kdpiutang" & vbCrLf & _
        "left join customer c on b.kdcustomer=c.kdcustomer group by a.tglTT ,b.kdcustomer,c.nmcustomer,c.alamat order by a.tglTT DESC ,c.nmcustomer"
    Else
        If CMBCARI.ListIndex < 2 Then
        sql = "select a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat ,count(a.kdpiutang) as jmlkw  from tanda_terima a left join piutangsewa b on a.kdpiutang=b.kdpiutang" & vbCrLf & _
            "left join customer c on b.kdcustomer=c.kdcustomer where " & kategori & " like '%" & TXTCARI & "%' group by a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat order by a.tglTT DESC,c.nmcustomer"
        Else
        sql = "select a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat ,count(a.kdpiutang) as jmlkw  from tanda_terima a left join piutangsewa b on a.kdpiutang=b.kdpiutang" & vbCrLf & _
            "left join customer c on b.kdcustomer=c.kdcustomer where a.tglTT = '" & Format(TXTCARI, "yyyy/MM/dd") & "' group by a.tglTT,b.kdcustomer,c.nmcustomer,c.alamat order by a.tglTT DESC,c.nmcustomer"
        End If
    End If
    
End If

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

MousePointer = vbDefault

Call LG
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "c.nmcustomer"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "c.alamat"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "a.tglTT"
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
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub


Private Sub Form_Load()

GradientForm Me, 0

Me.Height = Me.Height - 1170


CMBCARI.AddItem "CUSTOMER"
CMBCARI.AddItem "ALAMAT"
CMBCARI.AddItem "TGL TT"

CMBCARI.ListIndex = 0



TimerAll.Interval = 10
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If KODE = 2 Or KODE = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerAll.Interval = 0

MousePointer = vbDefault

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
'ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
' Call tbh
'ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
' If rs.RecordCount <> 0 Then
' Call ubh
' End If
'ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
' If rs.RecordCount <> 0 Then
' Call hps
' End If
'ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
' Call all
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

