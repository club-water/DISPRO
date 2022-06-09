VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Piutang_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcari 
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
      Left            =   225
      TabIndex        =   1
      Top             =   1305
      Width           =   2490
   End
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   2
      Top             =   855
      Width           =   10410
      _Version        =   524288
      _ExtentX        =   18362
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   6210
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
      Picture         =   "Piutang_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   4110
      Left            =   225
      TabIndex        =   0
      Top             =   1710
      Width           =   10320
      _cx             =   18203
      _cy             =   7250
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Piutang_BR.frx":6862
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
   Begin VB.Label lblpos 
      Height          =   285
      Left            =   4050
      TabIndex        =   8
      Top             =   7380
      Width           =   960
   End
   Begin VB.Label lblkdkategori 
      Caption         =   "lblkategori"
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   7695
      Width           =   1155
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   10665
      Picture         =   "Piutang_BR.frx":6975
      Stretch         =   -1  'True
      Top             =   450
      Width           =   285
   End
   Begin VB.Label lbljudul 
      BackStyle       =   0  'Transparent
      Caption         =   "Piutang Sewa"
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
      Left            =   585
      TabIndex        =   6
      Top             =   135
      Width           =   7755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   270
      TabIndex        =   5
      Top             =   990
      Width           =   1500
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   7695
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   6765
      Left            =   45
      Picture         =   "Piutang_BR.frx":6D35
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11220
   End
End
Attribute VB_Name = "Piutang_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim sql1, sql2, sql As String

Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
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


Private Sub LG()
On Error GoTo hell

With datagrid1.Columns(0)
.Width = 120
.Caption = "NO KWITANSI"
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "BLN"
.Width = 40
.Alignment = dbgCenter
End With

With datagrid1.Columns(2)
.Caption = "TAHUN"
.Width = 60
.Alignment = dbgCenter
End With

With datagrid1.Columns(3)
.Caption = "kdcustomer"
.Width = 0
.Alignment = dbgCenter
End With

With datagrid1.Columns(4)
.Caption = "JML PIUTANG"
.Width = 100
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With

With datagrid1.Columns(5)
.Caption = "JML BAYAR"
.Width = 100
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With

With datagrid1.Columns(6)
.Caption = "POTONGAN"
.Width = 100
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With


With datagrid1.Columns(7)
.Caption = "SISA PIUTANG"
.Width = 100
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With





Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell


If TXTCARI = "" Then
sql1 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
        "sum(jmlpiutang - jmlbayar - potongan) as sisa from (" & vbCrLf & _
        "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar,0 as potongan from piutangsewa" & vbCrLf & _
        "Union" & vbCrLf & _
        "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
        "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"


sql = "select a.kdpiutang,c.bln,c.tahun,a.kdcustomer,a.jmlpiutang,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a " & vbCrLf & _
     "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join Tanda_terima b on a.kdpiutang=b.kdpiutang where a.kdcustomer='" & TTerima_D.lblkdcustomer & "' and a.sisa<>0 and c.tt=0 order by c.tahun,c.bln"

Else

sql1 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
        "sum(jmlpiutang - jmlbayar - potongan) as sisa from (" & vbCrLf & _
        "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar,0 as potongan from piutangsewa" & vbCrLf & _
        "Union" & vbCrLf & _
        "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
        "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"


sql = "select a.kdpiutang,c.bln,c.tahun,a.kdcustomer,a.jmlpiutang,a.jmlbayar,a.potongan,a.sisa from (" & sql1 & ") a " & vbCrLf & _
      "left join piutangsewa c on a.kdpiutang=c.kdpiutang left join Tanda_terima b on a.kdpiutang=b.kdpiutang where a.kdcustomer='" & TTerima_D.lblkdcustomer & "' and a.sisa<>0 and c.tt=0 and a.kdpiutang like '%" & TXTCARI & "%'  order by c.tahun,c.bln"


End If



Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub DataGrid1_DblClick()
On Error GoTo hell

KODE = 2
Call max


sqlX = "insert into Tanda_terima values('" & rs!kdpiutang & "','" & Format(TTerima_D.txttglTT, "yyyy/MM/dd") & "')"
con.Execute (sqlX)

sqlX = "update piutangsewa set tt=1 where kdpiutang='" & rs!kdpiutang & "'"
con.Execute (sqlX)

TimerAll.Interval = 10
TTerima_D.TimerAll.Interval = 10
TTerima.TimerAll.Interval = 10
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyUp Then

    If rs.AbsolutePosition = 1 Then
    TXTCARI.SetFocus
    End If

ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

On Error GoTo hell

If KeyAscii = 13 Then
    
KODE = 2
Call max
    
sqlX = "insert into Tanda_terima values('" & rs!kdpiutang & "','" & Format(TTerima_D.txttglTT, "yyyy/MM/dd") & "')"
con.Execute (sqlX)

sqlX = "update piutangsewa set tt=1 where kdpiutang='" & rs!kdpiutang & "'"
con.Execute (sqlX)



TimerAll.Interval = 10
TTerima_D.TimerAll.Interval = 10
TTerima.TimerAll.Interval = 10

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerAll.Interval = 10
End Sub




Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If KODE = 2 Then
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
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub










