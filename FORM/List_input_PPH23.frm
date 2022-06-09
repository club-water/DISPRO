VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form List_input_PPH23 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   17235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   1
      Top             =   585
      Width           =   15990
      _Version        =   524288
      _ExtentX        =   28205
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1215
      TabIndex        =   2
      Top             =   5265
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
      Picture         =   "List_input_PPH23.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   4380
      Left            =   135
      TabIndex        =   0
      Top             =   675
      Width           =   15990
      _cx             =   28205
      _cy             =   7726
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"List_input_PPH23.frx":6862
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
      Begin VB.TextBox DGPPH23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11790
         TabIndex        =   5
         Text            =   "dgtglplan"
         Top             =   720
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   16290
      Picture         =   "List_input_PPH23.frx":69B1
      Stretch         =   -1  'True
      Top             =   225
      Width           =   285
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input PPH 23"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   765
      TabIndex        =   4
      Top             =   45
      Width           =   5280
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   5805
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   5730
      Left            =   0
      Picture         =   "List_input_PPH23.frx":6D71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17205
   End
End
Attribute VB_Name = "List_input_PPH23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub datagrid1_DblClick()
If datagrid1.Col = 9 Then
KODE = 2
lblpos = rs.AbsolutePosition

DGPPH23.Top = datagrid1.Top + datagrid1.CellTop - 50
DGPPH23.Left = datagrid1.Left + datagrid1.CellLeft - 30


DGPPH23.Text = FormatNumber(rs!rpPPH23, 0)
DGPPH23.Visible = True
DGPPH23.Height = datagrid1.CellHeight
DGPPH23.Width = datagrid1.CellWidth
SendKeys "{Home}+{End}"
DGPPH23.SetFocus
End If

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub DGPPH23_Change()
Call nul(DGPPH23)
End Sub

Private Sub DGPPH23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

MousePointer = vbHourglass

con.Execute ("update byrpiutangsewa set jmlbayar= " & (CCur(rs!jmlbayar) + CCur(rs!rpPPH23)) - CCur(DGPPH23) & ",rpPPH23=" & CCur(DGPPH23) & " where kdbyrpiutang='" & rs!kdbyrPiutang & "'")



DGPPH23.Visible = False

TimerALL.Interval = 10

MousePointer = vbDefault

End If
End Sub

Private Sub DGPPH23_LostFocus()
DGPPH23.Visible = False
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
'On Error GoTo hell
''
'With datagrid1.Columns(0)
'.Width = 90
'.Caption = "KODE"
'.Alignment = dbgCenter
'End With
'
'With datagrid1.Columns(1)
'.Caption = "BARANG"
'.Width = 180
'End With
'
'With datagrid1.Columns(2)
'.Caption = "SATUAN"
'.Width = 70
'.Alignment = dbgCenter
'End With
'
'With datagrid1.Columns(3)
'.Caption = "S. AWAL"
'.Width = 70
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'End With
'
'With datagrid1.Columns(4)
'.Caption = "KELUAR"
'.Width = 70
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'End With
'
'
'With datagrid1.Columns(5)
'.Caption = "S. AKHIR"
'.Width = 70
'.Alignment = dbgRight
'.NumberFormat = "#,###0"
'
'End With
'
'
'
'
'Exit Sub
'hell:

End Sub

Private Sub all()
On Error GoTo hell

sql1 = "select kdpiutang from LHP a left join Customer b on left(a.kdpiutang,6)=b.kdcustomer where b.PPH23=1 and a.tglLHP='" & Format(LHP_D.txttglLHP, "yyyy/MM/dd") & "' and a.status='TERTAGIH' "

sql2 = "select * from byrpiutangSewa where kdpiutang in (" & sql1 & " ) and tglbayar='" & Format(LHP_D.txttglCLR, "yyyy/MM/dd") & "' and keterangan='LHP' and trf=0 and kdkolektor='" & LHP_D.lblkdkolektor & "'"

sql = "select a.kdbyrpiutang,a.kdpiutang,a.urut,a.tglbayar,a.kdcustomer,b.nmcustomer,b.alamat,a.jmlbayar,a.rpPPH23,a.potongan from (" & sql2 & ") a left join customer b on a.kdcustomer=b.kdcustomer"

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
'Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerALL.Interval = 10
End Sub




Private Sub TimerAll_Timer()
Call all

TimerALL.Interval = 0
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub








