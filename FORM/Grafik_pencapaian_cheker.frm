VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafik_pencapaian_cheker 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10230
   ScaleWidth      =   19830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbkategori 
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
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   540
      Width           =   2310
   End
   Begin VB.Timer TimerAll 
      Left            =   11925
      Top             =   270
   End
   Begin VB.ComboBox CMBJNS 
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
      Left            =   6660
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   45
      Width           =   2670
   End
   Begin VB.TextBox txtperiode 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   10755
      TabIndex        =   1
      Top             =   45
      Width           =   1500
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   9645
      Left            =   5715
      OleObjectBlob   =   "Grafik_pencapaian_cheker.frx":0000
      TabIndex        =   0
      Top             =   540
      Width           =   14010
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   6075
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   4
      Top             =   450
      Width           =   19590
      _Version        =   524288
      _ExtentX        =   34555
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   16740
      TabIndex        =   5
      Top             =   45
      Width           =   3165
      _ExtentX        =   5583
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
      Picture         =   "Grafik_pencapaian_cheker.frx":28C9
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   10260
      TabIndex        =   6
      ToolTipText     =   "Simpan"
      Top             =   0
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
      Picture         =   "Grafik_pencapaian_cheker.frx":912B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   1950
      Left            =   135
      TabIndex        =   9
      Top             =   945
      Width           =   5550
      _cx             =   9790
      _cy             =   3440
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16744448
      ForeColorFixed  =   65535
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
      HighLight       =   0
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
      FormatString    =   $"Grafik_pencapaian_cheker.frx":B95D
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "KATEGORI :"
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
      Left            =   225
      TabIndex        =   11
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Perbandingan Pencapaian Cheker"
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
      Height          =   465
      Left            =   270
      TabIndex        =   8
      Top             =   -45
      Width           =   8475
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
      Left            =   9585
      TabIndex        =   7
      Top             =   90
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   11085
      Left            =   90
      Picture         =   "Grafik_pencapaian_cheker.frx":BA20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16350
   End
End
Attribute VB_Name = "Grafik_pencapaian_cheker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rsT As ADODB.Recordset
Dim rsA As ADODB.Recordset
Dim rsB As ADODB.Recordset
Dim i As Integer

Dim color As Long, flag As Byte

Private Sub CMBJNS_Click()
TimerAll.Interval = 10
End Sub

Private Sub cmbkategori_Click()
TimerAll.Interval = 10
End Sub

Private Sub cmdBR1_Click()
fixrute_BR2.Show vbModal
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub all()

MousePointer = vbHourglass

If CMBJNS.ListIndex <= 9 Then
MSChart1.chartType = CMBJNS.ListIndex
Else
MSChart1.chartType = 14
End If


sqlA1 = "select b.nmteknisi,a.jmlcustomer,a.jmlcustomer1,convert(float,a.jmlcustomer1 ) /CONVERT(float,a.jmlcustomer) as PCN_A,convert(float,a.jmlcustomer1 ) /CONVERT(float,a.jmlcustomer) * 100 as PCN_B  from V_rekap_plan_VS_Real a left join teknisi b on a.kdteknisi=b.kdteknisi where a.nmrute='" & txtperiode & "'"
sqlA = "select nmteknisi,jmlcustomer,jmlcustomer1,pcn_A from (" & sqlA1 & ") x order by pcn_A desc"

sqlB = "select nmteknisi,pcn_B as PENCAPAIAN from (" & sqlA1 & ") x order by pcn_B desc"


Set rsA = con.Execute(sqlA)
Set rsB = con.Execute(sqlB)

Set datagrid1.DataSource = rsA
Set DataGrid2.DataSource = rsB
Set MSChart1.DataSource = rsB


'
'With MSChart1.Legend
'
'    .Location.LocationType = VtChLocationTypeRight
'    .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
'    .TextLayout.Orientation = VtOrientationDown
'    .TextLayout.WordWrap = False
'
'End With

MSChart1.FootnoteText = " Data Tarikan Per : " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " "
MSChart1.Title = "Perbandingan Pencapaian Cheker ( " & txtperiode & " ) By Customer"
MousePointer = vbDefault

End Sub

Private Sub all1()

MousePointer = vbHourglass

If CMBJNS.ListIndex <= 9 Then
MSChart1.chartType = CMBJNS.ListIndex
Else
MSChart1.chartType = 14
End If

sqlX = "select B.kdteknisi,a.kdcustomer,sum(a.unit-a.Runit)as qty from V_brg_split a left join " & vbCrLf & _
       "(select kdteknisi,kdcustomer,data_pertgl from ROUTE_PLAN where nmrute='" & txtperiode & "' group by kdteknisi,kdcustomer,data_pertgl ) b on a.kdcustomer=b.kdcustomer where a.tgl<=b.data_pertgl  group by a.kdcustomer,b.kdteknisi"



sqlA1 = "select a.kdteknisi,a.kdcustomer,b.qty from route_plan a left join (" & sqlX & ") b on a.kdcustomer=b.kdcustomer where a.nmrute='" & txtperiode & "'"
sqlA2 = "select kdteknisi,sum(qty) as qty, 0 as qty1 from (" & sqlA1 & ") Q group by kdteknisi"


sqlB1 = "select kdteknisi,kdbarang,1 as qty1 from real_cek where nmrute='" & txtperiode & "'and tglcek <= getdate() "
sqlB2 = "select kdteknisi, 0 as qty,sum(qty1) as qty1 from (" & sqlB1 & ") R group by kdteknisi"

sqlC1 = "select kdteknisi,sum(qty) as qty, sum(qty1) as qty1 from (" & sqlA2 & " union all " & sqlB2 & ") x group by kdteknisi"
sqlC2 = "select b.nmteknisi,a.qty,a.qty1,convert(float,a.qty1) / convert(float,a.qty) as PCN_A,(convert(float,a.qty1) / convert(float,a.qty)) * 100 as PCN_B from (" & sqlC1 & ") a left join teknisi b on a.kdteknisi=b.kdteknisi"

sqlA = "select nmteknisi,qty,qty1,pcn_A from (" & sqlC2 & ") x order by pcn_A desc"

sqlB = "select nmteknisi,pcn_B as PENCAPAIAN from (" & sqlC2 & ") x order by pcn_B desc"




Set rsA = con.Execute(sqlA)
Set rsB = con.Execute(sqlB)

Set datagrid1.DataSource = rsA
Set DataGrid2.DataSource = rsB
Set MSChart1.DataSource = rsB


MSChart1.FootnoteText = " Data Tarikan Per : " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " "
MSChart1.Title = "Perbandingan Pencapaian Cheker ( " & txtperiode & " ) By Jml Unit (Qty)"
MousePointer = vbDefault

End Sub











Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()

With CMBJNS
.AddItem "3D Bar"
.AddItem "2D Bar"
.AddItem "3D Line"
.AddItem "2D LIne"
.AddItem "3D Area"
.AddItem "2D Area"
.AddItem "3D Step"
.AddItem "2D Step"
.AddItem "3D Combination"
.AddItem "2D Combination"
.AddItem "2D Pie"
.ListIndex = 1
End With

cmbkategori.AddItem "BY JML CUSTOMER"
cmbkategori.AddItem "BY JML UNIT (QTY)"
cmbkategori.ListIndex = 0

TimerAll.Interval = 10
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width
Image1.Height = Me.Height

MSChart1.Height = Me.Height - 800
MSChart1.Width = Me.Width - 6200
cmdCANCEL.Width = Me.Width - 100
End Sub

Private Sub TimerAll_Timer()
On Error GoTo hell

If cmbkategori.ListIndex = 0 Then
Call all
Else
Call all1
End If

TimerAll.Interval = 0

MousePointer = vbDefault
Exit Sub
hell:
MousePointer = vbDefault
MsgBox err.Description
TimerAll.Interval = 0
End Sub

Private Sub txtperiode_Change()
TimerAll.Interval = 10
End Sub


