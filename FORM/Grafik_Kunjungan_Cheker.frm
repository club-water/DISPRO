VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafik_Kunjungan_Cheker 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   11115
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkdetail 
      BackColor       =   &H00000000&
      Caption         =   "Tampilkan Detail"
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
      Left            =   11745
      TabIndex        =   1
      Top             =   90
      Width           =   1950
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   8700
      Left            =   180
      OleObjectBlob   =   "Grafik_Kunjungan_Cheker.frx":0000
      TabIndex        =   5
      Top             =   2250
      Width           =   19725
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
      Left            =   9855
      TabIndex        =   9
      Top             =   45
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1815
      Left            =   8910
      TabIndex        =   6
      Top             =   855
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
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   45
      Width           =   2670
   End
   Begin VB.Timer TimerAll 
      Left            =   11925
      Top             =   270
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   2
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
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   1635
      Left            =   180
      TabIndex        =   7
      Top             =   540
      Width           =   9330
      _cx             =   16457
      _cy             =   2884
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Grafik_Kunjungan_Cheker.frx":2929
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
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   16740
      TabIndex        =   8
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
      Picture         =   "Grafik_Kunjungan_Cheker.frx":2A0C
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   9360
      TabIndex        =   0
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
      Picture         =   "Grafik_Kunjungan_Cheker.frx":926E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid dataGridT1 
      Height          =   1635
      Left            =   9585
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   10230
      _cx             =   18045
      _cy             =   2884
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Grafik_Kunjungan_Cheker.frx":BAA0
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
      Left            =   8685
      TabIndex        =   10
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Performa Kujungan Cheker"
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
      TabIndex        =   3
      Top             =   -45
      Width           =   8475
   End
   Begin VB.Image Image1 
      Height          =   11085
      Left            =   0
      Picture         =   "Grafik_Kunjungan_Cheker.frx":BC18
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20355
   End
End
Attribute VB_Name = "Grafik_Kunjungan_Cheker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rsT As ADODB.Recordset
Dim rsA As ADODB.Recordset
Dim i As Integer

Dim color As Long, flag As Byte

Private Sub chkdetail_Click()
If chkdetail.Value = 0 Then
dataGridT1.Visible = False
Else
dataGridT1.Visible = True
End If
End Sub

Private Sub CMBJNS_Click()
TimerALL.Interval = 10
End Sub

Private Sub cmdBR1_Click()
fixrute_BR2.lblkode = "GRAFIK_KUNJUNGAN"
fixrute_BR2.Show vbModal
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

'Private Sub Form_Activate()
'    On Error GoTo err
'    color = vbBlue
'    flag = flag Or LWA_COLORKEY
'    SetTransparan1 Me.hwnd, color, 0, flag
'
'    Exit Sub
'err: MsgBox err.Description & " Source : " & err.Source
'End Sub





Private Sub all()

MousePointer = vbHourglass

If CMBJNS.ListIndex <= 9 Then
MSChart1.chartType = CMBJNS.ListIndex
Else
MSChart1.chartType = 14
End If

'MSChart1.ShowLegend = True

'sql1 = "select nmrute,kdteknisi,jmlcust,jmlcek,jmlcust - jmlCek as JmlBlomcek from (" & vbCrLf & _
'        "select nmrute,kdteknisi,SUM(jmlcust) as jmlcust,SUM(jmlcek) as jmlcek from (" & vbCrLf & _
'            "select nmrute,kdteknisi,COUNT(kdcustomer) as jmlcust,0 as jmlcek from route_plan where nmrute='" & txtperiode & "' group by nmrute,kdteknisi" & vbCrLf & _
'            "Union All" & vbCrLf & _
'            "select nmrute,kdteknisi,0 as jmlcust,SUM(jmlcek) as jmlcek from (" & vbCrLf & _
'                "select nmrute,kdteknisi,kdcustomer, 1 as jmlcek from Real_Cek where nmrute='" & txtperiode & "' and kdcustomer in (select kdcustomer from route_plan where nmrute='" & txtperiode & "' ) group by nmrute,kdteknisi,kdcustomer" & vbCrLf & _
'            ") a group by nmrute,kdteknisi" & vbCrLf & _
'        ") X group by nmrute,kdteknisi" & vbCrLf & _
'       ")  Y"
'
'sql = "select b.nmteknisi as CHECKER,a.jmlcust as TARGET,a.jmlcek as [SDH CEK],a.jmlblomcek as [BLM CEK] from (" & sql1 & ") a left join teknisi b on a.kdteknisi=b.kdteknisi"
       
sql = "exec sp_grafik1 @rute='" & txtperiode & "'"
sqlA = "exec sp_grafik1_a @rute='" & txtperiode & "'"

'sql1 = "exec sp_grafik_R1 @rute='" & txtperiode & "'"

sql1 = "select *, T2 - T1 as Ovr_Plan from (" & vbCrLf & _
       "select * ,isnull(DATEDIFF(DAY,tglP1,tglP2) + 1,0) as T1, isnull(DATEDIFF(DAY,tglR1,tglR2) +1 ,0) as T2 from (" & vbCrLf & _
       "select a.kdteknisi,e.nmteknisi ,a.data_perTgl ,Min(tglplan) as tglP1,ISNULL(b.tglP2,'1900/01/01') as tglP2,ISNULL(c.tglR1,'1900/01/01') as tglR1,ISNULL(d.tglR2,'1900/01/01') as tglR2  from ROUTE_PLAN a left join" & vbCrLf & _
       "(select kdteknisi,MAX(tglplan) as tglP2   from ROUTE_PLAN where nmrute='" & txtperiode & "' group by kdteknisi) b on a.kdteknisi =b.kdteknisi left join" & vbCrLf & _
       "(select kdteknisi,Min(tglcek) as tglR1 from real_cek where nmrute='" & txtperiode & "' group by kdteknisi) c on a.kdteknisi =c.kdteknisi left join" & vbCrLf & _
       "(select kdteknisi,Max(tglcek) as tglR2 from real_cek where nmrute='" & txtperiode & "' group by kdteknisi) d on a.kdteknisi =d.kdteknisi left join Teknisi e on a.kdteknisi =e.kdteknisi" & vbCrLf & _
       "where a.nmrute='" & txtperiode & "' group by a.kdteknisi,a.data_perTgl,b.tglP2,c.tglR1,d.tglR2,e.nmteknisi" & vbCrLf & _
    ") x ) y  order by nmteknisi"

Set rs = con.Execute(sql)
Set rs1 = con.Execute(sql1)
Set rsA = con.Execute(sqlA)

Set datagrid1.DataSource = rs
Set DataGrid2.DataSource = rsA
Set MSChart1.DataSource = rsA
'
Set dataGridT1.DataSource = rs1
'
With MSChart1.Legend
    .Location.LocationType = VtChLocationTypeBottom
    .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
    .TextLayout.WordWrap = True

End With
'
'MSChart1.Title = "Performa Kunjungan Cheker ( " & txtperiode & " )"
''With MSChart1.Title.VtFont
'    .Name = "verdana"
'    .Size = 14
'
'    .Effect = VtFontEffectUnderline
'End With


'For i = 1 To (DataGrid1.Rows - 1)
'
'DataGrid1.TextMatrix(i, 0) = i
'
'Next




'MSChart1.Footnote = "SDH CEK1 : Terkunjungi Dengan Melihat Unit" & vbCr & "SDH CEK2 : Terkunjungi Tanpa Melihat Unit"

'With MSChart1.Plot.Axis(1, 1)
'.AxisTitle.VtFont.Size = 10
'.AxisTitle.Visible = True
'.AxisTitle.VtFont.Effect = Bold
'.AxisTitle.VtFont.Style = verdana
'.AxisTitle.Text = "JML CUSTOMER"
'End With
''
'With MSChart1.Plot.Axis(0, 1)
'.AxisTitle.VtFont.Size = 10
'.AxisTitle.Visible = True
'.AxisTitle.VtFont.Effect = Bold
'.AxisTitle.VtFont.Style = verdana
'.AxisTitle.Text = "NAMA CHEKER"
'End With
'
'' mengatur judul grafik
'    MSChart1.Title = "Performa Tim Cheker"
'    With MSChart1.Title.VtFont
'        .Name = "Calibri"
'        .Size = 20
'        .Effect = VtFontEffectUnderline
'        .VtColor.Blue = True
'    End With

MousePointer = vbDefault

End Sub




Private Sub datagrid1_DblClick()
If rs.RecordCount <> 0 Then
sqlT = "select * from teknisi where nmteknisi='" & rs!cheker & "'"
Set rsT = con.Execute(sqlT)

Grafik_D.lblkdteknisi = rsT!kdteknisi
Grafik_D.lbljmlcustomer = rs![sdh cek1] + rs![sdh cek2] + rs![kdl cek]
Grafik_D.Show vbModal
End If
End Sub


Private Sub dataGridT1_DblClick()
If rs.RecordCount <> 0 Then
Grafik_D1.lblkdteknisi = rs1!kdteknisi
Grafik_D1.lblnmteknisi = rs1!nmteknisi
Grafik_D1.lbldata_perTgl = rs1!data_pertgl
Grafik_D1.lblnmrute = Grafik_Kunjungan_Cheker.txtperiode
Grafik_D1.Show vbModal
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()

'GradientForm Me, 0



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
.ListIndex = 10
End With

chkdetail.Value = 1

TimerALL.Interval = 10
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width
Image1.Height = Me.Height

MSChart1.Height = Me.Height - 2500
MSChart1.Width = Me.Width - 600
cmdCANCEL.Width = Me.Width - 100
End Sub

Private Sub TimerAll_Timer()
On Error GoTo hell
Call all


TimerALL.Interval = 0

MousePointer = vbDefault
Exit Sub
hell:
MousePointer = vbDefault
MsgBox err.Description
TimerALL.Interval = 0
End Sub

Private Sub txtperiode_Change()
TimerALL.Interval = 10
End Sub
