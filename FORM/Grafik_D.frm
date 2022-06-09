VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafik_D 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtoff 
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
      Left            =   6570
      TabIndex        =   0
      Text            =   "0"
      Top             =   90
      Width           =   780
   End
   Begin VB.Timer TimerAll 
      Left            =   270
      Top             =   1305
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
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   1770
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6045
      Left            =   3735
      OleObjectBlob   =   "Grafik_D.frx":0000
      TabIndex        =   4
      Top             =   1395
      Width           =   8250
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   780
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   4785
      _cx             =   8440
      _cy             =   1376
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Grafik_D.frx":2716
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1005
      Left            =   495
      TabIndex        =   7
      Top             =   8010
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1773
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
   Begin VSFlex8UCtl.VSFlexGrid dataGridT1 
      Height          =   780
      Left            =   4905
      TabIndex        =   2
      Top             =   540
      Width           =   7035
      _cx             =   12409
      _cy             =   1376
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Grafik_D.frx":2794
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   315
      TabIndex        =   12
      Top             =   2250
      Width           =   2625
   End
   Begin VB.Label lbljmlcustomer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   315
      TabIndex        =   11
      Top             =   1800
      Width           =   2625
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total yg Terkunjungi"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   315
      TabIndex        =   10
      Top             =   1395
      Width           =   3030
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
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
      Left            =   7380
      TabIndex        =   9
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "OFF Kunjungan :"
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
      Left            =   5220
      TabIndex        =   8
      Top             =   135
      Width           =   1500
   End
   Begin VB.Label lblkdteknisi 
      Caption         =   "Label2"
      Height          =   510
      Left            =   4005
      TabIndex        =   6
      Top             =   8235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Vs Realisasi"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   135
      TabIndex        =   5
      Top             =   0
      Width           =   8475
   End
   Begin VB.Image Image1 
      Height          =   7530
      Left            =   0
      Picture         =   "Grafik_D.frx":2950
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "Grafik_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsX As ADODB.Recordset
Dim i As Integer

Dim color As Long, flag As Byte

Private Sub CMBJNS_Click()
TimerAll.Interval = 10
End Sub



Private Sub all()

MousePointer = vbHourglass

If CMBJNS.ListIndex <= 9 Then
MSChart1.chartType = CMBJNS.ListIndex
Else
MSChart1.chartType = 14
End If


sqlX = "select * from V_rekap_real_cek where kdteknisi='" & lblkdteknisi & "' and nmrute='" & Grafik_Kunjungan_Cheker.txtperiode & "'"
Set rsX = con.Execute(sqlX)


sqlP1 = "select '1' as kode,tgloff from OFF_kunjungan_all union all select '1' as kode,tgloFF_C as tgloff from OFF_kunjungan_C where kdteknisi='" & lblkdteknisi & "'"
sqlP2 = "select kode,tgloff from (" & sqlP1 & ") x where tgloff between '" & Format(rsX!tglawal1, "yyyy/MM/dd") & "' and '" & Format(rsX!tglakhir1, "yyyy/MM/dd") & "'  group by kode,tgloff"
sqlP = "select kode,count(tgloff) as jmloff from (" & sqlP2 & ") x group by kode"
Set rsP = con.Execute(sqlP)

If rsP.RecordCount <> 0 Then
txtoff = rsP!jmloff
Else
txtoff = 0
End If
       
sql = "exec sp_grafik2 @rute='" & Grafik_Kunjungan_Cheker.txtperiode & "',@kdteknisi='" & lblkdteknisi & "'"
sql1 = "exec sp_grafik_R2 @rute='" & Grafik_Kunjungan_Cheker.txtperiode & "',@kdteknisi='" & lblkdteknisi & "',@off_kjgn = " & txtoff & " , @jmlcust=" & lbljmlcustomer & ""




Set rs = con.Execute(sql)
Set rs1 = con.Execute(sql1)

Set DataGrid2.DataSource = rs
Set datagrid1.DataSource = rs
Set dataGridT1.DataSource = rs1

Set MSChart1.DataSource = rs


With MSChart1.Legend
    .Location.LocationType = VtChLocationTypeRight
    .TextLayout.VertAlignment = VtVerticalAlignmentCenter
    .TextLayout.WordWrap = True
    
End With


MousePointer = vbDefault

End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()

'GradientForm Me, 0
Me.Top = Grafik_Kunjungan_Cheker.datagrid1.Top + 300
Me.Left = Grafik_Kunjungan_Cheker.datagrid1.Left + 500


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

If UTAMA.lblstatus = 0 Then
txtoff.Enabled = False
Else
txtoff.Enabled = True
End If

TimerAll.Interval = 10
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width
Image1.Height = Me.Height
End Sub

Private Sub TimerAll_Timer()
On Error GoTo hell
Call all


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


Private Sub txtoff_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtoff_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtoff_KeyPress(KeyAscii As Integer)
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

Private Sub txtoff_LostFocus()
On Error GoTo hell

txtoff = FormatNumber(txtoff, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtoff.SetFocus

End Sub



