VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form List_non_route 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5790
   ScaleWidth      =   19290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   2
      Top             =   585
      Width           =   17925
      _Version        =   524288
      _ExtentX        =   31618
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
      TabIndex        =   1
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
      Picture         =   "List_non_route.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   4380
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   17970
      _cx             =   31697
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"List_non_route.frx":6862
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
      Begin VB.TextBox DGTglplan 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10710
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Label lblpos 
      Caption         =   "0"
      Height          =   285
      Left            =   10935
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   5805
      Width           =   1155
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Non Route yang Blom dikunjungi"
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
      TabIndex        =   3
      Top             =   45
      Width           =   8835
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   18270
      Picture         =   "List_non_route.frx":6A1B
      Stretch         =   -1  'True
      Top             =   270
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   5730
      Left            =   0
      Picture         =   "List_non_route.frx":6DDB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19275
   End
End
Attribute VB_Name = "List_non_route"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim kode As Integer
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
On Error Resume Next
If datagrid1.Col = 1 Then
kode = 2
lblpos = rs.AbsolutePosition

DGTglplan.Top = datagrid1.Top + datagrid1.CellTop - 50
DGTglplan.Left = datagrid1.Left + datagrid1.CellLeft - 30

DGTglplan = Date
DGTglplan.Text = rs!tglplan
DGTglplan.Visible = True
DGTglplan.Height = datagrid1.CellHeight
DGTglplan.Width = datagrid1.CellWidth
SendKeys "{Home}+{End}"
DGTglplan.SetFocus
End If

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub DGTGLPlan_Change()
Call nul(DGTglplan)
End Sub

Private Sub DGTGLPlan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
If KeyAscii = 13 Then


    If DGTglplan <> "" Then
    con.Execute ("delete from plan_non_route where idrute = '" & fixrute_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & fixrute_TU.txtperiode & "'")
    con.Execute ("insert into plan_non_route values ('" & fixrute_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & fixrute_TU.txtperiode & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & rs!kdcustomer & "','" & Format(DGTglplan, "yyyy/MM/dd") & "',getdate(),'" & UTAMA.lblkduser & "')")
    
    Else
    con.Execute ("delete from plan_non_route where idrute = '" & fixrute_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & fixrute_TU.txtperiode & "'")
    
    End If
    TimerALL.Interval = 10
    DGTglplan.Visible = False
    datagrid1.SetFocus

End If

Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"

End Sub

Private Sub DGTGLPlan_LostFocus()
DGTglplan.Visible = False
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




Private Sub LG()
End Sub

Private Sub all()
On Error GoTo hell

    sqlNRX = "select kdcustomer from route_plan where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi ='" & fixrute_TU.lblkdteknisi & "'  union all" & vbCrLf & _
             "select kdcustomer from real_cek where nmrute='" & fixrute_TU.txtperiode & "' and kdcustomer not in (select kdcustomer from Route_plan  where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi ='" & fixrute_TU.lblkdteknisi & "')"
    
    sqlNR1 = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
                "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
                "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
                "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                    "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= getdate() group by kdcustomer,kdkategori" & vbCrLf & _
                ") a group by kdcustomer " & vbCrLf & _
           ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2+RG <>0"
    
    
    sqlNR2 = "select d.nmareaC,e.nmteknisi,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,a.disp,a.showC,a.RG from (" & sqlNR1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
             "left join  area_cheker d on b.kdareaC=d.kdareaC left join teknisi e on b.kdteknisi= e.kdteknisi where b.kdteknisi='" & fixrute_TU.lblkdteknisi & "'"
          
    sqlNR3 = "select kdcustomer,tglplan from plan_non_route where nmrute='" & fixrute_TU.txtperiode & "' and kdteknisi='" & fixrute_TU.lblkdteknisi & "'"
          
    sqlNR = "select b.tglplan,c.tglSJ,a.nmareaC,a.nmteknisi,a.kdcustomer,a.nmcustomer,a.alamat,a.cp,a.telp,a.disp,a.showC,a.RG,a.disp+a.showC+a.RG as total from (" & sqlNR2 & ") a left join (" & sqlNR3 & ") b on a.kdcustomer=b.kdcustomer left join V_tglsj c on a.kdcustomer=c.kdcustomer where a.kdcustomer not in (" & sqlNRX & ") order by c.tglSJ desc,a.nmareaC,a.nmteknisi,a.nmcustomer,a.alamat"
    
    Set rs = con.Execute(sqlNR)
    Set datagrid1.DataSource = rs
    
    If rs.RecordCount <> 0 Then
    
        For i = 1 To (datagrid1.Rows - 1)
        For j = 1 To (datagrid1.Cols - 1)
    
        
        datagrid1.TextMatrix(i, 0) = i
    
        
        If datagrid1.TextMatrix(i, 1) = "" Then
        datagrid1.Cell(flexcpForeColor, i, j) = vbRed
        End If
        
        Next
        Next
    End If
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
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerALL.Interval = 10
End Sub




Private Sub TimerAll_Timer()
On Error Resume Next

Call all

If kode = 2 Then
rs.AbsolutePosition = lblpos
End If

TimerALL.Interval = 0
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub









