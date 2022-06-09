VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Rute_Cheker_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9225
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglplan 
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
      Left            =   1710
      TabIndex        =   0
      Top             =   945
      Width           =   1590
   End
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
      Left            =   1395
      TabIndex        =   3
      Top             =   1530
      Width           =   2490
   End
   Begin VB.Timer TimerALL 
      Left            =   8460
      Top             =   3060
   End
   Begin VB.Timer TimerG 
      Left            =   7920
      Top             =   3060
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   4
      Top             =   765
      Width           =   15945
      _Version        =   524288
      _ExtentX        =   28125
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   8685
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
      Picture         =   "Rute_Cheker_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   6270
      Left            =   135
      TabIndex        =   6
      Top             =   1890
      Width           =   15945
      _cx             =   28125
      _cy             =   11060
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   7.5
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Rute_Cheker_BR.frx":6862
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
      Begin VB.Timer Timerflood 
         Left            =   8820
         Top             =   1170
      End
      Begin C1SizerLibCtl.C1Elastic flood 
         Height          =   465
         Left            =   6705
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   4155
         _cx             =   7329
         _cy             =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   255
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   1
         FloodPercent    =   0
         CaptionPos      =   4
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
   End
   Begin Threed.SSCommand cmdBR4 
      Height          =   420
      Left            =   14175
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   900
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
      Picture         =   "Rute_Cheker_BR.frx":69D9
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC4 
      Height          =   420
      Left            =   14670
      TabIndex        =   2
      Top             =   900
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
      Picture         =   "Rute_Cheker_BR.frx":920B
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   17
      Top             =   1440
      Width           =   16080
      _Version        =   524288
      _ExtentX        =   28363
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   870
      Left            =   16200
      TabIndex        =   20
      ToolTipText     =   "Pilih Semua"
      Top             =   1890
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
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
      Picture         =   "Rute_Cheker_BR.frx":B855
      ButtonStyle     =   4
   End
   Begin VB.Label lblpos 
      Caption         =   "1"
      Height          =   330
      Left            =   7695
      TabIndex        =   21
      Top             =   8505
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbltgl1 
      BackStyle       =   0  'Transparent
      Caption         =   "20/12/2019"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   13770
      TabIndex        =   19
      Top             =   1620
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PER TGL : "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   12060
      TabIndex        =   18
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL ROUTE PLAN :"
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
      Left            =   270
      TabIndex        =   16
      Top             =   990
      Width           =   1500
   End
   Begin VB.Label lblnmareaC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   10710
      TabIndex        =   15
      Top             =   945
      Width           =   3480
   End
   Begin VB.Label lblkdareaC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9675
      TabIndex        =   14
      Top             =   945
      Width           =   1005
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "AREA CHEKER :"
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
      Left            =   8460
      TabIndex        =   13
      Top             =   990
      Width           =   1185
   End
   Begin VB.Label lblnmteknisi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5355
      TabIndex        =   12
      Top             =   945
      Width           =   2670
   End
   Begin VB.Label lblkdteknisi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4455
      TabIndex        =   11
      Top             =   945
      Width           =   870
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   " CHEKER :"
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
      Left            =   3645
      TabIndex        =   10
      Top             =   990
      Width           =   915
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3960
      Picture         =   "Rute_Cheker_BR.frx":102A0
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   16245
      Picture         =   "Rute_Cheker_BR.frx":1D150
      Stretch         =   -1  'True
      Top             =   360
      Width           =   285
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rute Cheker"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   855
      TabIndex        =   9
      Top             =   135
      Width           =   5280
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
      Left            =   225
      TabIndex        =   8
      Top             =   1530
      Width           =   1500
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   270
      TabIndex        =   7
      Top             =   9270
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   9240
      Left            =   0
      Picture         =   "Rute_Cheker_BR.frx":1D510
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   17115
   End
End
Attribute VB_Name = "Rute_Cheker_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim rsAreaC As ADODB.Recordset
Dim rsteknisi As ADODB.Recordset
Dim kata, kata1 As String
Dim sql, sql1, sql2, sqlX As String
Dim rs1 As ADODB.Recordset


Private Sub cmdBR4_Click()
ACekher_BR.LBLKODE = "RUTE_CHEKER_BR"
ACekher_BR.Show vbModal
End Sub




Private Sub cmdC4_Click()
lblkdareaC = ""
lblnmareaC = ""
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub cmdsimpan_Click()
LBLKODE = 1

If txttglplan = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Route Plan harus diisi !!", vbCritical, "Error !"
    Exit Sub
ElseIf CDate(txttglplan) < CDate(fixrute_TU.txttglspk1) Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Planing tidak boleh kurang dari tgl cut off !!", vbCritical, "Error !"
    Exit Sub
ElseIf CDate(txttglplan) < Date And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Planing tidak boleh kurang dari Hari ini !!", vbCritical, "Error !"
    Exit Sub
    
Else
'    flood.Visible = True
'    Timerflood.Interval = 10
'
    If txtcari = "" Then
    sqladd1 = "select nmareaC,nmteknisi,kdcustomer,nmcustomer,alamat,cp,telp,disp,showC,RG,disp+showC+RG as total from (" & sql2 & ") a where kdcustomer not in (" & sqlX & ") "
    Else
    sqladd1 = "select nmareaC,nmteknisi,kdcustomer,nmcustomer,alamat,cp,telp,disp,showC,RG,disp+showC+RG as total from (" & sql2 & ") a where (kdcustomer like '%" & txtcari & "%' or nmcustomer like '%" & txtcari & "%' or alamat like '%" & txtcari & "%') and kdcustomer not in (" & sqlX & ") "
    End If
    
    sqladd = "insert into route_plan select '" & fixrute_TU.lblkdteknisi & "/" & "' + kdcustomer + '" & "/" & fixrute_TU.txtperiode & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','','','" & Format(txttglplan, "yyyy/MM/dd") & "',kdcustomer,total,'" & Format(lbltgl1, "yyyy/MM/dd") & "',getdate(),'" & UTAMA.lblkduser & "' from (" & sqladd1 & ") x"
    
    con.Execute (sqladd)
    
    
    
    fixrute_TU.TimerALL.Interval = 10
    TimerALL.Interval = 10
    
End If
End Sub

Private Sub datagrid1_Click()
On Error Resume Next
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
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

If rs.RecordCount <> 0 Then
cmdsimpan.Enabled = True
Else
cmdsimpan.Enabled = False
End If

Exit Sub
hell:

End Sub

Private Sub all()
MousePointer = vbHourglass


sqlX = "select kdcustomer from route_plan where nmrute='" & fixrute_TU.txtperiode & "' "


sql1 = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
            "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
            "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
            "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= '" & Format(lbltgl1, "yyyy/MM/dd") & "' group by kdcustomer,kdkategori" & vbCrLf & _
            ") a group by kdcustomer " & vbCrLf & _
       ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2+RG <>0"


sql2 = "select d.nmareaC,e.nmteknisi,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,a.disp,a.showC,a.RG from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
       "left join  area_cheker d on b.kdareaC=d.kdareaC left join teknisi e on b.kdteknisi= e.kdteknisi where " & kata & " and " & kata1 & "  "
      
      
If txtcari = "" Then
sql = "select nmareaC,nmteknisi,kdcustomer,nmcustomer,alamat,cp,telp,disp,showC,RG,disp+showC+RG as total from (" & sql2 & ") a where kdcustomer not in (" & sqlX & ") order by nmareaC,nmteknisi,nmcustomer,alamat"
Else
sql = "select nmareaC,nmteknisi,kdcustomer,nmcustomer,alamat,cp,telp,disp,showC,RG,disp+showC+RG as total from (" & sql2 & ") a where (kdcustomer like '%" & txtcari & "%' or nmcustomer like '%" & txtcari & "%' or alamat like '%" & txtcari & "%') and kdcustomer not in (" & sqlX & ") order by nmareaC,nmteknisi,nmcustomer,alamat"
End If


Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
Call LG


For i = 1 To (datagrid1.Rows - 1)
'For j = 1 To (datagrid1.Cols - 1)

If rs.RecordCount <> 0 Then
datagrid1.TextMatrix(i, 0) = i
End If

'If datagrid1.TextMatrix(i, 12) = 1 And datagrid1.TextMatrix(i, 13) < lbltgl1 Then
'datagrid1.Cell(flexcpForeColor, i, j) = &HFF00FF
'End If

'Next
Next


MousePointer = vbDefault
End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell

LBLKODE = 2



If txttglplan = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Route Plan harus diisi !!", vbCritical, "Error !"
    Exit Sub
    

ElseIf CDate(txttglplan) < CDate(fixrute_TU.txttglspk1) Then
    
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Planing tidak boleh kurang dari tgl cut off !!", vbCritical, "Error !"
    Exit Sub

ElseIf CDate(txttglplan) < Date And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Planing tidak boleh kurang dari Hari ini !!", vbCritical, "Error !"
    Exit Sub
Else
    
    
    
     sqladd = "insert into route_plan values ( '" & fixrute_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & fixrute_TU.txtperiode & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','','','" & Format(txttglplan, "yyyy/MM/dd") & "','" & rs!kdcustomer & "'," & rs!total & ",'" & Format(lbltgl1, "yyyy/MM/dd") & "',getdate(),'" & UTAMA.lblkduser & "' )"
     con.Execute (sqladd)
    
    
'     flood.Visible = True
'     Timerflood.Interval = 10
       
    
    fixrute_TU.TimerALL.Interval = 10
    TimerALL.Interval = 10
    
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyUp Then

    If rs.AbsolutePosition = 1 Then
    txtcari.SetFocus
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
    
    If LBLKODE = "CUSTOMER_TU" Then
    Customer_TU.lblkdareaC = rs!kdareaC
    'Customer_TU.lblnmareaC = rs!nmareaC
    ElseIf LBLKODE = "CETAK_9A1" Then
    Cetak_9A1.lblkdareaC = rs!kdareaC

    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

LBLKODE = 1

kata = "b.kdareaC <> '@@@'"
kata1 = "b.kdteknisi <> '@@@'"

txttglplan = Date

TimerALL.Interval = 10
End Sub




Private Sub lblkdareaC_Change()
sqlAreaC = "select a.*,isnull(b.nmteknisi,'') as nmteknisi from area_cheker a left join teknisi b on a.kdteknisi=b.kdteknisi where a.kdareaC='" & lblkdareaC & "'"
Set rsAreaC = con.Execute(sqlAreaC)


If rsAreaC.RecordCount <> 0 Then
lblnmareaC = rsAreaC!nmareaC
Else
lblnmareaC = ""
End If

If lblkdareaC = "" Then
kata = "b.kdareaC <> '@@@'"
Else
kata = "b.kdareaC ='" & lblkdareaC & "'"
End If


TimerALL.Interval = 10
End Sub

Private Sub lblkdteknisi_Change()
sqlteknisi = "select * from teknisi where kdteknisi='" & lblkdteknisi & "'"
Set rsteknisi = con.Execute(sqlteknisi)

If rsteknisi.RecordCount <> 0 Then
lblnmteknisi = rsteknisi!nmteknisi
Else
lblnmteknisi = ""
End If

If lblkdteknisi = "" Then
kata1 = "b.kdteknisi <> '@@@'"
Else
kata1 = "b.kdteknisi ='" & lblkdteknisi & "'"
End If

TimerALL.Interval = 10
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If LBLKODE = 2 Then
rs.AbsolutePosition = lblpos
End If


TimerALL.Interval = 0

MousePointer = vbDefault


End Sub

Private Sub Timerflood_Timer()
 Dim j%
  Static i%

  If i > 90 Then
  i = 0
  
  End If
  
  i = i + 10
  
  If i = 100 Then
  Timerflood.Interval = 0
  flood.Visible = False
  
  
'    If rs.RecordCount = 0 Then
'        SetTimer hwnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
'         MsgBox "Data gak ada Broo !", vbInformation, "Info !"
'    End If

  End If
  

  For j = 0 To 10
    flood.FloodPercent = i
    flood.Caption = i & "%"
  Next j


End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub txtcari_Change()
If txtcari = "" Then
TimerALL.Interval = 10
End If
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
 TimerALL.Interval = 10

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub












Private Sub txttglplan_Change()
Call nul(txttglplan)


End Sub

Private Sub txttglplan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglplan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglplan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttglplan_LostFocus()
On Error GoTo hell

txttglplan = FormatDateTime(txttglplan, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglplan.SetFocus

End Sub

