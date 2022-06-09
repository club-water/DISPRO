VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Real_cek_TU 
   BorderStyle     =   0  'None
   Caption         =   "f"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerG 
      Left            =   2295
      Top             =   4050
   End
   Begin VB.Timer TimerAll 
      Left            =   1800
      Top             =   4050
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
      Left            =   1620
      TabIndex        =   1
      Top             =   945
      Width           =   1500
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
      Left            =   1350
      TabIndex        =   0
      Top             =   1485
      Width           =   2490
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   450
      TabIndex        =   2
      Top             =   720
      Width           =   18780
      _Version        =   524288
      _ExtentX        =   33126
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   180
      TabIndex        =   3
      Top             =   1440
      Width           =   19050
      _Version        =   524288
      _ExtentX        =   33602
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
      Left            =   19395
      TabIndex        =   4
      ToolTipText     =   "Tambah"
      Top             =   1845
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
      Picture         =   "Real_cek_TU.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   19395
      TabIndex        =   5
      ToolTipText     =   "Report Route Plan"
      Top             =   6570
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
      Picture         =   "Real_cek_TU.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   19395
      TabIndex        =   6
      ToolTipText     =   "Hapus Per Customer"
      Top             =   2790
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
      Picture         =   "Real_cek_TU.frx":6DD9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   19395
      TabIndex        =   7
      ToolTipText     =   "Refresh"
      Top             =   4680
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
      Picture         =   "Real_cek_TU.frx":9E72
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   19395
      TabIndex        =   8
      ToolTipText     =   "Cetak"
      Top             =   5625
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
      Picture         =   "Real_cek_TU.frx":CFEE
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1350
      TabIndex        =   9
      Top             =   10935
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
      Picture         =   "Real_cek_TU.frx":10A4B
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   1125
      TabIndex        =   10
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
      Picture         =   "Real_cek_TU.frx":172AD
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   5
      Left            =   19395
      TabIndex        =   11
      ToolTipText     =   "Hapus Semua"
      Top             =   3735
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
      Picture         =   "Real_cek_TU.frx":19ADF
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   8655
      Left            =   225
      TabIndex        =   12
      Top             =   1845
      Width           =   18960
      _cx             =   33443
      _cy             =   15266
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
      AllowUserResizing=   3
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
      FormatString    =   $"Real_cek_TU.frx":1E0C2
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
      Begin VB.ComboBox DGKeterangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9045
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   630
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.TextBox DGTGLcek 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4455
         TabIndex        =   21
         Text            =   "dgtglplan"
         Top             =   765
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin MSComCtl2.DTPicker DTPCari 
      Height          =   330
      Left            =   7110
      TabIndex        =   23
      Top             =   1485
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16761024
      CheckBox        =   -1  'True
      CustomFormat    =   "dd / MM / yyyy"
      Format          =   90505217
      CurrentDate     =   43923
   End
   Begin VB.Label lblfrm 
      Caption         =   "lblfrm"
      Height          =   330
      Left            =   11115
      TabIndex        =   25
      Top             =   135
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Cek :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   6210
      TabIndex        =   24
      Top             =   1485
      Width           =   960
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   330
      Left            =   6210
      TabIndex        =   20
      Top             =   11070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   7470
      TabIndex        =   19
      Top             =   11070
      Visible         =   0   'False
      Width           =   1545
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
      Left            =   450
      TabIndex        =   18
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Realisasi Route Plan"
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
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      Top             =   45
      Width           =   7260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CHEKER :"
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
      Left            =   3375
      TabIndex        =   16
      Top             =   990
      Width           =   735
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
      Left            =   4095
      TabIndex        =   15
      Top             =   945
      Width           =   870
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
      Left            =   4995
      TabIndex        =   14
      Top             =   945
      Width           =   2940
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3870
      Picture         =   "Real_cek_TU.frx":1E25A
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   420
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
      TabIndex        =   13
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   11490
      Left            =   0
      Picture         =   "Real_cek_TU.frx":2B10A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20445
   End
End
Attribute VB_Name = "Real_cek_TU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rsL1, rsL2 As ADODB.Recordset
Dim rsK As ADODB.Recordset
Dim a As Integer
Dim kode As Integer
Dim rsX As ADODB.Recordset
Dim sqlA, sqlB, sqlC, sqlA1, sqlA2 As String
Dim color As Long, flag As Byte
Dim rsA As ADODB.Recordset
Dim rsB As ADODB.Recordset
Dim sql1, sqlK1 As String


Private Sub cmdBR1_Click()
Fixrute_BR.LBLKODE = "REAL_CEK_TU"
Fixrute_BR.Show vbModal

End Sub

Private Sub cmdBR1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If

End Sub


Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then

    If txtperiode <> "" Then
        REAL_PS_BR.lblkdteknisi = lblkdteknisi
        REAL_PS_BR.lblnmteknisi = lblnmteknisi
        REAL_PS_BR.txttgl1 = Format(Date, "dd/MM/yyyy")
        
        If UTAMA.lblstatus = 0 Then
        REAL_PS_BR.txttgl1.Enabled = False
        Else
        REAL_PS_BR.txttgl1.Enabled = True
        End If
        
        REAL_PS_BR.Show vbModal
    Else
        MsgBox "Isi Dulu Nama Rutenya !", vbExclamation, "Warning !"
        End
    End If

ElseIf Index = 2 Then
Call hps

ElseIf Index = 5 Then
Call hps_ALL

ElseIf Index = 4 Then
Real_cek_List.Show vbModal

ElseIf Index = 3 Then
TimerAll.Interval = 10

    

End If
End Sub






Private Sub datagrid1_DblClick()
On Error Resume Next

If datagrid1.Col = 2 And UTAMA.lblstatus = 1 Then
kode = 2
lblpos = rs.AbsolutePosition

DGTGLcek.Top = datagrid1.Top + datagrid1.CellTop - 110
DGTGLcek.Left = datagrid1.Left + datagrid1.CellLeft

DGTGLcek = rs!tglcek
DGTGLcek.Visible = True
DGTGLcek.Height = datagrid1.CellHeight
DGTGLcek.Width = datagrid1.CellWidth
DGTGLcek.SetFocus
SendKeys "{Home}+{End}"

ElseIf datagrid1.Col = 11 Then
kode = 2
lblpos = rs.AbsolutePosition

DGKeterangan.Top = datagrid1.Top + datagrid1.CellTop - 150
DGKeterangan.Left = datagrid1.Left + datagrid1.CellLeft


DGKeterangan.Text = rs!keterangan
DGKeterangan.Visible = True
DGKeterangan.Height = datagrid1.CellHeight
DGKeterangan.Width = datagrid1.CellWidth
DGKeterangan.SetFocus


End If

End Sub



Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyHome Then
rs.MoveFirst
ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
End If
End Sub

Private Sub DGketerangan_KeyPress(KeyAscii As Integer)
On Error GoTo hell

If KeyAscii = 13 Then

MousePointer = vbHourglass


con.Execute ("update real_cek set keterangan='" & UCase(DGKeterangan.Text) & "' where nmrute='" & txtperiode & "' and kdcustomer='" & rs!kdcustomer & "'")
DGKeterangan.Visible = False

ms = InputBox("Input Detail keterangan !", "Detail Keterangan", rs!det_keterangan)

con.Execute ("update real_cek set det_keterangan='" & Trim(UCase(ms)) & "' where nmrute='" & txtperiode & "' and kdcustomer='" & rs!kdcustomer & "'")


TimerAll.Interval = 10
    If lblfrm = "FIXRUTE_TU" Then
       fixrute_TU.TimerAll.Interval = 10
    End If

MousePointer = vbDefault

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
MousePointer = vbDefault
End Sub

Private Sub DGketerangan_LostFocus()
DGKeterangan.Visible = False
End Sub

Private Sub DGTGLcek_Change()
Call nul(DGTGLcek)
End Sub

Private Sub DGTGLcek_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub DGTGLcek_KeyPress(KeyAscii As Integer)
On Error GoTo hell


If KeyAscii = 13 Then

MousePointer = vbHourglass

DGTGLcek = FormatDateTime(DGTGLcek, vbGeneralDate)

con.Execute ("update real_cek set tglcek='" & Format(DGTGLcek, "yyyy/MM/dd") & "' where idcek='" & rs!idcek & "'")
DGTGLcek.Visible = False

TimerAll.Interval = 10

    If lblfrm = "FIXRUTE_TU" Then
       fixrute_TU.TimerAll.Interval = 10
    End If

End If

MousePointer = vbDefault

Exit Sub
hell:
MsgBox "Format Tgl Tidak Sesuai", vbCritical, "Error !"
DGTGLcek.SetFocus
SendKeys "{Home}+{End}"
MousePointer = vbDefault
End Sub

Private Sub DGTGLcek_LostFocus()
DGTGLcek.Visible = False
End Sub

Private Sub DTPCari_Change()
TimerAll.Interval = 10
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Cetak()

Unload AR_PObeli

sqlX = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.keterangan from pobeli_d a left join barang b " & vbCrLf & _
       "on a.kdbarang=b.kdbarang where a.kdpobeli='" & txtkdPO & "' order by a.kdbarang"

Set rsX = con.Execute(sqlX)

With AR_PObeli.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_PObeli
.fldunit.DataField = "unit"
.fldnmbarang.DataField = "nmbarang"
.fldsatuan.DataField = "satuan"
.fldketerangan.DataField = "keterangan"

.lblnoPO = txtkdPO
.lblnmgudang = lblnmgudang
.lbltglPO = Format(txttglPO, "dd/MM/yyyy")
.lbljudul1 = "CUSTOMER"

If txtketerangan = "" Then
.lblNB = ""
Else
.lblNB = "NB : " & txtketerangan
End If

.lbljudul2.Visible = False
.lblkategori.Visible = False


AR_PObeli.Show vbModal

End With

End Sub


Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub

Private Sub tbl()
If rs.RecordCount = 0 Then
    txtperiode.Enabled = True
    cmdBR1.Enabled = True
   
    cmdT(2).Enabled = False
    datagrid1.Enabled = False
    
   


Else
    txtperiode.Enabled = False
    cmdBR1.Enabled = False
   
    cmdT(2).Enabled = True
    datagrid1.Enabled = True
    
   
   
End If
End Sub


Private Sub LG()
On Error GoTo hell
Call tbl

Exit Sub
hell:
End Sub


Private Sub all()

MousePointer = vbHourglass
    
    
    If txtcari = "" Then
    sql1 = "select a.idcek,a.tglcek,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,a.kdbarang,d.kd1,d.nmbarang,a.unit,a.keterangan,a.det_keterangan,a.tglinput from Real_cek a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join barang d on a.kdbarang=d.kdbarang where a.kdteknisi='" & lblkdteknisi & "' and  nmrute= '" & txtperiode & "'"
    Else
    sql1 = "select a.idcek,a.tglcek,c.nmareaC,a.kdcustomer,b.nmcustomer ,b.alamat,a.kdbarang,d.kd1,d.nmbarang,a.unit,a.keterangan,a.det_keterangan,a.tglinput from Real_cek a left join Customer b " & vbCrLf & _
           "on a.kdcustomer=b.kdcustomer left join area_cheker c on b.kdareaC=c.kdareaC left join barang d on a.kdbarang=d.kdbarang where a.kdteknisi='" & lblkdteknisi & "' and  nmrute= '" & txtperiode & "' and (a.kdcustomer like '%" & txtcari & "%' or b.nmcustomer like '%" & txtcari & "%' or b.alamat like '%" & txtcari & "%' or A.kdbarang like '%" & txtcari & "%' or d.kd1 like '%" & txtcari & "%' or d.nmbarang like '%" & txtcari & "%' or a.keterangan like '%" & txtcari & "%')"
    

    End If
    
    If IsNull(DTPCari.Value) Then
    sql = "select * from (" & sql1 & ") a  order by a.tglcek desc,a.tglinput,a.nmcustomer ,a.alamat"
    
    Else
    sql = "select * from (" & sql1 & ") a where a.tglcek='" & Format(DTPCari, "yyyy/MM/dd") & "'  order by a.tglcek desc,a.tglinput,a.nmcustomer ,a.alamat"
    
    End If


Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs


For i = 1 To (datagrid1.Rows - 1)

If rs.RecordCount <> 0 Then
datagrid1.TextMatrix(i, 0) = i
End If

If datagrid1.TextMatrix(i, 11) <> "" Then
datagrid1.Cell(flexcpForeColor, i, 11) = vbRed
datagrid1.Cell(flexcpBackColor, i, 11) = vbYellow
datagrid1.Cell(flexcpFontBold, i, 11) = True

datagrid1.Cell(flexcpForeColor, i, 12) = vbRed
datagrid1.Cell(flexcpBackColor, i, 12) = vbYellow
datagrid1.Cell(flexcpFontBold, i, 12) = True
End If


Next

Call LG


MousePointer = vbDefault
End Sub



Private Sub hps()
On Error GoTo hell


    MousePointer = vbHourglass
    kode = 2
    Call max
    
    
    ms = MsgBox("Apakah anda ingin menghapus data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql = "delete from real_cek where idcek  ='" & rs!idcek & "' "
        con.Execute (sql)
        TimerAll.Interval = 10
        
        
        
    End If

    MousePointer = vbDefault

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
MousePointer = vbDefault
End Sub


Private Sub hps_ALL()
On Error Resume Next

MousePointer = vbHourglass

    kode = 2
    Call max
    

    ms = MsgBox("Apakah anda ingin menghapus Semua data ini ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
    
        If IsNull(DTPCari.Value) Then
        sql2 = "select * from (" & sql1 & ") a "
        Else
        sql2 = "select * from (" & sql1 & ") a where a.tglcek='" & Format(DTPCari, "yyyy/MM/dd") & "' "
        End If
        
        sql = "delete from real_cek where idcek in (select idcek from (" & sql2 & ") a)"
        con.Execute (sql)
        
        
        
        TimerAll.Interval = 10
       
        
    End If
MousePointer = vbDefault

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

DTPCari.Value = Date
DTPCari.Value = Null


Call nul(txtperiode)
Call nul(lblkdteknisi)
Call nul(lblnmteknisi)

sqlK = "Select * from alasan_cek where kebutuhan='R' order by nmalasan"
Set rsK = con.Execute(sqlK)

rsK.MoveFirst

Do While Not rsK.EOF
DGKeterangan.AddItem rsK!nmalasan
rsK.MoveNext
Loop

DGKeterangan.ListIndex = 0

TimerAll.Interval = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lblfrm = "FIXRUTE_TU" Then
        fixrute_TU.TimerAll.Interval = 1000
    Else
        Real_Cek.TimerAll.Interval = 1000
    End If
End Sub

Private Sub lblkdteknisi_Change()
Call nul(lblkdteknisi)
End Sub

Private Sub lblnmteknisi_Change()
Call nul(lblnmteknisi)
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next

Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerAll.Interval = 0
MousePointer = vbDefault

End Sub



Private Sub TXTCARI_Change()
If txtcari = "" Then
TimerAll.Interval = 0
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
TimerAll.Interval = 10
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub


Private Sub txtperiode_Change()
Call nul(txtperiode)
End Sub

Private Sub txtperiode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtperiode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
Beep
KeyAscii = 0
End If
End Sub

Private Sub txtperiode_LostFocus()
txtperiode = UCase(txtperiode)
End Sub


