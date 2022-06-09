VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form REAL_PS_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9795
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttgl1 
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
      Left            =   17775
      TabIndex        =   6
      Top             =   1575
      Width           =   1410
   End
   Begin VB.Timer TimerG 
      Left            =   7920
      Top             =   3105
   End
   Begin VB.Timer TimerALL 
      Left            =   8460
      Top             =   3105
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
      Top             =   1575
      Width           =   2490
   End
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
      Left            =   1215
      TabIndex        =   0
      Top             =   990
      Width           =   1590
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   8
      Top             =   810
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
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1395
      TabIndex        =   9
      Top             =   9315
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
      Picture         =   "REAL_PS.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR4 
      Height          =   420
      Left            =   13680
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   945
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
      Picture         =   "REAL_PS.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC4 
      Height          =   420
      Left            =   14220
      TabIndex        =   2
      Top             =   945
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
      Picture         =   "REAL_PS.frx":9094
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   10
      Top             =   1485
      Width           =   19230
      _Version        =   524288
      _ExtentX        =   33920
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   870
      Left            =   19395
      TabIndex        =   7
      ToolTipText     =   "Pilih Semua"
      Top             =   1980
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
      Picture         =   "REAL_PS.frx":B6DE
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   6900
      Left            =   270
      TabIndex        =   5
      Top             =   1980
      Width           =   18915
      _cx             =   33364
      _cy             =   12171
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"REAL_PS.frx":10129
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
   Begin MSComCtl2.DTPicker DTPCari 
      Height          =   330
      Left            =   5805
      TabIndex        =   4
      Top             =   1575
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
      Format          =   91095041
      CurrentDate     =   43923
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Plan :"
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
      Left            =   4860
      TabIndex        =   23
      Top             =   1575
      Width           =   960
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   45
      TabIndex        =   22
      Top             =   9450
      Visible         =   0   'False
      Width           =   1155
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
      TabIndex        =   21
      Top             =   1575
      Width           =   1500
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pinjaman dan Sewa"
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
      TabIndex        =   20
      Top             =   180
      Width           =   5280
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   19350
      Picture         =   "REAL_PS.frx":102FE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3960
      Picture         =   "REAL_PS.frx":106BE
      Stretch         =   -1  'True
      Top             =   1530
      Width           =   420
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
      Left            =   3150
      TabIndex        =   19
      Top             =   1035
      Width           =   915
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
      Left            =   3960
      TabIndex        =   18
      Top             =   990
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
      Left            =   4860
      TabIndex        =   17
      Top             =   990
      Width           =   2670
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
      Left            =   7965
      TabIndex        =   16
      Top             =   1035
      Width           =   1185
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
      Left            =   9180
      TabIndex        =   15
      Top             =   990
      Width           =   1005
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
      Left            =   10215
      TabIndex        =   14
      Top             =   990
      Width           =   3480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL CHEK :"
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
      TabIndex        =   13
      Top             =   1035
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
      Left            =   16110
      TabIndex        =   12
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblpos 
      Caption         =   "1"
      Height          =   330
      Left            =   7695
      TabIndex        =   11
      Top             =   8550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   9780
      Left            =   0
      Picture         =   "REAL_PS.frx":1D56E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20445
   End
End
Attribute VB_Name = "REAL_PS_BR"
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
Dim rsL As ADODB.Recordset
Dim sql, sql1, sql2, sql3, sql4, sql5, sqlX, sqlZ, sqlY, sqlH As String
Dim rs1 As ADODB.Recordset
Dim rsH As ADODB.Recordset


Private Sub cmdBR4_Click()
ACekher_BR.LBLKODE = "REAL_PS_BR"
ACekher_BR.Show vbModal
End Sub

Private Sub cmdC4_Click()
lblnmareaC = ""
lblkdareaC = ""
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

sqlH = "select kode,min(tglno) as tglno from (" & sql5 & ") x group by kode"
Set rsH = con.Execute(sqlH)


If txttglplan = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Cek harus diisi !!", vbCritical, "Error !"
    Exit Sub
ElseIf CDate(txttglplan) < rsH!tglno And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL cek ada yg tidak sesuai", vbCritical, "Error !"
    Exit Sub
Else

'    If IsNull(DTPCari.Value) Then
'    sqladd1 = "select * from (" & sql4 & ") R "
'    Else
'    sqladd1 = "select * from (" & sql4 & ") R where tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' "
'    End If
    
    
    sqladd = "insert into real_cek select '" & Real_cek_TU.lblkdteknisi & "/" & "' + kdcustomer + '" & "/" & Real_cek_TU.txtperiode & "/" & "' + kdbarang ,'" & Real_cek_TU.lblkdteknisi & "/" & "' + kdcustomer + '" & "/" & Real_cek_TU.txtperiode & "','" & Real_cek_TU.txtperiode & "','" & Real_cek_TU.lblkdteknisi & "','" & Format(txttglplan, "yyyy/MM/dd") & "',kdcustomer,kdbarang,total,'','',getdate(),'" & UTAMA.lblkduser & "','" & UCase(UTAMA.lblnmcom) & "' from (" & sql5 & " ) a"
    con.Execute (sqladd)
    
    con.Execute ("update route_plan set keterangan='',det_keterangan='' where idrute in (select '" & Real_cek_TU.lblkdteknisi & "/" & "' + kdcustomer + '" & "/" & Real_cek_TU.txtperiode & "' from (" & sql5 & ") x )")
        

    TimerALL.Interval = 10
    
End If
End Sub


Private Sub datagrid1_Click()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
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

sqlZ = "select kdbarang + '/' + kdcustomer as kdBRX from real_cek where nmrute='" & Real_cek_TU.txtperiode & "' and kdteknisi='" & Real_cek_TU.lblkdteknisi & "'"


sqlY = "select kdbarang,max(tglsj) as tglsj from (" & vbCrLf & _
       "select a.kdbarang,b.tglpinjam as tglSJ from pinjam_d a left join pinjam b on a.kdpinjam=b.kdpinjam Union all select a.kdbarang,b.tglsewa as tglSJ from sewa_d a left join sewa b on a.kdsewa=b.kdsewa" & vbCrLf & _
       ") a where tglsj <= '" & Format(txttgl1, "yyyy/MM/dd") & "' group by kdbarang"


sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' group by a.kdcustomer,b.kdbarang  " & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"


sql2 = "select b.kdareaC,d.nmareaC,a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.kd1,c.nmbarang,f.tglsj,a.pjm,a.swa,a.pjm + a.swa as total,d.lama_cek from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
       "left join barang c on a.kdbarang=c.kdbarang left join  area_cheker d on b.kdareaC=d.kdareaC left join teknisi e on b.kdteknisi= e.kdteknisi left join (" & sqlY & ") f on a.kdbarang=f.kdbarang where " & kata & " and " & kata1 & " and (c.kdkategori between '04' and '10' )  and (a.pjm <> 0 or a.swa<>0)  "


If TXTCARI = "" Then
sql3 = "select * from (" & sql2 & ") a where kdbarang + '/' + kdcustomer not in (select * from (" & sqlZ & ") z) "
Else
sql3 = "select * from (" & sql2 & ") a where kdbarang + '/' + kdcustomer not in (select * from (" & sqlZ & ") z) and (kdbarang like '%" & TXTCARI & "%' or kd1 like '%" & TXTCARI & "%' or kdcustomer like '%" & TXTCARI & "%' or nmcustomer like '%" & TXTCARI & "%' or alamat like '%" & TXTCARI & "%' or nmbarang like '%" & TXTCARI & "%' or nmareac like '% txtcari  %') "
End If

sql4 = "select x.*,y.tglPLAN,convert(date,getdate() - x.lama_cek ) as tglNO from (" & sql3 & ") x left join (select kdcustomer,tglplan from route_plan where nmrute='" & Real_cek_TU.txtperiode & "') y on x.kdcustomer=y.kdcustomer "

If IsNull(DTPCari.Value) Then
sql5 = "select *,'1' as kode from (" & sql4 & ") R "
Else
sql5 = "select *,'1' as kode from (" & sql4 & ") R where tglplan='" & Format(DTPCari, "yyyy/MM/dd") & "' "
End If

sql = "select * from (" & sql5 & ") T order by tglplan,nmareaC,nmcustomer,alamat"

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

Call LG

MousePointer = vbDefault
End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell

LBLKODE = 2
If txttglplan = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL Cek harus diisi !!", vbCritical, "Error !"
    Exit Sub
ElseIf txttglplan < CDate(rs!tglno) And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "TGL cek maximal " & CInt(rs!lama_cek) & " hari dari tgl Skrg !!", vbCritical, "Error !"
    Exit Sub
Else

    MousePointer = vbHourglass


    sqladd = "insert into real_cek values ('" & Real_cek_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & Real_cek_TU.txtperiode & "/" & rs!kdbarang & "','" & Real_cek_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & Real_cek_TU.txtperiode & "','" & Real_cek_TU.txtperiode & "','" & Real_cek_TU.lblkdteknisi & "','" & Format(txttglplan, "yyyy/MM/dd") & "','" & rs!kdcustomer & "','" & rs!kdbarang & "'," & rs!total & ",'','',getdate(),'" & UTAMA.lblkduser & "','" & UCase(UTAMA.lblnmcom) & "' )"
    con.Execute (sqladd)
    
    con.Execute ("update route_plan set keterangan='',det_keterangan='' where idrute='" & Real_cek_TU.lblkdteknisi & "/" & rs!kdcustomer & "/" & Real_cek_TU.txtperiode & "'")
    
    TimerALL.Interval = 10
    

'    If Real_cek_TU.lblfrm = "FIXRUTE_TU" Then
'    fixrute_TU.TimerALL.Interval = 10
'    End If
    
    MousePointer = vbDefault
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
MousePointer = vbDefault
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
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
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

txttgl1 = Date

txttglplan = Date

DTPCari.Value = Date
DTPCari.Value = Null


TimerALL.Interval = 10
End Sub




Private Sub Form_Unload(Cancel As Integer)
Real_cek_TU.TimerALL.Interval = 1000
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
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
If TXTCARI = "" Then
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












Private Sub txttgl1_Change()
Call nul(txttgl1)
TimerALL.Interval = 10
End Sub

Private Sub txttgl1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttgl1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttgl1_KeyPress(KeyAscii As Integer)
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

Private Sub txttgl1_LostFocus()
On Error GoTo hell

txttgl1 = FormatDateTime(txttgl1, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttgl1.SetFocus

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



