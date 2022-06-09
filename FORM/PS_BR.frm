VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form PS_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   12840
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
      TabIndex        =   0
      Top             =   1485
      Width           =   2490
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   2
      Top             =   855
      Width           =   11850
      _Version        =   524288
      _ExtentX        =   20902
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   990
      TabIndex        =   1
      Top             =   6930
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
      Picture         =   "PS_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   4875
      Left            =   225
      TabIndex        =   7
      Top             =   1845
      Width           =   11850
      _cx             =   20902
      _cy             =   8599
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"PS_BR.frx":6862
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
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   7695
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
      Left            =   270
      TabIndex        =   5
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Label lbljudul 
      BackStyle       =   0  'Transparent
      Caption         =   "Pinjam Pakai"
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
      TabIndex        =   4
      Top             =   135
      Width           =   7755
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   12240
      Picture         =   "PS_BR.frx":6908
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Label lblkdkategori 
      Caption         =   "lblkategori"
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   7695
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   7440
      Left            =   0
      Picture         =   "PS_BR.frx":6CC8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "PS_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim sql1, sql2, sql As String


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
.Width = 60
.Caption = "KODE"
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "CUSTOMER"
.Width = 300
End With

With datagrid1.Columns(2)
.Caption = "ALAMAT"
.Width = 380
End With



Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

sql1 = "select a.kdcustomer,sum(b.unit) as unit,'PINJAM' as status from Pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam group by kdcustomer Union All" & vbCrLf & _
       "select a.kdcustomer,-sum(b.unit) as unit,'PINJAM' as status from RPinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam group by kdcustomer Union All" & vbCrLf & _
       "select a.kdcustomer,sum(b.unit) as unit,'SEWA' as status from SEWA a left join SEWA_d b on a.kdsewa=b.kdsewa group by kdcustomer Union All" & vbCrLf & _
       "select a.kdcustomer,-sum(b.unit) as unit,'SEWA' as status from RSewa a left join Rsewa_d b on a.kdRSewa=b.kdRSewa group by kdcustomer"

If LBLKODE = "RPINJAM_D" Then
sql2 = "select kdcustomer,sum(unit) as unit from (" & sql1 & ") a where status='PINJAM' group by kdcustomer"
    
ElseIf LBLKODE = "RSEWA_D" Then
sql2 = "select kdcustomer,sum(unit) as unit from (" & sql1 & ") a where status='SEWA' group by kdcustomer"

End If

If TXTCARI = "" Then
sql = "select a.kdcustomer,b.nmcustomer,b.alamat,a.unit from (" & sql2 & ") a left join customer b on a.kdcustomer =b.kdcustomer where a.unit<>0  order by b.nmcustomer,b.alamat"
Else
sql = "select a.kdcustomer,b.nmcustomer,b.alamat,a.unit from (" & sql2 & ") a left join customer b on a.kdcustomer =b.kdcustomer where a.unit<>0 and (b.nmcustomer like '%" & TXTCARI & "%' or a.kdcustomer like '%" & TXTCARI & "%' or b.alamat like '%" & TXTCARI & "%') order by b.nmcustomer,b.alamat"
End If



Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell
If LBLKODE = UCase("RPINJAM_D") Then
Rpinjam_D.lblkdcustomer = rs!kdcustomer
Rpinjam_D.lblnmcustomer = rs!nmcustomer
Rpinjam_D.lblalamat = rs!alamat
ElseIf LBLKODE = UCase("RSEWA_D") Then
RSewa_d.lblkdcustomer = rs!kdcustomer
RSewa_d.lblnmcustomer = rs!nmcustomer
RSewa_d.lblalamat = rs!alamat

End If
Unload Me

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
    
    If LBLKODE = UCase("RPINJAM_D") Then
    Rpinjam_D.lblkdcustomer = rs!kdcustomer
    Rpinjam_D.lblnmcustomer = rs!nmcustomer
    Rpinjam_D.lblalamat = rs!alamat
    ElseIf LBLKODE = UCase("RSEWA_D") Then
    RSewa_d.lblkdcustomer = rs!kdcustomer
    RSewa_d.lblnmcustomer = rs!nmcustomer
    RSewa_d.lblalamat = rs!alamat
       
    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("v") Or KeyAscii = Asc("V") Then
    
 If LBLKODE = UCase("RPINJAM_D") Then
 PS_D_BR.LBLKODE = "RPINJAM_D"
 PS_D_BR.lblKd = rs!kdpinjam
 PS_D_BR.Show vbModal
 ElseIf LBLKODE = UCase("RSEWA_D") Then
 PS_D_BR.LBLKODE = "RSEWA_D"
 PS_D_BR.lblKd = rs!kdsewa
 PS_D_BR.lbljudul = "Detail Sewa"
 PS_D_BR.Show vbModal
 End If

End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerALL.Interval = 10
End Sub




Private Sub TimerAll_Timer()
On Error Resume Next
Call all

TimerALL.Interval = 0
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerALL.Interval = 10
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








