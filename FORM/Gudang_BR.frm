VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Gudang_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10200
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
      Width           =   9465
      _Version        =   524288
      _ExtentX        =   16695
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   945
      TabIndex        =   1
      Top             =   6435
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
      Picture         =   "Gudang_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   4110
      Left            =   135
      TabIndex        =   6
      Top             =   1935
      Width           =   9375
      _cx             =   16536
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
      FormatString    =   $"Gudang_BR.frx":6862
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
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   180
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gudang"
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
      TabIndex        =   3
      Top             =   135
      Width           =   2715
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   9675
      Picture         =   "Gudang_BR.frx":68F0
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2790
      Picture         =   "Gudang_BR.frx":6CB0
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   6900
      Left            =   0
      Picture         =   "Gudang_BR.frx":13B60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10185
   End
End
Attribute VB_Name = "Gudang_BR"
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

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub LG()
On Error GoTo hell

With datagrid1.Columns(0)
.Width = 70
.Caption = "KODE"
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "GUDANG"
.Width = 250
End With

With datagrid1.Columns(2)
.Caption = "ALAMAT"
.Width = 250
End With

With datagrid1.Columns(3)
.Caption = "KETERANGAN"
.Width = 0
End With



Exit Sub
hell:

End Sub

Private Sub ALL()
On Error GoTo hell

If txtcari = "" Then
sql = "select * from Gudang  order by nmgudang"
Else
sql = "select * from gudang where kdgudang like '%" & txtcari & "%' or nmgudang like '%" & txtcari & "%' order by nmgudang"
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
If LBLKODE = "PObeli_D" Then
PObeli_d.lblkdgudang = rs!kdgudang
PObeli_d.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "PO_D" Then
PO_D.lblkdgudang = rs!kdgudang
PO_D.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "RPINJAM_D" Then
Rpinjam_D.lblkdgudang = rs!kdgudang
Rpinjam_D.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "PERBAIKAN_D" Then
Perbaikan_D.lblkdgudang2 = rs!kdgudang
Perbaikan_D.lblnmgudang2 = rs!nmgudang
ElseIf LBLKODE = "RSEWA_D" Then
RSewa_d.lblkdgudang = rs!kdgudang
RSewa_d.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "1A1" Then
Cetak_1A1.lblkdgudang = rs!kdgudang
Cetak_1A1.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "3A1" Then
Cetak_3A1.lblkdgudang = rs!kdgudang
Cetak_3A1.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "3A2" Then
Cetak_3A2.lblkdgudang = rs!kdgudang
Cetak_3A2.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "1A2" Then
Cetak_1A2.lblkdgudang = rs!kdgudang
Cetak_1A2.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "4A1" Then
Cetak_4A1.lblkdgudang = rs!kdgudang
Cetak_4A1.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "4A2" Then
Cetak_4A2.lblkdgudang = rs!kdgudang
Cetak_4A2.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "5A1" Then
Cetak_5A1.lblkdgudang = rs!kdgudang
Cetak_5A1.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "5A2" Then
Cetak_5A2.lblkdgudang = rs!kdgudang
Cetak_5A2.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "5A3" Then
Cetak_5A3.lblkdgudang = rs!kdgudang
Cetak_5A3.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "5A4" Then
Cetak_5A4.lblkdgudang = rs!kdgudang
Cetak_5A4.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "6A1" Then
Cetak_6A1.lblkdgudang = rs!kdgudang
Cetak_6A1.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "6A2" Then
Cetak_6A2.lblkdgudang = rs!kdgudang
Cetak_6A2.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "3A5" Then
Cetak_3A5.lblkdgudang = rs!kdgudang
Cetak_3A5.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "1A3" Then
Cetak_1A3.lblkdgudang = rs!kdgudang
Cetak_1A3.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "1A4A1" Then
Cetak_1A4A1.lblkdgudang = rs!kdgudang
Cetak_1A4A1.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "1A5" Then
Cetak_1A5.lblkdgudang = rs!kdgudang
Cetak_1A5.lblnmgudang = rs!nmgudang
ElseIf LBLKODE = "S_OPNAME" Then
S_OPname.lblkdgudang = rs!kdgudang
S_OPname.lblnmgudang = rs!nmgudang

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
    
    If LBLKODE = "PObeli_D" Then
    PObeli_d.lblkdgudang = rs!kdgudang
    PObeli_d.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "PO_D" Then
    PO_D.lblkdgudang = rs!kdgudang
    PO_D.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "RPINJAM_D" Then
    Rpinjam_D.lblkdgudang = rs!kdgudang
    Rpinjam_D.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "PERBAIKAN_D" Then
    Perbaikan_D.lblkdgudang2 = rs!kdgudang
    Perbaikan_D.lblnmgudang2 = rs!nmgudang
    ElseIf LBLKODE = "RSEWA_D" Then
    RSewa_d.lblkdgudang = rs!kdgudang
    RSewa_d.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "PERBAIKAN_D" Then
    Perbaikan_D.lblkdgudang2 = rs!kdgudang
    Perbaikan_D.lblnmgudang2 = rs!nmgudang
    ElseIf LBLKODE = "1A1" Then
    Cetak_1A1.lblkdgudang = rs!kdgudang
    Cetak_1A1.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "3A1" Then
    Cetak_3A1.lblkdgudang = rs!kdgudang
    Cetak_3A1.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "3A2" Then
    Cetak_3A2.lblkdgudang = rs!kdgudang
    Cetak_3A2.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "1A2" Then
    Cetak_1A2.lblkdgudang = rs!kdgudang
    Cetak_1A2.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "4A1" Then
    Cetak_4A1.lblkdgudang = rs!kdgudang
    Cetak_4A1.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "4A2" Then
    Cetak_4A2.lblkdgudang = rs!kdgudang
    Cetak_4A2.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "5A1" Then
    Cetak_5A1.lblkdgudang = rs!kdgudang
    Cetak_5A1.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "5A2" Then
    Cetak_5A2.lblkdgudang = rs!kdgudang
    Cetak_5A2.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "5A3" Then
    Cetak_5A3.lblkdgudang = rs!kdgudang
    Cetak_5A3.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "5A4" Then
    Cetak_5A4.lblkdgudang = rs!kdgudang
    Cetak_5A4.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "6A1" Then
    Cetak_6A1.lblkdgudang = rs!kdgudang
    Cetak_6A1.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "6A2" Then
    Cetak_6A2.lblkdgudang = rs!kdgudang
    Cetak_6A2.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "3A5" Then
    Cetak_3A5.lblkdgudang = rs!kdgudang
    Cetak_3A5.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "1A3" Then
    Cetak_1A3.lblkdgudang = rs!kdgudang
    Cetak_1A3.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "1A4A1" Then
    Cetak_1A4A1.lblkdgudang = rs!kdgudang
    Cetak_1A4A1.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "1A5" Then
    Cetak_1A5.lblkdgudang = rs!kdgudang
    Cetak_1A5.lblnmgudang = rs!nmgudang
    ElseIf LBLKODE = "S_OPNAME" Then
    S_OPname.lblkdgudang = rs!kdgudang
    S_OPname.lblnmgudang = rs!nmgudang

    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call ALL
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerALL.Interval = 10
End Sub




Private Sub TimerALL_Timer()
On Error Resume Next
Call ALL

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





