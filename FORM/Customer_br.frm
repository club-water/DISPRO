VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Customer_br 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkR 
      BackColor       =   &H00000000&
      Caption         =   "TAMPILKAN :"
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
      Height          =   330
      Left            =   11340
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   1395
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.TextBox txtR 
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
      Left            =   12915
      TabIndex        =   3
      Text            =   "100"
      Top             =   1395
      Width           =   735
   End
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
      Left            =   270
      TabIndex        =   0
      Top             =   1350
      Width           =   2490
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   6
      Top             =   855
      Width           =   14685
      _Version        =   524288
      _ExtentX        =   25903
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
      TabIndex        =   5
      Top             =   9045
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
      Picture         =   "Customer_br.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   2745
      TabIndex        =   1
      ToolTipText     =   "Tambah Customer Baru"
      Top             =   1350
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
      Picture         =   "Customer_br.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7215
      Left            =   270
      TabIndex        =   4
      Top             =   1755
      Width           =   14595
      _cx             =   25744
      _cy             =   12726
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
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777152
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Customer_br.frx":902B
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RECORD"
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
      Left            =   13680
      TabIndex        =   10
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   135
      TabIndex        =   9
      Top             =   9720
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
      Left            =   315
      TabIndex        =   8
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      TabIndex        =   7
      Top             =   135
      Width           =   2715
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   14940
      Picture         =   "Customer_br.frx":90AA
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1440
      Picture         =   "Customer_br.frx":946A
      Stretch         =   -1  'True
      Top             =   990
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   9690
      Left            =   0
      Picture         =   "Customer_br.frx":1631A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15810
   End
End
Attribute VB_Name = "Customer_br"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte

Private Sub ChkR_Click()
TimerALL.Interval = 10

If ChkR.Value = 0 Then
txtR.Enabled = False
Else
txtR.Enabled = True
End If

End Sub

Private Sub ChkR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdBR_Click()
Customer_TU.LBLKODE = 1
Customer_TU.lblfrm = "CUSTOMER_BR"
Customer_TU.Show vbModal
End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If

End Sub

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



Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

sql1 = "select kdpiutang, kdcustomer,sum(jmlpiutang) as jmlpiutang, sum(jmlbayar) as jmlbayar,sum(potongan) as potongan," & vbCrLf & _
        "sum(jmlpiutang - jmlbayar - potongan) as sisa from (" & vbCrLf & _
        "select 'a' as kode,kdpiutang,kdcustomer,jmlpiutang, 0 as jmlbayar,0 as potongan from piutangsewa" & vbCrLf & _
        "Union" & vbCrLf & _
        "select 'b' as kode,kdpiutang,kdcustomer,0 as jmlpiutang,sum(jmlbayar) as jmlbayar,sum(potongan) as potongan  from byrpiutangsewa" & vbCrLf & _
        "group by kdpiutang,kdcustomer ) a group by kdpiutang, kdcustomer"


sqlA = "select a.kdcustomer from (" & sql1 & ") a " & vbCrLf & _
       "left join piutangsewa c on a.kdpiutang=c.kdpiutang where a.sisa<>0 and c.tt=0"


If LBLKODE = "TTERIMA_D" Then
    If TXTCARI = "" Then
    sql = "select kdcustomer,nmcustomer,alamat from customer where kdcustomer in (" & sqlA & " ) order by kdcustomer"
    Else
    sql = "select kdcustomer,nmcustomer,alamat from customer where kdcustomer in (" & sqlA & " ) and (kdcustomer like '%" & TXTCARI & "%' or nmcustomer like '%" & TXTCARI & "%' or alamat like '%" & TXTCARI & "%') order by kdcustomer"
    End If
ElseIf LBLKODE = "TEKNISILUAR_D" Then
    If TXTCARI = "" Then
    sql = "select kdcustomer,nmcustomer,alamat from customer where kdcustomer in (select kdcustomer from rekap_pjm_sewa) order by kdcustomer"
    Else
    sql = "select kdcustomer,nmcustomer,alamat from customer where kdcustomer in (select kdcustomer from rekap_pjm_sewa) and (kdcustomer like '%" & TXTCARI & "%' or nmcustomer like '%" & TXTCARI & "%' or alamat like '%" & TXTCARI & "%') order by kdcustomer"
    End If
Else
    
    
    If ChkR.Value = 0 Then
        If TXTCARI = "" Then
        sql = "select kdcustomer,nmcustomer,alamat from customer where non_aktif=0 order by kdcustomer desc"
        Else
        sql = "select kdcustomer,nmcustomer,alamat from customer where (kdcustomer like '%" & TXTCARI & "%' or nmcustomer like '%" & TXTCARI & "%' or alamat like '%" & TXTCARI & "%' or noSPK like '%" & TXTCARI & "%') and non_aktif=0  order by kdcustomer desc"
        End If
    Else
        If TXTCARI = "" Then
        sql = "select TOP " & CLng(txtR) & " kdcustomer,nmcustomer,alamat from customer where non_aktif=0 order by kdcustomer desc"
        Else
        sql = "select TOP " & CLng(txtR) & " kdcustomer,nmcustomer,alamat from customer where (kdcustomer like '%" & TXTCARI & "%' or nmcustomer like '%" & TXTCARI & "%' or alamat like '%" & TXTCARI & "%' or noSPK like '%" & TXTCARI & "%') and non_aktif=0  order by kdcustomer desc"
        End If
    End If
    
End If

Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell
If LBLKODE = "PO_D" Then
PO_D.lblkdcustomer = rs!kdcustomer
PO_D.lblnmcustomer = rs!nmcustomer
PO_D.lblalamat = rs!alamat
ElseIf LBLKODE = "3A4" Then
Cetak_3A4.lblkdcustomer = rs!kdcustomer
Cetak_3A4.lblnmcustomer = rs!nmcustomer
Cetak_3A4.lblalamat = rs!alamat
ElseIf LBLKODE = "4A4" Then
Cetak_4A4.lblkdcustomer = rs!kdcustomer
Cetak_4A4.lblnmcustomer = rs!nmcustomer
Cetak_4A4.lblalamat = rs!alamat
ElseIf LBLKODE = "TTERIMA_D" Then
TTerima_D.lblkdcustomer = rs!kdcustomer
TTerima_D.lblnmcustomer = rs!nmcustomer
TTerima_D.lblalamat = rs!alamat
ElseIf LBLKODE = "6A3" Then
Cetak_6A3.lblkdcustomer = rs!kdcustomer
Cetak_6A3.lblnmcustomer = rs!nmcustomer
Cetak_6A3.lblalamat = rs!alamat
ElseIf LBLKODE = "7A1" Then
Cetak_7A1.lblkdcustomer = rs!kdcustomer
Cetak_7A1.lblnmcustomer = rs!nmcustomer
Cetak_7A1.lblalamat = rs!alamat
ElseIf LBLKODE = "7A2" Then
Cetak_7A2.lblkdcustomer = rs!kdcustomer
Cetak_7A2.lblnmcustomer = rs!nmcustomer
Cetak_7A2.lblalamat = rs!alamat

ElseIf LBLKODE = "2A5" Then
Cetak_2A5.lblkdcustomer = rs!kdcustomer
Cetak_2A5.lblnmcustomer = rs!nmcustomer
Cetak_2A5.lblalamat = rs!alamat

ElseIf LBLKODE = "KWITANSI_GAB" Then
Kwitansi_GAB.lblkdcustomer = rs!kdcustomer
Kwitansi_GAB.lblnmcustomer = rs!nmcustomer
Kwitansi_GAB.lblalamat = rs!alamat

ElseIf LBLKODE = "TEKNISILUAR_D" Then
TeknisiLuar_D.lblkdcustomer = rs!kdcustomer
TeknisiLuar_D.lblnmcustomer = rs!nmcustomer
TeknisiLuar_D.lblalamat = rs!alamat

ElseIf LBLKODE = "9A2" Then
Cetak_9A2.lblkdcustomer = rs!kdcustomer
Cetak_9A2.lblnmcustomer = rs!nmcustomer
Cetak_9A2.lblalamat = rs!alamat


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
    
    If LBLKODE = "PO_D" Then
    PO_D.lblkdcustomer = rs!kdcustomer
    PO_D.lblnmcustomer = rs!nmcustomer
    PO_D.lblalamat = rs!alamat
    ElseIf LBLKODE = "3A4" Then
    Cetak_3A4.lblkdcustomer = rs!kdcustomer
    Cetak_3A4.lblnmcustomer = rs!nmcustomer
    Cetak_3A4.lblalamat = rs!alamat
    ElseIf LBLKODE = "4A4" Then
    Cetak_4A4.lblkdcustomer = rs!kdcustomer
    Cetak_4A4.lblnmcustomer = rs!nmcustomer
    Cetak_4A4.lblalamat = rs!alamat
    ElseIf LBLKODE = "TTERIMA_D" Then
    TTerima_D.lblkdcustomer = rs!kdcustomer
    TTerima_D.lblnmcustomer = rs!nmcustomer
    TTerima_D.lblalamat = rs!alamat
    ElseIf LBLKODE = "6A3" Then
    Cetak_6A3.lblkdcustomer = rs!kdcustomer
    Cetak_6A3.lblnmcustomer = rs!nmcustomer
    Cetak_6A3.lblalamat = rs!alamat
    ElseIf LBLKODE = "7A1" Then
    Cetak_7A1.lblkdcustomer = rs!kdcustomer
    Cetak_7A1.lblnmcustomer = rs!nmcustomer
    Cetak_7A1.lblalamat = rs!alamat
    ElseIf LBLKODE = "7A2" Then
    Cetak_7A2.lblkdcustomer = rs!kdcustomer
    Cetak_7A2.lblnmcustomer = rs!nmcustomer
    Cetak_7A2.lblalamat = rs!alamat
    ElseIf LBLKODE = "2A5" Then
    Cetak_2A5.lblkdcustomer = rs!kdcustomer
    Cetak_2A5.lblnmcustomer = rs!nmcustomer
    Cetak_2A5.lblalamat = rs!alamat

    ElseIf LBLKODE = "KWITANSI_GAB" Then
    Kwitansi_GAB.lblkdcustomer = rs!kdcustomer
    Kwitansi_GAB.lblnmcustomer = rs!nmcustomer
    Kwitansi_GAB.lblalamat = rs!alamat
    
    ElseIf LBLKODE = "TEKNISILUAR_D" Then
    TeknisiLuar_D.lblkdcustomer = rs!kdcustomer
    TeknisiLuar_D.lblnmcustomer = rs!nmcustomer
    TeknisiLuar_D.lblalamat = rs!alamat
    
    ElseIf LBLKODE = "9A2" Then
    Cetak_9A2.lblkdcustomer = rs!kdcustomer
    Cetak_9A2.lblnmcustomer = rs!nmcustomer
    Cetak_9A2.lblalamat = rs!alamat


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
    DataGrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    DataGrid1.SetFocus
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



Private Sub txtR_Change()
Call nul(txtR)
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
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

Private Sub txtR_LostFocus()
On Error GoTo hell

txtR = FormatNumber(txtR, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtR.SetFocus

End Sub

