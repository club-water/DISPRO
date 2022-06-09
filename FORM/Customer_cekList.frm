VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Customer_cekList 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6840
      TabIndex        =   2
      Text            =   "5"
      Top             =   900
      Width           =   735
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
      TabIndex        =   1
      Top             =   900
      Width           =   2490
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   3
      Top             =   810
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
      TabIndex        =   4
      Top             =   8730
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
      Picture         =   "Customer_cekList.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   6900
      Left            =   135
      TabIndex        =   0
      Top             =   1305
      Width           =   15945
      _cx             =   28125
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Customer_cekList.frx":6862
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
      Caption         =   "TOTAL UNIT LEBIH DARI :"
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
      Left            =   4770
      TabIndex        =   11
      Top             =   945
      Width           =   2085
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   270
      TabIndex        =   10
      Top             =   9315
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
      TabIndex        =   9
      Top             =   900
      Width           =   1500
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      TabIndex        =   8
      Top             =   180
      Width           =   5280
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   16245
      Picture         =   "Customer_cekList.frx":69B3
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3960
      Picture         =   "Customer_cekList.frx":6D73
      Stretch         =   -1  'True
      Top             =   855
      Width           =   420
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
      Left            =   12870
      TabIndex        =   7
      Top             =   945
      Width           =   1815
   End
   Begin VB.Label lbltgl1 
      BackStyle       =   0  'Transparent
      Caption         =   "20/12/20109"
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
      Left            =   14580
      TabIndex        =   6
      Top             =   945
      Width           =   1500
   End
   Begin VB.Label lblpos 
      Caption         =   "1"
      Height          =   330
      Left            =   7695
      TabIndex        =   5
      Top             =   8550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   9240
      Left            =   0
      Picture         =   "Customer_cekList.frx":13C23
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17115
   End
End
Attribute VB_Name = "Customer_cekList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim sql, sql1, sql2, sqlX As String
Dim rs1 As ADODB.Recordset




Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub datagrid1_Click()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub

Private Sub datagrid1_GotFocus()
DataGrid1.HighLight = flexHighlightAlways
End Sub

Private Sub datagrid1_LostFocus()
DataGrid1.HighLight = flexHighlightNever
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub ALL()
MousePointer = vbHourglass

sql1 = "select kdcustomer,(disp1 + disp2 + disp3 +disp4) as disp , (show1 + show2) as showC,RG from ( " & vbCrLf & _
            "select kdcustomer, SUM(case kdkategori when '04' then unit else 0 end) as disp1, SUM(case kdkategori when '05' then unit else 0 end) as disp2," & vbCrLf & _
            "SUM(case kdkategori when '06' then unit else 0 end) as disp3, SUM(case kdkategori when '07' then unit else 0 end) as disp4,SUM(case kdkategori when '08' then unit else 0 end) as show1," & vbCrLf & _
            "SUM(case kdkategori when '09' then unit else 0 end) as show2,SUM(case kdkategori when '10' then unit else 0 end) as RG from (" & vbCrLf & _
                "select kdcustomer,kdkategori,sum(unit-Runit)as unit from V_brg_split where tgl  <= '" & Format(lbltgl1, "yyyy/MM/dd") & "' group by kdcustomer,kdkategori" & vbCrLf & _
            ") a group by kdcustomer " & vbCrLf & _
       ") a where disp1 + disp2 + disp3 +disp4 + show1 + show2+RG <>0"


sql2 = "select d.nmareaC,e.nmteknisi,a.kdcustomer,b.nmcustomer,b.alamat,b.cp,b.telp,a.disp,a.showC,a.RG from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
       "left join  area_cheker d on b.kdareaC=d.kdareaC left join teknisi e on b.kdteknisi= e.kdteknisi  "
      
      
If txtcari = "" Then
sql = "select nmareaC,nmteknisi,kdcustomer,nmcustomer,alamat,cp,telp,disp,showC,RG,disp+showC+RG as total from (" & sql2 & ") a where disp+showC+RG > " & CLng(txtR) & " order by nmareaC,nmteknisi,nmcustomer,alamat"
Else
sql = "select nmareaC,nmteknisi,kdcustomer,nmcustomer,alamat,cp,telp,disp,showC,RG,disp+showC+RG as total from (" & sql2 & ") a where disp+showC+RG > " & CLng(txtR) & " and (kdcustomer like '%" & txtcari & "%' or nmcustomer like '%" & txtcari & "%' or alamat like '%" & txtcari & "%' or nmareaC like '%" & txtcari & "%' or nmteknisi like '%" & txtcari & "%')  order by nmareaC,nmteknisi,nmcustomer,alamat"
End If


Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs

For i = 1 To (DataGrid1.Rows - 1)

DataGrid1.TextMatrix(i, 0) = i
DataGrid1.Cell(flexcpForeColor, i, 11) = vbRed
DataGrid1.Cell(flexcpBackColor, i, 11) = vbGreen
DataGrid1.Cell(flexcpFontSize, i, 11) = 8
Next


MousePointer = vbDefault

End Sub



Private Sub DataGrid1_DblClick()
On Error GoTo hell

Cetak_ceklist.lblkdcustomer = rs!kdcustomer
Cetak_ceklist.lblnmcustomer = rs!nmcustomer
Cetak_ceklist.lblalamat = rs!alamat

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
    
    Cetak_ceklist.lblkdcustomer = rs!kdcustomer
    Cetak_ceklist.lblnmcustomer = rs!nmcustomer
    Cetak_ceklist.lblalamat = rs!alamat
    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

LBLKODE = 1


TimerALL.Interval = 10
End Sub






Private Sub lblkdareaC_Click()

End Sub

Private Sub TimerALL_Timer()
On Error Resume Next
Call ALL

If LBLKODE = 2 Then
rs.AbsolutePosition = lblpos
End If


TimerALL.Interval = 0
End Sub


Private Sub TimerG_Timer()

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
    
'    Else
'    CMBCARI.SetFocus
    End If
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    DataGrid1.SetFocus
    
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


