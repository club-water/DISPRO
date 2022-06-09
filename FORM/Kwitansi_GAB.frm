VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form Kwitansi_GAB 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttahun2 
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
      Left            =   5715
      TabIndex        =   4
      Text            =   "2020"
      Top             =   1845
      Width           =   780
   End
   Begin VB.ComboBox cmbbln2 
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
      Left            =   4005
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1845
      Width           =   1635
   End
   Begin VB.TextBox txttahun1 
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
      Left            =   2790
      TabIndex        =   2
      Text            =   "2020"
      Top             =   1845
      Width           =   780
   End
   Begin VB.ComboBox CMBbln1 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1845
      Width           =   1635
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   8
      Top             =   720
      Width           =   7800
      _Version        =   524288
      _ExtentX        =   13758
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   450
      TabIndex        =   7
      Top             =   5850
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
      Picture         =   "Kwitansi_GAB.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   7425
      TabIndex        =   0
      Top             =   990
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
      Picture         =   "Kwitansi_GAB.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   3210
      Left            =   90
      TabIndex        =   6
      Top             =   2385
      Width           =   7845
      _cx             =   13838
      _cy             =   5662
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Kwitansi_GAB.frx":9094
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
      Begin VB.Timer TimerALL 
         Left            =   3105
         Top             =   1485
      End
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   0
      Left            =   8100
      TabIndex        =   5
      ToolTipText     =   "Cetak"
      Top             =   2385
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1455
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
      Picture         =   "Kwitansi_GAB.frx":917E
      ButtonStyle     =   4
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   6120
      TabIndex        =   19
      Top             =   5670
      Width           =   1410
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   5085
      TabIndex        =   18
      Top             =   5670
      Width           =   960
   End
   Begin VB.Label lbltgl2 
      Caption         =   "Label6"
      Height          =   375
      Left            =   5625
      TabIndex        =   17
      Top             =   6435
      Width           =   1005
   End
   Begin VB.Label lbltgl1 
      Caption         =   "Label5"
      Height          =   375
      Left            =   4365
      TabIndex        =   16
      Top             =   6435
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   1890
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BULAN :"
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
      TabIndex        =   14
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label lblalamat 
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
      Left            =   1080
      TabIndex        =   13
      Top             =   1395
      Width           =   6855
   End
   Begin VB.Label lblkdcustomer 
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
      Left            =   1080
      TabIndex        =   12
      Top             =   1035
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER :"
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
      Left            =   45
      TabIndex        =   11
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label lblnmcustomer 
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
      Left            =   2250
      TabIndex        =   10
      Top             =   1035
      Width           =   5190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kwitansi Gabungan"
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
      Left            =   495
      TabIndex        =   9
      Top             =   0
      Width           =   5370
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   8100
      Picture         =   "Kwitansi_GAB.frx":CBDB
      Stretch         =   -1  'True
      Top             =   135
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   6315
      Left            =   0
      Picture         =   "Kwitansi_GAB.frx":CF9B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8970
   End
End
Attribute VB_Name = "Kwitansi_GAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsC1 As ADODB.Recordset
Dim rsC2 As ADODB.Recordset
Dim rsT As ADODB.Recordset
Dim rsX As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim color As Long, flag As Byte


Private Sub Cetak()


Unload AR_KWITANSI_GAB

 ms = MsgBox("Cetak Dengan Stempel TSP ?", vbYesNo + vbQuestion, "Info")
    If ms = vbNo Then
        AR_KWITANSI_GAB.IMG_STEMPEL.Visible = False
        AR_KWITANSI_GAB.lbltgl_STEMPEL.Visible = False
        AR_KWITANSI_GAB.Image2.Visible = True
    Else
        AR_KWITANSI_GAB.IMG_STEMPEL.Visible = True
        AR_KWITANSI_GAB.lbltgl_STEMPEL.Visible = True
        AR_KWITANSI_GAB.Image2.Visible = False
    End If
    
    

sqlX1 = "select kdcustomer,sum(jmlpiutang) as jmlpiutang from piutangsewa where tglposting between '" & Format(lbltgl1, "yyyy/MM/dd") & "' and '" & Format(lbltgl2, "yyyy/MM/dd") & "' and kdcustomer ='" & lblkdcustomer & "'  group by kdcustomer"

sqlX = "select a.kdcustomer,b.nmcustomer,b.alamat,a.jmlpiutang,c.nmbank,c.norek,c.atas_nama from (" & sqlX1 & ") a " & vbCrLf & _
       "left join customer b on a.kdcustomer=b.kdcustomer left join bank c on b.kdbank=c.kdbank "
       
Set rsX = con.Execute(sqlX)

With AR_KWITANSI_GAB.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_KWITANSI_GAB
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.flduang.DataField = "jmlpiutang"
.fldjmlpiutang.DataField = "jmlpiutang"
.fldtglposting = Format(Date, "dd/MM/yyyy")
.fldnorek.DataField = "norek"
.fldnmbank.DataField = "nmbank"
.fldAtas_nama.DataField = "atas_nama"

Select Case CLng(Month(Date))
       Case 1
      .lbltgl_STEMPEL = Day(Date) & " JAN " & Year(Date)
       Case 2
      .lbltgl_STEMPEL = Day(Date) & " FEB " & Year(Date)
       Case 3
      .lbltgl_STEMPEL = Day(Date) & " MAR " & Year(Date)
       Case 4
      .lbltgl_STEMPEL = Day(Date) & " APR " & Year(Date)
       Case 5
      .lbltgl_STEMPEL = Day(Date) & " MEI " & Year(Date)
       Case 6
      .lbltgl_STEMPEL = Day(Date) & " JUN " & Year(Date)
       Case 7
      .lbltgl_STEMPEL = Day(Date) & " JUL " & Year(Date)
       Case 8
      .lbltgl_STEMPEL = Day(Date) & " AGS " & Year(Date)
       Case 9
      .lbltgl_STEMPEL = Day(Date) & " SEP " & Year(Date)
       Case 10
      .lbltgl_STEMPEL = Day(Date) & " OKT " & Year(Date)
       Case 11
      .lbltgl_STEMPEL = Day(Date) & " NOV " & Year(Date)
       Case 12
      .lbltgl_STEMPEL = Day(Date) & " DES " & Year(Date)

End Select

.Zoom = 140

AR_KWITANSI_GAB.Show vbModal


End With

End Sub




Private Sub ALL()

sql1 = "select '1' as kode,kdpiutang,bln,tahun,unit,harga,jmlpiutang from piutangsewa where tglposting between '" & Format(lbltgl1, "yyyy/MM/dd") & "' and '" & Format(lbltgl2, "yyyy/MM/dd") & "' and kdcustomer ='" & lblkdcustomer & "' "

sql = sql1 & "order by tahun,bln"

sqlT = "select kode, sum(jmlpiutang) as jmlpiutang from (" & sql1 & ") a group by kode"



Set rs = con.Execute(sql)
Set rsT = con.Execute(sqlT)
Set datagrid1.DataSource = rs


If rsT.RecordCount <> 0 Then
lblTotal = FormatNumber(rsT!jmlpiutang, 0)
Else
lblTotal = 0
End If

If rs.RecordCount = 0 Then
cmdT(0).Enabled = False
Else
cmdT(0).Enabled = True
End If



End Sub

Private Sub CMBbln1_Click()
On Error GoTo hell
sqlC1 = "select * from piutangsewa where bln=" & CLng(CMBbln1.ListIndex) + 1 & " and tahun=" & txttahun1 & " and kdcustomer='" & lblkdcustomer & "'"

Set rsC1 = con.Execute(sqlC1)

If rsC1.RecordCount <> 0 Then
lbltgl1 = rsC1!tglposting
Else
lbltgl1 = "01/01/3000"
End If

TimerALL.Interval = 10

Exit Sub
hell:
lbltgl1 = "01/01/3000"
TimerALL.Interval = 10

End Sub

Private Sub CMBbln1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmbbln2_Click()
On Error GoTo hell
sqlC2 = "select * from piutangsewa where bln=" & CLng(cmbbln2.ListIndex) + 1 & " and tahun=" & txttahun2 & " and kdcustomer='" & lblkdcustomer & "'"

Set rsC2 = con.Execute(sqlC2)

If rsC2.RecordCount <> 0 Then
lbltgl2 = rsC2!tglposting
Else
lbltgl2 = "01/01/1900"
End If

TimerALL.Interval = 10

Exit Sub
hell:
lbltgl2 = "01/01/1900"
TimerALL.Interval = 10
End Sub

Private Sub cmbbln2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab

End Sub

Private Sub cmdBR_Click()
Customer_br.LBLKODE = "KWITANSI_GAB"
Customer_br.Show vbModal
End Sub


Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then

Call Cetak
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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("p") Or KeyAscii = Asc("P") Then
    If rs.RecordCount <> 0 Then Call Cetak
End If

End Sub

Private Sub Form_Load()
GradientForm Me, 0


CMBbln1.AddItem "Januari"
CMBbln1.AddItem "Februari"
CMBbln1.AddItem "Maret"
CMBbln1.AddItem "April"
CMBbln1.AddItem "Mei"
CMBbln1.AddItem "Juni"
CMBbln1.AddItem "Juli"
CMBbln1.AddItem "Agustus"
CMBbln1.AddItem "September"
CMBbln1.AddItem "Oktober"
CMBbln1.AddItem "November"
CMBbln1.AddItem "Desember"

CMBbln1.ListIndex = Month(Now) - 1
txttahun1 = Year(Now)



cmbbln2.AddItem "Januari"
cmbbln2.AddItem "Februari"
cmbbln2.AddItem "Maret"
cmbbln2.AddItem "April"
cmbbln2.AddItem "Mei"
cmbbln2.AddItem "Juni"
cmbbln2.AddItem "Juli"
cmbbln2.AddItem "Agustus"
cmbbln2.AddItem "September"
cmbbln2.AddItem "Oktober"
cmbbln2.AddItem "November"
cmbbln2.AddItem "Desember"

cmbbln2.ListIndex = Month(Now) - 1
txttahun2 = Year(Now)




End Sub

Private Sub lblkdcustomer_Change()
CMBbln1_Click
cmbbln2_Click

'TimerALL.Interval = 10

End Sub

Private Sub TimerALL_Timer()
On Error GoTo hell

Call ALL

TimerALL.Interval = 0

Exit Sub
hell:
MsgBox err.Description
TimerALL.Interval = 0
End Sub

Private Sub txttahun1_Change()
Call nul(txttahun1)
CMBbln1_Click
End Sub

Private Sub txttahun1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttahun1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttahun1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then

    cekTBL = InStr("1234567890", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub


Private Sub txttahun2_Change()
Call nul(txttahun2)
cmbbln2_Click
End Sub

Private Sub txttahun2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttahun2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttahun2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
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


