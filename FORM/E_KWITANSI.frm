VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form E_KWITANSI 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10140
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerEmail 
      Left            =   8100
      Top             =   675
   End
   Begin VB.Timer TimerPDF 
      Left            =   7650
      Top             =   675
   End
   Begin VB.Timer Timercetak 
      Left            =   7200
      Top             =   675
   End
   Begin VB.Timer TimerAll 
      Left            =   6750
      Top             =   675
   End
   Begin VB.TextBox txttahun 
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
      Left            =   3780
      TabIndex        =   2
      Text            =   "2017"
      Top             =   990
      Width           =   960
   End
   Begin VB.ComboBox CMBBLN 
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
      Left            =   1935
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   915
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   2715
      Left            =   135
      TabIndex        =   0
      Top             =   1485
      Width           =   17610
      _cx             =   31062
      _cy             =   4789
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
      GridColor       =   -2147483638
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
      FormatString    =   $"E_KWITANSI.frx":0000
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
      FrozenCols      =   6
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   4680
      Left            =   135
      TabIndex        =   5
      Top             =   4230
      Width           =   17595
      _ExtentX        =   31036
      _ExtentY        =   8255
      SectionData     =   "E_KWITANSI.frx":01AA
   End
   Begin Threed.SSCommand cmdPDF 
      Height          =   825
      Left            =   17910
      TabIndex        =   6
      ToolTipText     =   "Create PDF"
      Top             =   1485
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
      Picture         =   "E_KWITANSI.frx":01E6
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdEmail 
      Height          =   825
      Left            =   17910
      TabIndex        =   9
      ToolTipText     =   "Ubah"
      Top             =   2340
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
      Picture         =   "E_KWITANSI.frx":33CD
      ButtonStyle     =   4
   End
   Begin C1SizerLibCtl.C1Elastic flood 
      Height          =   420
      Left            =   315
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8955
      Visible         =   0   'False
      Width           =   17400
      _cx             =   30692
      _cy             =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      BackColor       =   16777215
      ForeColor       =   0
      FloodColor      =   16776960
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
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
      FrameStyle      =   2
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
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   9585
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
      Picture         =   "E_KWITANSI.frx":74DF
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   315
      TabIndex        =   15
      Top             =   765
      Width           =   17430
      _Version        =   524288
      _ExtentX        =   30745
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   17865
      Picture         =   "E_KWITANSI.frx":DD41
      Stretch         =   -1  'True
      Top             =   450
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E - KWITANSI"
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
      Left            =   1440
      TabIndex        =   16
      Top             =   45
      Width           =   5685
   End
   Begin VB.Label lblnm_pengirim 
      Caption         =   "Label1"
      Height          =   420
      Left            =   11340
      TabIndex        =   12
      Top             =   990
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblpass 
      Caption         =   "Label1"
      Height          =   420
      Left            =   10845
      TabIndex        =   11
      Top             =   990
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblemail_pengirim 
      Caption         =   "Label1"
      Height          =   420
      Left            =   10350
      TabIndex        =   10
      Top             =   990
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   420
      Left            =   9855
      TabIndex        =   8
      Top             =   990
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lbljml 
      Caption         =   "lbljml"
      Height          =   420
      Left            =   9360
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3105
      TabIndex        =   4
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TAGIHAN BLN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   675
      TabIndex        =   3
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   0
      Picture         =   "E_KWITANSI.frx":E101
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18870
   End
End
Attribute VB_Name = "E_KWITANSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsT As ADODB.Recordset
Dim rsX As ADODB.Recordset
Dim rsE As ADODB.Recordset
Dim nm_file As String
Dim ket_bln As String

Dim color As Long, flag As Byte

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Sub Kirim_Email()
On Error GoTo hell


Select Case CLng(CMBBLN.Text)
       Case 1
      ket_bln = "JAN " & txttahun
       Case 2
      ket_bln = "FEB " & txttahun
       Case 3
      ket_bln = "MAR " & txttahun
       Case 4
      ket_bln = "APR " & txttahun
       Case 5
      ket_bln = "MEI " & txttahun
       Case 6
      ket_bln = "JUN " & txttahun
       Case 7
      ket_bln = "JUL " & txttahun
       Case 8
      ket_bln = "AGS " & txttahun
       Case 9
      ket_bln = "SEP " & txttahun
       Case 10
      ket_bln = "OKT " & txttahun
       Case 11
      ket_bln = "NOV " & txttahun
       Case 12
      ket_bln = "DES " & txttahun

End Select



Set mg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set fn = iConf.Fields
schema = "http://schemas.microsoft.com/cdo/configuration/"
fn.Item(schema & "sendusing") = 2
fn.Item(schema & "smtpserver") = "smtp.gmail.com"
fn.Item(schema & "smtpserverport") = 465
fn.Item(schema & "smtpauthenticate") = 1
fn.Item(schema & "sendusername") = "" & lblemail_pengirim & ""
fn.Item(schema & "sendpassword") = "" & lblpass & ""
fn.Item(schema & "smtpusessl") = 1
fn.Update



With mg

.To = "" & rs!email_penerima & ""
.From = "" & lblnm_pengirim & " <" & lblemail_pengirim & ">"
'.CC = "" & txtCC_penerima & ""
.Subject = "Tagihan Sewa Dispencer ( " & rs!kdcustomer & " ) " & rs!nmcustomer & " --- " & ket_bln & " "
.HTMLBody = "Berikut kami lampirkan Tagihan Sewa Dispencer ( " & rs!kdcustomer & " ) " & rs!nmcustomer & " ---- Periode : " & ket_bln & " .......Don't Replay"
.Sender = "" & lblnm_pengirim & ""
.Organization = "TSP NGANJUK"
.AddAttachment "" & App.Path & "\EMAIL\" & rs!kdcustomer & "_" & CMBBLN.Text & "_" & Right(txttahun, 2) & ".pdf" & ""
'.ReplyTo = "" & lblnmpengirim & ""



Set .Configuration = iConf
SendEmailGmail = .Send
End With

Set mg = Nothing
Set iConf = Nothing
Set fn = Nothing
Set fn = Nothing

Exit Sub
hell:
MsgBox err.Description, vbInformation, "Error !"

End

End Sub



Private Sub Cetak()
Unload AR_Kwitansi

AR_Kwitansi.IMG_STEMPEL.Visible = True
AR_Kwitansi.lbltgl_STEMPEL.Visible = True
AR_Kwitansi.Image2.Visible = False


nm_file = rs!kdcustomer & "_" & CMBBLN.Text & "_" & Right(txttahun, 2)



sqlX = "select a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,c.nmbank,c.norek,a.tglcetak,c.atas_nama from piutangsewa a " & vbCrLf & _
      "left join customer b on a.kdcustomer=b.kdcustomer left join bank c on b.kdbank=c.kdbank where a.bln=" & CMBBLN.ListIndex + 1 & " and a.tahun=" & txttahun & " and a.kdcustomer='" & rs!kdcustomer & "' order by a.kdpiutang,a.kdcustomer"

Set rsX = con.Execute(sqlX)

With AR_Kwitansi.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_Kwitansi
.fldnokwitansi.DataField = "kdpiutang"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.flduang.DataField = "jmlpiutang"
.fldbln.DataField = "bln"
.FLDTHN.DataField = "tahun"
.fldunit.DataField = "unit"
.fldharga.DataField = "harga"
.fldjmlpiutang.DataField = "jmlpiutang"
.fldtglposting.DataField = "tglcetak"
.fldnorek.DataField = "norek"
.fldnmbank.DataField = "nmbank"
.fldAtas_nama.DataField = "atas_nama"

Select Case CLng(Month(rsX!tglcetak))
       Case 1
      .lbltgl_STEMPEL = "23 JAN " & Kwitansi.txttahun
       Case 2
      .lbltgl_STEMPEL = "23 FEB " & Kwitansi.txttahun
       Case 3
      .lbltgl_STEMPEL = "23 MAR " & Kwitansi.txttahun
       Case 4
      .lbltgl_STEMPEL = "23 APR " & Kwitansi.txttahun
       Case 5
      .lbltgl_STEMPEL = "23 MEI " & Kwitansi.txttahun
       Case 6
      .lbltgl_STEMPEL = "23 JUN " & Kwitansi.txttahun
       Case 7
      .lbltgl_STEMPEL = "23 JUL " & Kwitansi.txttahun
       Case 8
      .lbltgl_STEMPEL = "23 AGS " & Kwitansi.txttahun
       Case 9
      .lbltgl_STEMPEL = "23 SEP " & Kwitansi.txttahun
       Case 10
      .lbltgl_STEMPEL = "23 OKT " & Kwitansi.txttahun
       Case 11
      .lbltgl_STEMPEL = "23 NOV " & Kwitansi.txttahun
       Case 12
      .lbltgl_STEMPEL = "23 DES " & Kwitansi.txttahun

End Select

If rs!st_pdf <> "X" Then
con.Execute ("insert into status_E_kwitansi values ('" & rsX!kdpiutang & "',1,0,getdate(),'" & UTAMA.lblkduser & "')")
End If


Set Me.ARV1.ReportSource = AR_Kwitansi




If CLng(lblpos) < CLng(lbljml) Then
rs.MoveNext
lblpos = lblpos + 1
End If




End With

End Sub




Private Sub ALL()

sql1 = "select  '1' as kode,a.kdpiutang,a.bln,a.tahun,a.kdcustomer,b.nmcustomer,b.alamat,a.unit,a.harga,jmlpiutang,a.tglposting,a.tglcetak,isnull(c.email_penerima,'') as email_penerima, (case when isnull(st_pdf,0)=1 then 'X' else '' end) as ST_PDF ,(case when isnull(st_email,0)=1 then 'X' else '' end) as ST_email   from piutangsewa a " & vbCrLf & _
       "left join customer b on a.kdcustomer=b.kdcustomer left join list_email_cust c on a.kdcustomer=c.kdcustomer left join status_e_kwitansi d on a.kdpiutang=d.kdpiutang" & vbCrLf & _
       "where a.bln=" & CMBBLN.ListIndex + 1 & "  and a.tahun=" & txttahun & " and a.kdcustomer in (select kdcustomer from list_email_cust where non_aktif=0)"
      
sql = sql1 & "order by a.kdpiutang,a.kdcustomer"



Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

sqlT = "select kode,sum(convert(int,kode)) as jml from (" & sql1 & ") a group by kode"
Set rsT = con.Execute(sqlT)
       
       
If rsT.RecordCount <> 0 Then
lbljml = rsT!jml
Else
lbljml = 0
End If


End Sub


Private Sub ARV1_LoadCompleted()
If CLng(lblpos) < CLng(lbljml) Then
TimerCetak.Interval = 1000
End If

On Error Resume Next
If rs!st_pdf = "X" Then
    ms = MsgBox("Data sudah pernah di Create PDF, apa ingin Create Ulang ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        TimerPdf.Interval = 10
    End If
Else
    TimerPdf.Interval = 10
End If


flood.FloodPercent = (CLng(lblpos) / CLng(lbljml)) * 100
flood.Caption = "Proses Create PDF : " & FormatNumber((CLng(lblpos) / CLng(lbljml)) * 100, 1) & "%"

If flood.FloodPercent = 100 Then
MsgBox "Create PDF berhasil", vbInformation, "Info !"
TimerPdf.Interval = 10
flood.Visible = False
flood.FloodPercent = 0
Call ALL
End If

End Sub

Private Sub CMBBLN_Click()
TimerALL.Interval = 10
End Sub

Private Sub CMBBLN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub cmdEmail_Click()

sqlE = "select * from email_pengirim where kdpengirim='A'"
Set rsE = con.Execute(sqlE)

If rsE.RecordCount <> 0 Then
lblemail_pengirim = rsE!alamat_email
lblpass = rsE!pass_pengirim
lblnm_pengirim = rsE!nm_pengirim
Else
lblemail_pengirim = ""
lblpass = ""
lblnm_pengirim = ""
End If

Call ALL

flood.Visible = True

TimerEmail.Interval = 10


End Sub

Private Sub cmdPDF_Click()
flood.Visible = True
rs.MoveFirst
lblpos = 0
Call Cetak
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If

End Sub

Private Sub Form_Load()

GradientForm Me, 0

CMBBLN.AddItem "1"
CMBBLN.AddItem "2"
CMBBLN.AddItem "3"
CMBBLN.AddItem "4"
CMBBLN.AddItem "5"
CMBBLN.AddItem "6"
CMBBLN.AddItem "7"
CMBBLN.AddItem "8"
CMBBLN.AddItem "9"
CMBBLN.AddItem "10"
CMBBLN.AddItem "11"
CMBBLN.AddItem "12"
CMBBLN.ListIndex = Month(Date) - 1

txttahun = Year(Date)

TimerALL.Interval = 10
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub TimerALL_Timer()
On Error GoTo hell
Call ALL

TimerALL.Interval = 0

Exit Sub
hell:
TimerALL.Interval = 0
MsgBox err.Description

End Sub

Private Sub TimerCetak_Timer()
Static Z As Integer

Z = Z + 1

If Z = 3 Then
Call Cetak
Z = 0
TimerCetak.Interval = 0
End If

End Sub

Private Sub TimerEmail_Timer()
On Error Resume Next

Static Y As Integer

Y = Y + 1

If Y <= CLng(lbljml) Then

    flood.FloodPercent = (Y / CLng(lbljml)) * 100
    flood.Caption = "Kirim Email : " & FormatNumber((Y / CLng(lbljml)) * 100, 1) & "%"
    
   
    
    rs.AbsolutePosition = Y
    
    If rs!st_email = "X" Then
    ms = MsgBox("Data sudah pernah diemail, apa ingin email Ulang ?", vbYesNo + vbQuestion, "Info")
        If ms = vbYes Then
            Call Kirim_Email
        End If
    Else
    Call Kirim_Email
    End If

Else
    Y = 0
    TimerEmail.Interval = 0
    MsgBox "Kwitansi Berhasil di Email", vbInformation, "Info !!"
    flood.Visible = False
    flood.FloodPercent = 0
    
End If


If rs!st_email <> "X" Then
    con.Execute ("update status_E_kwitansi set st_email = 1,tglkirim=getdate() where kdpiutang='" & rs!kdpiutang & "'")
End If

End Sub

Private Sub TimerPDF_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF



pdf.filename = App.Path & "\EMAIL\" & nm_file & ".pdf"
pdf.Export AR_Kwitansi.Pages

TimerPdf.Interval = 0

Exit Sub
hell:
MsgBox err.Description
TimerPdf.Interval = 0

End Sub

Private Sub txttahun_Change()
TimerALL.Interval = 10
End Sub

Private Sub txttahun_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then

    cekTBL = InStr("1234567890.,", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If

End Sub

Private Sub txtharga_LostFocus()
On Error GoTo hell

txttahun = Format(txttahun, "####0")


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txttahun.SetFocus

End Sub

