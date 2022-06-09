VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form CustDisp_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   16830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2520
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
      Left            =   225
      TabIndex        =   1
      Top             =   1620
      Width           =   2490
   End
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6135
      Left            =   225
      TabIndex        =   2
      Top             =   1980
      Width           =   15450
      _ExtentX        =   27252
      _ExtentY        =   10821
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   14
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   4
      Top             =   855
      Width           =   15630
      _Version        =   524288
      _ExtentX        =   27570
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1125
      TabIndex        =   3
      Top             =   8370
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
      Picture         =   "CustDisp_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL RENCANA KUNJUNGAN :"
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
      Left            =   225
      TabIndex        =   9
      Top             =   990
      Width           =   2265
   End
   Begin VB.Label lblpos 
      Caption         =   "Label3"
      Height          =   375
      Left            =   2970
      TabIndex        =   8
      Top             =   8865
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1395
      Picture         =   "CustDisp_BR.frx":6862
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   15930
      Picture         =   "CustDisp_BR.frx":13712
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
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
      TabIndex        =   6
      Top             =   1305
      Width           =   1095
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   585
      TabIndex        =   5
      Top             =   8865
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   8835
      Left            =   0
      Picture         =   "CustDisp_BR.frx":13AD2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16800
   End
End
Attribute VB_Name = "CustDisp_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim KODE As Integer


Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
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
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub LG()
On Error GoTo hell

With DataGrid1.Columns(0)
.Width = 60
.Caption = "KODE"
.Alignment = dbgCenter
End With

With DataGrid1.Columns(1)
.Caption = "CUSTOMER"
.Width = 240
End With

With DataGrid1.Columns(2)
.Caption = "ALAMAT"
.Width = 290
End With

With DataGrid1.Columns(3)
.Caption = "KD BARANG"
.Width = 100
End With

With DataGrid1.Columns(4)
.Caption = "BARANG"
.Width = 150
End With

With DataGrid1.Columns(5)
.Caption = "NO DISP"
.Width = 120
End With

With DataGrid1.Columns(6)
.Caption = "JML"
.Width = 30
.Alignment = dbgRight
End With


Exit Sub
hell:

End Sub

Private Sub ALL()
On Error GoTo hell



sql1 = "select a.kdcustomer,c.nmcustomer,c.alamat,a.kdbarang,b.nmbarang,b.kd1 ,sum(uPjm + Usewa - URpinjam - URsewa) as jml from kartu_pelanggan a left join barang b on a.kdbarang=b.kdbarang" & vbCrLf & _
       "left join customer c on a.kdcustomer=c.kdcustomer where convert(int,b.kdkategori) > 3  group by a.kdcustomer,c.nmcustomer,c.alamat,a.kdbarang,b.nmbarang,b.kd1"


If txtcari = "" Then
    sql2 = "select * from (" & sql1 & ") a where jml<>0 "
Else
    sql2 = "select * from (" & sql1 & ") a where jml<>0 and (kdcustomer like '%" & txtcari & "%' or nmcustomer like '%" & txtcari & "%' or alamat like '%" & txtcari & "%' or kdbarang like '%" & txtcari & "%' or nmbarang like '%" & txtcari & "%' or kd1 like '%" & txtcari & "%') "
End If


sql = "select * from (" & sql2 & ") a where kdbarang not in (select kdbarang from fixrute where periode='" & fixrute_TU.txtperiode & "') order by nmcustomer,kdcustomer"

Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub DataGrid1_DblClick()
On Error GoTo hell

KODE = 2
Call max

sqlA1 = "select a.kdcustomer,c.nmcustomer,c.alamat,a.kdbarang,b.nmbarang,b.kd1 ,sum(uPjm + Usewa - URpinjam - URsewa) as jml from kartu_pelanggan a left join barang b on a.kdbarang=b.kdbarang" & vbCrLf & _
       "left join customer c on a.kdcustomer=c.kdcustomer where convert(int,b.kdkategori) > 3  group by a.kdcustomer,c.nmcustomer,c.alamat,a.kdbarang,b.nmbarang,b.kd1"



sqlA = "select * from (" & sqlA1 & ") a where jml<>0 "

sqlX = "insert into fixrute select kdbarang + '_' + '" & fixrute_TU.txtperiode & "','" & fixrute_TU.txtperiode & "','','" & Format(txttgl1, "yyyy/MM/dd") & "','" & fixrute_TU.lblkdteknisi & "',kdcustomer,kdbarang,'','1900/01/01','',0,0,'A01' from (" & sqlA & ") a where kdcustomer='" & rs!kdcustomer & "'"
con.Execute (sqlX)

TimerALL.Interval = 10
fixrute_TU.TimerALL.Interval = 10
fixrute.TimerALL.Interval = 10

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
    
KODE = 2
Call max

sqlA1 = "select a.kdcustomer,c.nmcustomer,c.alamat,a.kdbarang,b.nmbarang,b.kd1 ,sum(uPjm + Usewa - URpinjam - URsewa) as jml from kartu_pelanggan a left join barang b on a.kdbarang=b.kdbarang" & vbCrLf & _
       "left join customer c on a.kdcustomer=c.kdcustomer where convert(int,b.kdkategori) > 3  group by a.kdcustomer,c.nmcustomer,c.alamat,a.kdbarang,b.nmbarang,b.kd1"



sqlA = "select * from (" & sqlA1 & ") a where jml<>0 "

sqlX = "insert into fixrute select kdbarang + '_' + '" & fixrute_TU.txtperiode & "','" & fixrute_TU.txtperiode & "','','" & Format(txttgl1, "yyyy/MM/dd") & "','" & fixrute_TU.lblkdteknisi & "',kdcustomer,kdbarang,'','1900/01/01','',0,0,'A01' from (" & sqlA & ") a where kdcustomer='" & rs!kdcustomer & "'"
con.Execute (sqlX)

TimerALL.Interval = 10
fixrute_TU.TimerALL.Interval = 10
fixrute.TimerALL.Interval = 10


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

If KODE = 2 Or KODE = 3 Then
rs.AbsolutePosition = lblpos
End If

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









Private Sub txttgl1_Change()
Call nul(txttgl1)

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

