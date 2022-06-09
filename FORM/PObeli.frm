VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form PObeli 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   16905
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   16905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerG 
      Left            =   6165
      Top             =   4815
   End
   Begin VB.Timer TimerAll 
      Left            =   5625
      Top             =   4815
   End
   Begin VB.TextBox TXTCARI 
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
      Height          =   300
      Left            =   3555
      TabIndex        =   7
      Top             =   9315
      Width           =   2850
   End
   Begin VB.ComboBox CMBCARI 
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
      Height          =   345
      Left            =   1665
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   9315
      Width           =   1860
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7665
      Left            =   180
      TabIndex        =   0
      Top             =   990
      Width           =   15630
      _ExtentX        =   27570
      _ExtentY        =   13520
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
      Left            =   270
      TabIndex        =   8
      Top             =   675
      Width           =   15450
      _Version        =   524288
      _ExtentX        =   27252
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
      Left            =   15840
      TabIndex        =   1
      ToolTipText     =   "Tambah"
      Top             =   1350
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
      Picture         =   "PObeli.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   1
      Left            =   15840
      TabIndex        =   2
      ToolTipText     =   "Ubah"
      Top             =   2295
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
      Picture         =   "PObeli.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   2
      Left            =   15840
      TabIndex        =   3
      ToolTipText     =   "Hapus"
      Top             =   3240
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
      Picture         =   "PObeli.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   3
      Left            =   15840
      TabIndex        =   4
      ToolTipText     =   "Refresh"
      Top             =   4185
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
      Picture         =   "PObeli.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   15840
      TabIndex        =   5
      ToolTipText     =   "Cari Data"
      Top             =   5130
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
      Picture         =   "PObeli.frx":C086
      ButtonStyle     =   4
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   540
      TabIndex        =   12
      Top             =   9945
      Width           =   735
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA TIDAK ADA"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11250
      TabIndex        =   11
      Top             =   9405
      Width           =   2220
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   12060
      Picture         =   "PObeli.frx":EFAC
      Stretch         =   -1  'True
      Top             =   8910
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Pembelian Barang"
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
      Left            =   1035
      TabIndex        =   10
      Top             =   0
      Width           =   5685
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1440
      Top             =   8910
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Pencarian"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1665
      TabIndex        =   9
      Top             =   8955
      Width           =   4560
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6435
      Picture         =   "PObeli.frx":157FE
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   15885
      Picture         =   "PObeli.frx":226AE
      Stretch         =   -1  'True
      Top             =   315
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   9960
      Left            =   0
      Picture         =   "PObeli.frx":22A6E
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   16830
   End
End
Attribute VB_Name = "PObeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori, sqlcek As String
Dim kode As Integer
Dim rsmax As ADODB.Recordset
Dim rscek As ADODB.Recordset
Dim color As Long, flag As Byte

Private Sub cek_dalem()
sqlcek = "select * from PObeli_d where kdPObeli='" & rs!kdPObeli & "'"
Set rscek = con.Execute(sqlcek)
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub



'untuk set cursor pada saat dihapus
Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub


Private Sub tbl()
If rs.RecordCount = 0 Then
    cmdT(1).Enabled = False
    cmdT(2).Enabled = False
    DataGrid1.Enabled = False
    img1.Visible = True
    lbl1.Visible = True
Else
    cmdT(1).Enabled = True
    cmdT(2).Enabled = True
    DataGrid1.Enabled = True
    img1.Visible = False
    lbl1.Visible = False
End If
End Sub


Private Sub LG()
On Error GoTo hell

With DataGrid1.Columns(0)
.Width = 140
.Caption = "KODE"
.Alignment = dbgCenter
End With

With DataGrid1.Columns(1)
.Caption = "TANGGAL"
.Width = 85
.Alignment = dbgCenter
End With

With DataGrid1.Columns(2)
.Caption = "kdgudang"
.Width = 0
.Alignment = dbgCenter
End With

With DataGrid1.Columns(3)
.Caption = "GUDANG"
.Width = 250
End With

With DataGrid1.Columns(4)
.Caption = "KETERANGAN"
.Width = 420
End With


With DataGrid1.Columns(5)
.Caption = "NO. EASAP"
.Width = 100
.Alignment = dbgCenter
End With






Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()
PObeli_d.lblkode = 1
PObeli_d.Show vbModal
End Sub

Private Sub ubh()
PObeli_d.lblkode = 2
lblpos = rs.AbsolutePosition
kode = 2

PObeli_d.txtkdPO = rs!kdPObeli
PObeli_d.txttglPO = rs!tglPObeli
PObeli_d.lblkdgudang = rs!kdgudang
PObeli_d.lblnmgudang = rs!nmgudang
PObeli_d.txtketerangan = rs!keterangan
PObeli_d.txtnoEASAP = rs!noEASAP
PObeli_d.txttglPO.Enabled = False
PObeli_d.cmdBR.Enabled = False


PObeli_d.Show vbModal
End Sub

Private Sub hps()
On Error GoTo hell
kode = 3
Call max
    
    
    Call cek_dalem
    If rscek.RecordCount <> 0 Then
        MsgBox "Data Tidak dapat dihapus, karena Detail PO masih ada", vbCritical, "Error !"
        Exit Sub
    
    Else
        ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
        If ms = vbYes Then
            sql = "delete from PObeli where kdpobeli='" & rs!kdPObeli & "' "
            con.Execute (sql)
            
            TimerAll.Interval = 10
        Else
            Exit Sub
        End If
    End If
    

Exit Sub
hell:
MsgBox err.Description
End Sub


Private Sub all()
If TXTCARI = "" Then
sql = "select a.kdPObeli,a.tglPObeli,a.kdgudang,b.nmgudang,a.keterangan,a.noeasap from PObeli a left join gudang b on a.kdgudang=b.kdgudang where a.kdPObeli in (select kdPO from OTS_PO) order by a.tglPObeli desc,a.kdPObeli"
Else
    If CMBCARI.ListIndex < 4 Then
    sql = "select a.kdPObeli,a.tglPObeli,a.kdgudang,b.nmgudang,a.keterangan,a.noeasap from PObeli a left join gudang b on a.kdgudang=b.kdgudang where " & kategori & " like '%" & TXTCARI & "%'  and a.kdPObeli in (select kdPO from OTS_PO) order by a.tglPObeli desc,a.kdPObeli"
    Else
    sql = "select a.kdPObeli,a.tglPObeli,a.kdgudang,b.nmgudang,a.keterangan,a.noeasap from PObeli a left join gudang b on a.kdgudang=b.kdgudang where tglPObeli = '" & Format(TXTCARI, "yyyy/MM/dd") & "' and a.kdPObeli in (select kdPO from OTS_PO) order by a.tglPObeli desc,a.kdPObeli"
    End If
End If

Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs

Call LG
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "a.kdPObeli"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "b.nmgudang"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "a.keterangan"
ElseIf CMBCARI.ListIndex = 3 Then
kategori = "a.noeasap"
ElseIf CMBCARI.ListIndex = 4 Then
kategori = "a.TglPObeli"
End If

TimerAll.Interval = 10
End Sub

Private Sub CMBCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
     If rs.RecordCount <> 0 Then
     Call ubh
     End If
ElseIf Index = 2 Then
     If rs.RecordCount <> 0 Then
     Call hps
     End If
ElseIf Index = 3 Then
TXTCARI = ""
Call all
ElseIf Index = 4 Then
TXTCARI = ""
    If TXTCARI.Enabled = True Then
    Me.Height = Me.Height - 1170

    TXTCARI.Enabled = False
    CMBCARI.Enabled = False
    Else
    Me.Height = Me.Height + 1170

    TXTCARI.Enabled = True
    CMBCARI.Enabled = True
    End If
End If
End Sub



Private Sub cmdT_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
 TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub

Private Sub datagrid1_Click()
TimerG.Interval = 10
End Sub

Private Sub datagrid1_DblClick()
 If rs.RecordCount <> 0 Then
 Call ubh
 End If

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyLeft Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyRight Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub


Private Sub Form_Load()

GradientForm Me, 0

Me.Height = Me.Height - 1170


CMBCARI.AddItem "KODE PO"
CMBCARI.AddItem "GUDANG"
CMBCARI.AddItem "KETERANGAN"
CMBCARI.AddItem "NO EASAP"
CMBCARI.AddItem "TANGGAL PO"
CMBCARI.ListIndex = 0



TimerAll.Interval = 10
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerAll.Interval = 0

End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerAll.Interval = 10
End Sub

Private Sub TXTCARI_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTCARI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    DataGrid1.SetFocus
    TimerG.Interval = 10
    Else
    SendKeys vbTab
    End If
ElseIf KeyAscii = 27 Then
Unload Me
'ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
' Call tbh
'ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
' If rs.RecordCount <> 0 Then
' Call ubh
' End If
'ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
' If rs.RecordCount <> 0 Then
' Call hps
' End If
'ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
' Call all
End If
End Sub




