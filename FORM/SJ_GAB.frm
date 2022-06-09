VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form SJ_GAB 
   BorderStyle     =   0  'None
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4290
      Left            =   225
      TabIndex        =   2
      Top             =   1935
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   7567
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
      TabIndex        =   5
      Top             =   855
      Width           =   13515
      _Version        =   524288
      _ExtentX        =   23839
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
      TabIndex        =   4
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
      Picture         =   "SJ_GAB.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   375
      Left            =   2745
      TabIndex        =   1
      ToolTipText     =   "Tambah Customer Baru"
      Top             =   1485
      Width           =   420
      _ExtentX        =   741
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "SJ_GAB.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdprint 
      Height          =   735
      Left            =   13725
      TabIndex        =   3
      ToolTipText     =   "Cetak"
      Top             =   1980
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
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
      Picture         =   "SJ_GAB.frx":8C19
      ButtonStyle     =   4
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   13770
      Picture         =   "SJ_GAB.frx":C676
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Out SJ Gabungan"
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
      TabIndex        =   8
      Top             =   135
      Width           =   6720
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
      TabIndex        =   7
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   7695
      Width           =   1155
   End
   Begin VB.Image SJ_GAB 
      Height          =   6900
      Left            =   0
      Picture         =   "SJ_GAB.frx":CA36
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14505
   End
End
Attribute VB_Name = "SJ_GAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim rscp As ADODB.Recordset
Dim rsACC As ADODB.Recordset
Dim rsX As ADODB.Recordset

Private Sub Cetak()

Unload AR_SJ

sqlX = "select * from sj_detail where nosj='" & rs!noSJ & "' and tgl='" & Format(rs!tgl, "yyyy/MM/dd") & "' and kdcustomer='" & rs!kdcustomer & "'"

Set rsX = con.Execute(sqlX)

With AR_SJ.DC1
.ConnectionString = koneksi
.Source = sqlX
End With

With AR_SJ
.fldunit.DataField = "unit"
.fldnmbarang.DataField = "nmbarang"
.fldsatuan.DataField = "satuan"
.fldketerangan.DataField = "keterangan"
.fldkdbarang.DataField = "kdbarang"
.fldkdkategori.DataField = "kdkategori"


.lblnosj = "-"
.lblnosj1 = rs!noSJ
.lblnmcustomer = rs!nmcustomer
.lbltglSJ = Format(rs!tgl, "dd/MM/yyyy")
.lblalamat = rs!alamat


.lblNB = ""

sqlACC = "select * from Signature where kdFrm='" & rs!kdgudang & "'"
Set rsACC = con.Execute(sqlACC)

.lblAcc1 = rsACC!Acc1
.lblAcc2 = rsACC!Acc2
.lblAcc3 = rsACC!Acc3
.lblAcc4 = rsACC!Acc4


.lblCP = rs!CP
.lbltelp = rs!telp


AR_SJ.Show vbModal

End With

End Sub


Private Sub cmdBR_Click()
TimerAll.Interval = 10
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

Private Sub cmdT_Click(Index As Integer)

End Sub

Private Sub cmdprint_Click()
On Error GoTo hell
Call Cetak

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
End Sub

Private Sub cmdprint_KeyPress(KeyAscii As Integer)
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
.Width = 85
.Caption = "TANGGAL"
.Alignment = dbgCenter
End With
'
With DataGrid1.Columns(1)
.Caption = "NO SJ"
.Width = 100
.Alignment = dbgCenter
End With

With DataGrid1.Columns(2)
.Caption = "kdgudang"
.Width = 0
End With
'
With DataGrid1.Columns(3)
.Caption = "KD CUST"
.Width = 60
.Alignment = dbgCenter
End With

With DataGrid1.Columns(4)
.Caption = "CUSTOMER"
.Width = 250
End With
'
With DataGrid1.Columns(5)
.Caption = "ALAMAT"
.Width = 350
End With

With DataGrid1.Columns(6)
.Caption = "CP"
.Width = 0
End With

With DataGrid1.Columns(7)
.Caption = "TELP"
.Width = 0
End With


Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

sql = "select * from SJ_HEADER where nmcustomer like '%" & txtcari & "%' or nosj like '%" & txtcari & "%' or alamat like '%" & txtcari & "%'  order by tgl"

Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub DataGrid1_DblClick()
On Error GoTo hell
Call Cetak

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
    
   Call Cetak

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0




End Sub




Private Sub Image3_Click()

End Sub

Private Sub TimerAll_Timer()
Call all

TimerAll.Interval = 0
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
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
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub








