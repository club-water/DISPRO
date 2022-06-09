VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form PS_D_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   12900
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5235
      Left            =   225
      TabIndex        =   0
      Top             =   1395
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   9234
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
      TabIndex        =   1
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
      TabIndex        =   2
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
      Picture         =   "PS_D_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label lblkd 
      Caption         =   "Label1"
      Height          =   330
      Left            =   5130
      TabIndex        =   6
      Top             =   7740
      Width           =   1590
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   7695
      Width           =   1155
   End
   Begin VB.Label lbljudul 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Pinjam Pakai"
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
      Picture         =   "PS_D_BR.frx":6862
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
      Picture         =   "PS_D_BR.frx":6C22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "PS_D_BR"
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
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub LG()
On Error GoTo hell

With DataGrid1.Columns(0)
.Width = 100
.Caption = "KODE"
.Alignment = dbgCenter
End With

With DataGrid1.Columns(1)
.Caption = "BARANG"
.Width = 250
End With

With DataGrid1.Columns(2)
.Caption = "SISA"
.Width = 80
.Alignment = dbgRight
End With

With DataGrid1.Columns(3)
.Caption = "SATUAN"
.Width = 80
.Alignment = dbgCenter
End With

With DataGrid1.Columns(4)
.Caption = "HARGA"
.Width = 100
.NumberFormat = "#,###0"
.Alignment = dbgRight
End With

With DataGrid1.Columns(5)
.Caption = "RUPIAH"
.Width = 120
.NumberFormat = "#,###0"
.Alignment = dbgRight
End With






Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

If LBLKODE = "RPINJAM_D" Then

sql1 = "select kdbarang,harga,sum(unit) as unit from (select 'A' as kode,kdpinjam,kdbarang,unit,harga from pinjam_d where kdpinjam='" & lblkd & "' union " & vbCrLf & _
       "select 'B' as kode,a.kdpinjam,b.kdbarang,-sum(b.unit) as unit,b.harga from Rpinjam a left join Rpinjam_d b on a.kdRpinjam =b.kdRpinjam where a.kdpinjam='" & lblkd & "' group by a.kdpinjam,b.harga,b.kdbarang ) a group by kdpinjam,harga,kdbarang"

sql = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.harga,(a.unit * a.harga) as Rupiah from (" & sql1 & ") a left join barang b on a.kdbarang=b.kdbarang where a.unit <>0 "


ElseIf LBLKODE = "RSEWA_D" Then

sql1 = "select kdbarang,harga,sum(unit) as unit from (select 'A' as kode,kdsewa,kdbarang,unit,harga from sewa_d where kdsewa='" & lblkd & "' union " & vbCrLf & _
       "select 'B' as kode,a.kdsewa,b.kdbarang,-sum(b.unit) as unit,b.harga from Rsewa a left join Rsewa_d b on a.kdRsewa =b.kdRsewa where a.kdsewa='" & lblkd & "' group by a.kdsewa,b.harga,b.kdbarang ) a group by kdsewa,harga,kdbarang"

sql = "select a.kdbarang,b.nmbarang,a.unit,b.satuan,a.harga,(a.unit * a.harga) as Rupiah from (" & sql1 & ") a left join barang b on a.kdbarang=b.kdbarang where a.unit <>0 "


Else

End If



Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs
Call LG

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

If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then

 Call all

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

