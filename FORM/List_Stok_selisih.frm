VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form List_Stok_selisih 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   ScaleHeight     =   5670
   ScaleWidth      =   9825
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
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   1
      Top             =   855
      Width           =   8970
      _Version        =   524288
      _ExtentX        =   15822
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   5130
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
      Picture         =   "List_Stok_selisih.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3885
      Left            =   135
      TabIndex        =   0
      Top             =   945
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6853
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   255
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
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   630
      TabIndex        =   4
      Top             =   6435
      Width           =   1155
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Info Saldo Stok Kurang"
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
      Width           =   5280
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   9270
      Picture         =   "List_Stok_selisih.frx":6862
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   5595
      Left            =   0
      Picture         =   "List_Stok_selisih.frx":6C22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9780
   End
End
Attribute VB_Name = "List_Stok_selisih"
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

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
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
'
With datagrid1.Columns(0)
.Width = 90
.Caption = "KODE"
.Alignment = dbgCenter
End With

With datagrid1.Columns(1)
.Caption = "BARANG"
.Width = 180
End With

With datagrid1.Columns(2)
.Caption = "SATUAN"
.Width = 70
.Alignment = dbgCenter
End With

With datagrid1.Columns(3)
.Caption = "S. AWAL"
.Width = 70
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With

With datagrid1.Columns(4)
.Caption = "KELUAR"
.Width = 70
.Alignment = dbgRight
.NumberFormat = "#,###0"
End With


With datagrid1.Columns(5)
.Caption = "S. AKHIR"
.Width = 70
.Alignment = dbgRight
.NumberFormat = "#,###0"

End With




Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

If LBLKODE = "FREE" Then
sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
         "from RKP_stok where kdgudang='" & Free_D.lblkdgudang & "' and tgl <= '" & Format(Free_D.txttglfree, "yyyy/MM/dd") & "' and  kdbarang in (select kdbarang from free_d where kdfree='" & Free_D.lblKDFREE & "') group by kdbarang" & vbCrLf & _
         "Union All" & vbCrLf & _
         "select kdbarang,0 as unit,unit as UKeluar from free_d where kdfree='" & Free_D.lblKDFREE & "'"
         

ElseIf LBLKODE = "PINJAM" Then
sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
         "from RKP_stok where kdgudang='" & Pinjam_D.lblkdgudang & "' and tgl <= '" & Format(Pinjam_D.txttglpinjam, "yyyy/MM/dd") & "' and  kdbarang in (select kdbarang from pinjam_d where kdpinjam='" & Pinjam_D.lblkdPinjam & "') group by kdbarang" & vbCrLf & _
         "Union All" & vbCrLf & _
         "select kdbarang,0 as unit,unit as UKeluar from pinjam_d where kdpinjam='" & Pinjam_D.lblkdPinjam & "'"

ElseIf LBLKODE = "SEWA" Then
sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
         "from RKP_stok where kdgudang='" & Sewa_d.lblkdgudang & "' and tgl <= '" & Format(Sewa_d.txttglsewa, "yyyy/MM/dd") & "' and  kdbarang in (select kdbarang from sewa_d where kdsewa='" & Sewa_d.lblkdsewa & "') group by kdbarang" & vbCrLf & _
         "Union All" & vbCrLf & _
         "select kdbarang,0 as unit,unit as UKeluar from sewa_d where kdsewa='" & Sewa_d.lblkdsewa & "'"

ElseIf LBLKODE = "PERBAIKAN" Then
sqlCS1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as Unit,0 as UKeluar" & vbCrLf & _
         "from RKP_stok where kdgudang='" & Perbaikan_D.lblkdgudang1 & "' and tgl <= '" & Format(Perbaikan_D.txttglperbaikan, "yyyy/MM/dd") & "' and  kdbarang in (select kdbarang from perbaikan_d where kdperbaikan='" & Perbaikan_D.lblKDPerbaikan & "') group by kdbarang" & vbCrLf & _
         "Union All" & vbCrLf & _
         "select kdbarang,0 as unit,unit as UKeluar from perbaikan_d where kdperbaikan='" & Perbaikan_D.lblKDPerbaikan & "'"

End If

sqlCS2 = "select a.kdbarang,b.nmbarang,b.satuan,sum(a.unit) as unit,sum(a.UKeluar) as Ukeluar from (" & sqlCS1 & ") a left join barang b on a.kdbarang=b.kdbarang group by a.kdbarang,b.nmbarang,b.satuan"

sqlCS = "select kdbarang,nmbarang,satuan,unit + Ukeluar as Sawal, Ukeluar, Unit from (" & sqlCS2 & ") a where unit < 0 order by kdbarang"


Set rsCS = con.Execute(sqlCS)
Set datagrid1.DataSource = rsCS
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
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






