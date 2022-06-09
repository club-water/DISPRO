VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form PO_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19500
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   19500
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
      Height          =   6630
      Left            =   225
      TabIndex        =   1
      Top             =   1935
      Width           =   18060
      _ExtentX        =   31856
      _ExtentY        =   11695
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
      TabIndex        =   3
      Top             =   855
      Width           =   18060
      _Version        =   524288
      _ExtentX        =   31856
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
      TabIndex        =   2
      Top             =   8865
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
      Picture         =   "PO_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label lblkdkategori 
      Caption         =   "lblkategori"
      Height          =   315
      Left            =   1530
      TabIndex        =   7
      Top             =   9450
      Width           =   1155
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2790
      Picture         =   "PO_BR.frx":6862
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   18405
      Picture         =   "PO_BR.frx":13712
      Stretch         =   -1  'True
      Top             =   360
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Permintaan Barang"
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
      Left            =   855
      TabIndex        =   6
      Top             =   135
      Width           =   7755
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
      TabIndex        =   5
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   9450
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   9330
      Left            =   0
      Picture         =   "PO_BR.frx":13AD2
      Stretch         =   -1  'True
      Top             =   45
      Width           =   19455
   End
End
Attribute VB_Name = "PO_BR"
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

With DataGrid1.Columns(0)
.Width = 120
.Caption = "KODE"
.Alignment = dbgCenter
End With

With DataGrid1.Columns(1)
.Caption = "TANGGAL"
.Width = 80
.Alignment = dbgCenter
End With

With DataGrid1.Columns(2)
.Caption = "kdgudang"
.Width = 0
.Alignment = dbgCenter
End With

With DataGrid1.Columns(3)
.Caption = "GUDANG"
.Width = 150
End With

With DataGrid1.Columns(4)
.Caption = "KODE"
.Width = 60
.Alignment = dbgCenter
End With


If lblkdkategori <> "04" Then
    With DataGrid1.Columns(5)
    .Caption = "CUSTOMER"
    .Width = 220
    End With
    
    With DataGrid1.Columns(6)
    .Caption = "ALAMAT"
    .Width = 300
    End With
    
    With DataGrid1.Columns(7)
    .Caption = "kdkategori"
    .Width = 0
    .Alignment = dbgCenter
    End With
    
    With DataGrid1.Columns(8)
    .Caption = "KATEGORI"
    .Width = 90
    End With
    
    With DataGrid1.Columns(9)
    .Caption = "KETERANGAN"
    .Width = 150
    End With
    
    With DataGrid1.Columns(10)
    .Caption = "KD BRG"
    .Width = 0
    .Alignment = dbgCenter
    End With
    
    With DataGrid1.Columns(11)
    .Caption = "BARANG"
    .Width = 0
    End With
    
Else

 With DataGrid1.Columns(5)
    .Caption = "CUSTOMER"
    .Width = 200
    End With
    
    With DataGrid1.Columns(6)
    .Caption = "ALAMAT"
    .Width = 250
    End With
    
    With DataGrid1.Columns(7)
    .Caption = "kdkategori"
    .Width = 0
    .Alignment = dbgCenter
    End With
    
    With DataGrid1.Columns(8)
    .Caption = "KATEGORI"
    .Width = 80
    End With
    
    With DataGrid1.Columns(9)
    .Caption = "KETERANGAN"
    .Width = 0
    End With
    
    With DataGrid1.Columns(10)
    .Caption = "KD BRG"
    .Width = 80
    .Alignment = dbgCenter
    End With
    
    With DataGrid1.Columns(11)
    .Caption = "BARANG"
    .Width = 150
    End With

End If


If lblkdkategori = "04" Or lblkdkategori = "05" Then

    With DataGrid1.Columns(12)
    .Caption = "NO EASAP"
    .Width = 0
    .Alignment = dbgCenter
    End With
Else

    With DataGrid1.Columns(12)
    .Caption = "NO EASAP"
    .Width = 100
    .Alignment = dbgCenter
    End With
End If


Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

If lblkdkategori = "04" Then

    If txtcari = "" Then
    sql = "select a.kdPO,a.tglPO,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.kdkategori,d.nmkategori,a.keterangan,isnull(a.kdbarang,'') as kdbarang,isnull(e.nmbarang,'') as nmbarang,isnull(a.noEASAP,'') as noEASAP from PO a left join gudang b on a.kdgudang=b.kdgudang left join customer c on a.kdcustomer= c.kdcustomer " & vbCrLf & _
          "left join kategori d on a.kdkategori=d.kdkategori left join barang e on a.kdbarang=e.kdbarang where a.kdPO in (select kdpo from OTS_POminta where kdkategori in ('04','05')) order by a.kdPO"
    
    Else
    
    sql = "select a.kdPO,a.tglPO,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.kdkategori,d.nmkategori,a.keterangan,isnull(a.kdbarang,'') as kdbarang,isnull(e.nmbarang,'') as nmbarang,isnull(a.noEASAP,'') as noEASAP from PO a left join gudang b on a.kdgudang=b.kdgudang left join customer c on a.kdcustomer= c.kdcustomer " & vbCrLf & _
          "left join kategori d on a.kdkategori=d.kdkategori left join barang e on a.kdbarang=e.kdbarang where (c.nmcustomer like '%" & txtcari & "%' or a.kdpo like '%" & txtcari & "%' or b.nmgudang like '%" & txtcari & "%' or a.keterangan like '%" & txtcari & "%' or a.noEASAP like '%" & txtcari & "%') and  a.kdpo in (select kdpo from OTS_POminta where kdkategori in ('04','05')) order by a.kdPO"
    
    
    End If

Else
    If txtcari = "" Then
    sql = "select a.kdPO,a.tglPO,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.kdkategori,d.nmkategori,a.keterangan,isnull(a.kdbarang,'') as kdbarang,isnull(e.nmbarang,'') as nmbarang,isnull(a.noEASAP,'') as noEASAP from PO a left join gudang b on a.kdgudang=b.kdgudang left join customer c on a.kdcustomer= c.kdcustomer " & vbCrLf & _
          "left join kategori d on a.kdkategori=d.kdkategori left join barang e on a.kdbarang=e.kdbarang where a.kdPO in (select kdpo from OTS_POminta where kdkategori='" & lblkdkategori & "') order by a.kdPO"
    
    Else
    
    sql = "select a.kdPO,a.tglPO,a.kdgudang,b.nmgudang,a.kdcustomer,c.nmcustomer,c.alamat,a.kdkategori,d.nmkategori,a.keterangan,isnull(a.kdbarang,'') as kdbarang,isnull(e.nmbarang,'') as nmbarang,isnull(a.noEASAP,'') as noEASAP from PO a left join gudang b on a.kdgudang=b.kdgudang left join customer c on a.kdcustomer= c.kdcustomer " & vbCrLf & _
          "left join kategori d on a.kdkategori=d.kdkategori left join barang e on a.kdbarang=e.kdbarang where (c.nmcustomer like '%" & txtcari & "%' or a.kdpo like '%" & txtcari & "%' or b.nmgudang like '%" & txtcari & "%' or a.keterangan like '%" & txtcari & "%' or a.noEASAP like '%" & txtcari & "%') and  a.kdPO in (select kdpo from OTS_POminta where kdkategori='" & lblkdkategori & "') order by a.kdPO"
    
    
    End If

End If

Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub DataGrid1_DblClick()
On Error GoTo hell
If lblkode = UCase("FREE_D") Then
Free_D.txtkdPO = rs!kdPO
Free_D.lbltglPO = rs!tglPO
Free_D.lblkdgudang = rs!kdgudang
Free_D.lblnmgudang = rs!nmgudang
Free_D.lblkdcustomer = rs!kdcustomer
Free_D.lblnmcustomer = rs!nmcustomer
Free_D.lblalamat = rs!alamat
Free_D.txtketerangan = rs!keterangan
Free_D.lblnoEASAP = rs!noeasap
ElseIf lblkode = UCase("PINJAM_D") Then
Pinjam_D.txtkdPO = rs!kdPO
Pinjam_D.lbltglPO = rs!tglPO
Pinjam_D.lblkdgudang = rs!kdgudang
Pinjam_D.lblnmgudang = rs!nmgudang
Pinjam_D.lblkdcustomer = rs!kdcustomer
Pinjam_D.lblnmcustomer = rs!nmcustomer
Pinjam_D.lblalamat = rs!alamat
Pinjam_D.txtketerangan = rs!keterangan
Pinjam_D.lblnoEASAP = rs!noeasap
ElseIf lblkode = UCase("SEWA_D") Then
Sewa_d.txtkdPO = rs!kdPO
Sewa_d.lbltglPO = rs!tglPO
Sewa_d.lblkdgudang = rs!kdgudang
Sewa_d.lblnmgudang = rs!nmgudang
Sewa_d.lblkdcustomer = rs!kdcustomer
Sewa_d.lblnmcustomer = rs!nmcustomer
Sewa_d.lblalamat = rs!alamat
Sewa_d.txtketerangan = rs!keterangan
Sewa_d.lblnoEASAP = rs!noeasap
ElseIf lblkode = UCase("PERBAIKAN_D") Then
Perbaikan_D.txtkdPO = rs!kdPO
Perbaikan_D.lbltglPO = rs!tglPO
Perbaikan_D.lblkdgudang1 = rs!kdgudang
Perbaikan_D.lblnmgudang1 = rs!nmgudang
Perbaikan_D.txtketerangan = rs!keterangan
Perbaikan_D.lblkdbarang = rs!kdbarang
Perbaikan_D.lblnmbarang = rs!nmbarang
Perbaikan_D.lblkdkategori = rs!kdkategori

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
    
   If lblkode = UCase("Free_D") Then
    Free_D.txtkdPO = rs!kdPO
    Free_D.lbltglPO = rs!tglPO
    Free_D.lblkdgudang = rs!kdgudang
    Free_D.lblnmgudang = rs!nmgudang
    Free_D.lblkdcustomer = rs!kdcustomer
    Free_D.lblnmcustomer = rs!nmcustomer
    Free_D.lblalamat = rs!alamat
    Free_D.txtketerangan = rs!keterangan
    Free_D.lblnoEASAP = rs!noeasap
    ElseIf lblkode = UCase("PINJAM_D") Then
    Pinjam_D.txtkdPO = rs!kdPO
    Pinjam_D.lbltglPO = rs!tglPO
    Pinjam_D.lblkdgudang = rs!kdgudang
    Pinjam_D.lblnmgudang = rs!nmgudang
    Pinjam_D.lblkdcustomer = rs!kdcustomer
    Pinjam_D.lblnmcustomer = rs!nmcustomer
    Pinjam_D.lblalamat = rs!alamat
    Pinjam_D.txtketerangan = rs!keterangan
    Pinjam_D.lblnoEASAP = rs!noeasap
    ElseIf lblkode = UCase("SEWA_D") Then
    Sewa_d.txtkdPO = rs!kdPO
    Sewa_d.lbltglPO = rs!tglPO
    Sewa_d.lblkdgudang = rs!kdgudang
    Sewa_d.lblnmgudang = rs!nmgudang
    Sewa_d.lblkdcustomer = rs!kdcustomer
    Sewa_d.lblnmcustomer = rs!nmcustomer
    Sewa_d.lblalamat = rs!alamat
    Sewa_d.txtketerangan = rs!keterangan
    Sewa_d.lblnoEASAP = rs!noeasap
    ElseIf lblkode = UCase("PERBAIKAN_D") Then
    Perbaikan_D.txtkdPO = rs!kdPO
    Perbaikan_D.lbltglPO = rs!tglPO
    Perbaikan_D.lblkdgudang1 = rs!kdgudang
    Perbaikan_D.lblnmgudang1 = rs!nmgudang
    Perbaikan_D.txtketerangan = rs!keterangan
    Perbaikan_D.lblkdbarang = rs!kdbarang
    Perbaikan_D.lblnmbarang = rs!nmbarang
    Perbaikan_D.lblkdkategori = rs!kdkategori

    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerAll.Interval = 10
End Sub




Private Sub TimerAll_Timer()
On Error Resume Next
Call all

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






