VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Sewa_DTU 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtunit 
      Alignment       =   1  'Right Justify
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
      Text            =   "0"
      Top             =   1485
      Width           =   1050
   End
   Begin VB.TextBox txtketerangan 
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
      TabIndex        =   5
      Top             =   2205
      Width           =   6585
   End
   Begin VB.TextBox txtharga 
      Alignment       =   1  'Right Justify
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
      Left            =   4590
      TabIndex        =   3
      Text            =   "0"
      Top             =   1485
      Width           =   1275
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   8
      Top             =   720
      Width           =   7980
      _Version        =   524288
      _ExtentX        =   14076
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   675
      TabIndex        =   7
      Top             =   2835
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
      Picture         =   "Sewa_DTU.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   870
      Left            =   8235
      TabIndex        =   6
      ToolTipText     =   "Simpan"
      Top             =   1890
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
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
      Picture         =   "Sewa_DTU.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdhrg 
      Height          =   420
      Left            =   4095
      TabIndex        =   2
      ToolTipText     =   "Cari Harga Rata2 Barang "
      Top             =   1440
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
      Picture         =   "Sewa_DTU.frx":92CF
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   7605
      TabIndex        =   0
      Top             =   1080
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
      Picture         =   "Sewa_DTU.frx":BB01
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label lblmaxunit 
      Height          =   330
      Left            =   7335
      TabIndex        =   21
      Top             =   4545
      Width           =   1320
   End
   Begin VB.Label lblunit_awal 
      Caption         =   "Label5"
      Height          =   330
      Left            =   5265
      TabIndex        =   20
      Top             =   4365
      Width           =   1185
   End
   Begin VB.Label lblnmbarang 
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
      Left            =   2925
      TabIndex        =   19
      Top             =   1125
      Width           =   4695
   End
   Begin VB.Label lblkdbarang 
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
      Left            =   1395
      TabIndex        =   18
      Top             =   1125
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BARANG :"
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
      Left            =   135
      TabIndex        =   17
      Top             =   1170
      Width           =   1320
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   675
      TabIndex        =   16
      Top             =   4725
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UNIT :"
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
      Left            =   135
      TabIndex        =   15
      Top             =   1530
      Width           =   870
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN :"
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
      Left            =   135
      TabIndex        =   14
      Top             =   2250
      Width           =   1320
   End
   Begin VB.Label lblsatuan 
      BackStyle       =   0  'Transparent
      Caption         =   "SATUAN"
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
      Left            =   2475
      TabIndex        =   13
      Top             =   1530
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Barang"
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
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   3525
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   8280
      Picture         =   "Sewa_DTU.frx":E333
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "HARGA :"
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
      Left            =   3420
      TabIndex        =   11
      Top             =   1530
      Width           =   870
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "RUPIAH :"
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
      Left            =   135
      TabIndex        =   10
      Top             =   1890
      Width           =   870
   End
   Begin VB.Label lblrupiah 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   1395
      TabIndex        =   4
      Top             =   1845
      Width           =   2265
   End
   Begin VB.Label lblkdsewa_d 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2700
      TabIndex        =   9
      Top             =   4680
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   3390
      Left            =   0
      Picture         =   "Sewa_DTU.frx":E6F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "Sewa_DTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim a As Integer
Dim rsA As ADODB.Recordset
Dim color As Long, flag As Byte

Private Sub cek_POminta()
sqlA = "select a.kdbarang,sum(a.unit) as UKeluar,b.kdpo from sewa_d a left join sewa b  on a.kdsewa =b.kdsewa where b.kdpo ='" & Sewa_d.txtkdPO & "' group by a.kdbarang,b.kdpo"
            
sqlA1 = "select a.kdbarang,b.nmbarang,a.unit,isnull(c.Ukeluar,0) as Ukeluar,b.satuan,a.keterangan,a.kdpo_d from po_d a left join barang b " & vbCrLf & _
        "on a.kdbarang=b.kdbarang left join (" & sqlA & ") c on a.kdPO=c.kdPO and a.kdbarang=c.kdbarang where a.kdpo='" & Sewa_d.txtkdPO & "' "

sqlA2 = "select kdbarang,nmbarang,unit,ukeluar,unit - ukeluar as sisa,satuan,keterangan,kdPO_D from (" & sqlA1 & ") a  where kdbarang='" & lblkdbarang & "'"

Set rsA = con.Execute(sqlA2)
        
If rsA.RecordCount <> 0 Then
lblmaxunit = CCur(rsA!sisa) + CCur(lblunit_awal)
Else
lblmaxunit = CCur(lblunit_awal)
End If

End Sub


Private Sub cmdBR_Click()
Barang_BR.LBLKODE = UCase("SEWA_DTU")
Barang_BR.Show vbModal

End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdhrg_Click()

sql1 = "select kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - repair) as unit," & vbCrLf & _
        "sum(U_beliX + U_RpinjamX + U_RsewaX + M_unitX - U_freeX - U_pinjamX - U_sewaX  - K_unitX - repairX) as rupiah" & vbCrLf & _
        "from RKP_stok where kdgudang='" & Sewa_d.lblkdgudang & "' and tgl <= '" & Format(Sewa_d.txttglsewa, "yyyy/MM/dd") & "' and kdbarang='" & lblkdbarang & "' group by kdbarang"

Set rs1 = con.Execute(sql1)

If rs1.RecordCount <> 0 Then
 If CCur(rs1!unit) + CCur(txtunit) = 0 Then
 txtharga = 0
 Else
 txtharga = Format((CCur(rs1!rupiah) + CCur(lblrupiah)) / (CCur(rs1!unit) + CCur(txtunit)), "#,###0")
 End If
Else
 txtharga = 0
 MsgBox "Tidak ada Referensi Harga !", vbCritical, "Error !"
End If

cmdhrg.Enabled = False

End Sub

Private Sub cmdhrg_KeyPress(KeyAscii As Integer)
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






Private Sub cmdsimpan_Click()
'On Error GoTo hell

    If lblnmbarang = "" Or lblkdbarang = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !"
    Exit Sub
    Else
         If LBLKODE = 1 Then
         con.Execute ("delete from PO_d where kdpo='" & Sewa_d.txtkdPO & "' and kdbarang in (select kdbarang from barang where kdkategori='02') ")
         con.Execute ("delete from sewa_d where kdsewa='" & Sewa_d.lblkdsewa & "' and kdbarang in (select kdbarang from barang where kdkategori='02') ")
         
         con.Execute ("insert into PO_D values ('" & lblkdbarang & "_" & Sewa_d.txtkdPO & "','" & Sewa_d.txtkdPO & "','" & lblkdbarang & "'," & CCur(txtunit) & ",'" & UCase(txtketerangan) & "')")
         con.Execute ("insert into sewa_D values ('" & lblkdbarang & "_" & Sewa_d.lblkdsewa & "','" & Sewa_d.lblkdsewa & "','" & lblkdbarang & "'," & CCur(txtunit) & "," & CCur(txtharga) & "," & CCur(lblrupiah) & ",'" & UCase(txtketerangan) & "')")
         
         SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
         MsgBox "Data Telah di Tambah di POnya Juga", vbInformation, "Info !"
         Sewa_d.TimerALL.Interval = 10

         ElseIf LBLKODE = 2 Then
             Call cek_POminta
         
             If CCur(lblmaxunit) < CCur(txtunit) Then
                SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
                 MsgBox "PO yg tersedia = " & lblmaxunit & " Unit", vbCritical, "Error !"
                 
                 Exit Sub
             Else

                 sql = "update sewa_D set unit=" & CCur(txtunit) & ",harga=" & CCur(txtharga) & ",rupiah=" & CCur(lblrupiah) & ",keterangan='" & UCase(txtketerangan) & "' where kdsewa_D='" & lblkdsewa_d & "'"
                 con.Execute (sql)
                 SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
                 MsgBox "Data Telah di Ubah", vbInformation, "Info !"
    
                 Sewa_d.TimerALL.Interval = 10
             
             End If

         End If
         
         Unload Me
    End If
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

Call nul(lblnmbarang)
Call nul(lblkdbarang)
End Sub





Private Sub lblkdbarang_Change()
Call nul(lblkdbarang)
End Sub

Private Sub lblnmbarang_Change()
Call nul(lblnmbarang)
End Sub

Private Sub txtharga_Change()
Call nul(txtharga)
On Error GoTo hell
lblrupiah = CCur(txtunit) * CCur(txtharga)
lblrupiah = FormatNumber(lblrupiah, 0)

Exit Sub
hell:
lblrupiah = 0

End Sub

Private Sub txtharga_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtharga_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txtharga_KeyPress(KeyAscii As Integer)
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

Private Sub txtharga_LostFocus()
On Error GoTo hell

txtharga = FormatNumber(txtharga, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtharga.SetFocus

End Sub

Private Sub txtketerangan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtketerangan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txtketerangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtketerangan_LostFocus()
txtketerangan = UCase(txtketerangan)
End Sub



Private Sub txtunit_Change()
Call nul(txtunit)
On Error GoTo hell
lblrupiah = CCur(txtunit) * CCur(txtharga)
lblrupiah = FormatNumber(lblrupiah, 0)

Exit Sub
hell:
lblrupiah = 0

End Sub

Private Sub txtunit_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtunit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtunit_KeyPress(KeyAscii As Integer)
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

Private Sub txtunit_LostFocus()
On Error GoTo hell

txtunit = FormatNumber(txtunit, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtunit.SetFocus

End Sub




