VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Barang_TU 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5730
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglnon_aktif 
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
      Left            =   2745
      TabIndex        =   11
      Top             =   4770
      Width           =   1455
   End
   Begin VB.TextBox txtkdSAP 
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
      TabIndex        =   7
      Top             =   3780
      Width           =   2175
   End
   Begin VB.Timer TimerQR 
      Left            =   5445
      Top             =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   2610
      ScaleHeight     =   140
      ScaleMode       =   0  'User
      ScaleWidth      =   140
      TabIndex        =   26
      Top             =   6300
      Width           =   2385
   End
   Begin VB.CheckBox ChkNA 
      BackColor       =   &H00000000&
      Caption         =   "NON AKTIF"
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
      Height          =   330
      Left            =   1395
      TabIndex        =   10
      Top             =   4770
      Width           =   1320
   End
   Begin VB.TextBox TXTKD1 
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
      TabIndex        =   4
      Top             =   2700
      Width           =   2175
   End
   Begin VB.CheckBox ChkBFS 
      BackColor       =   &H00000000&
      Caption         =   "Buffer Stok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1395
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   4365
      Width           =   1365
   End
   Begin VB.TextBox txtBFS 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   2700
      TabIndex        =   9
      Text            =   "0"
      Top             =   4365
      Width           =   1050
   End
   Begin VB.ComboBox CMBKATEGORI 
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
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2295
      Width           =   2130
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
      Top             =   3060
      Width           =   4605
   End
   Begin VB.TextBox lblkdbarang 
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
      TabIndex        =   0
      Top             =   1215
      Width           =   2175
   End
   Begin VB.TextBox txtsatuan 
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
      TabIndex        =   2
      Top             =   1935
      Width           =   2175
   End
   Begin VB.Timer TimerCMB 
      Left            =   4500
      Top             =   1035
   End
   Begin VB.TextBox TXTnmbarang 
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
      Top             =   1575
      Width           =   4605
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   915
      Left            =   6120
      TabIndex        =   12
      ToolTipText     =   "Simpan"
      Top             =   1890
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Barang_TU.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   90
      TabIndex        =   22
      Top             =   720
      Width           =   5955
      _Version        =   524288
      _ExtentX        =   10504
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
      TabIndex        =   14
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
      Picture         =   "Barang_TU.frx":2A6D
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdQR 
      Height          =   915
      Left            =   6120
      TabIndex        =   13
      ToolTipText     =   "Simpan"
      Top             =   2835
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Barang_TU.frx":92CF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.TextBox txtmerk 
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
      TabIndex        =   6
      Top             =   3420
      Width           =   4605
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "KDSAP :"
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
      TabIndex        =   28
      Top             =   3825
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MERK :"
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
      TabIndex        =   27
      Top             =   3465
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "KD BAJA PUTIH :"
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
      Left            =   90
      TabIndex        =   25
      Top             =   2745
      Width           =   1320
   End
   Begin VB.Label lblfrm 
      Height          =   285
      Left            =   2880
      TabIndex        =   24
      Top             =   7785
      Width           =   1050
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   6210
      Picture         =   "Barang_TU.frx":D430
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   3600
      Picture         =   "Barang_TU.frx":D7F0
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   465
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
      Left            =   495
      TabIndex        =   23
      Top             =   0
      Width           =   3525
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "KATEGORI :"
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
      TabIndex        =   21
      Top             =   2340
      Width           =   1320
   End
   Begin VB.Label lblkdkategori 
      Height          =   330
      Left            =   3555
      TabIndex        =   20
      Top             =   2340
      Visible         =   0   'False
      Width           =   510
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
      TabIndex        =   19
      Top             =   3150
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SATUAN :"
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
      Left            =   180
      TabIndex        =   18
      Top             =   1980
      Width           =   1320
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   495
      TabIndex        =   17
      Top             =   7740
      Width           =   1545
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE :"
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
      TabIndex        =   16
      Top             =   1260
      Width           =   645
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
      TabIndex        =   15
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   5685
      Left            =   -45
      Picture         =   "Barang_TU.frx":EAAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Barang_TU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Adodb.Recordset
Dim rs1 As Adodb.Recordset
Dim sql As String
Dim sql1 As String
Dim a As Integer

Dim color As Long, flag As Byte

Private Sub Check1_Click()

End Sub

Private Sub ChkBFS_Click()
If ChkBFS.Value = 1 Then
txtBFS.Enabled = True
Else
txtBFS = 0
txtBFS.Enabled = False
End If
End Sub

Private Sub ChkBFS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub ChkNA_Click()
If ChkNA.Value = 1 Then
txttglnon_aktif = Format(Date, "dd/MM/yyyy")
txttglnon_aktif.Enabled = True
Else
txttglnon_aktif = "01/01/1900"
txttglnon_aktif.Enabled = False
End If
End Sub

Private Sub ChkNA_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
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

Private Sub cmdQR_Click()
Unload AR_QR

With AR_QR

   Set cQrCode = New ClassQR
   .Image1.Picture = cQrCode.GetPictureQrCode(lblkdbarang, 140, 140)
   If .Image1.Picture Is Nothing Then MsgBox "Error!"

.fldkdbarang.Text = lblkdbarang




AR_QR.Show vbModal

End With
End Sub

Private Sub cmdQR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub



Private Sub CMBKATEGORI_Click()
On Error Resume Next
sql1 = "select * from KATEGORIBRG where nmKATEGORI='" & CMBKATEGORI.Text & "'"
Set rs1 = con.Execute(sql1)

lblkdkategori = rs1!kdkategori

If lblkdkategori = "04" Or lblkdkategori = "10" Then
Call nul(TXTKD1)
TXTKD1.Enabled = True
Else
TXTKD1 = ""
TXTKD1.BackColor = vbWhite
TXTKD1.Enabled = False
End If
End Sub

Private Sub CMBKATEGORI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub cmdsimpan_Click()
On Error GoTo hell

    If TXTnmbarang = "" Or lblkdbarang = "" Then
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !!"
    Exit Sub
    Else
        
        If (lblkdkategori = "04" Or lblkdkategori = "10") And TXTKD1 = "" Then
        MsgBox "inputan belum lengkap !!", vbInformation, "Info !!"
        Exit Sub
        Else
             If LBLKODE = 1 Then
                 sql = "insert into barang values ('" & Trim(UCase(lblkdbarang)) & "','" & UCase(TXTnmbarang) & "','" & UCase(txtsatuan) & "','" & lblkdkategori & "','" & UCase(txtketerangan) & "'," & ChkBFS.Value & "," & CCur(txtBFS) & ",'" & UCase(TXTKD1) & "'," & ChkNA.Value & ",'" & UCase(txtmerk) & "','" & UCase(txtkdSAP) & "','1900/01/01')"
                 con.Execute (sql)
                 MsgBox "Data Telah Tersimpan", vbInformation, "Informasi !"
         
                 
             Else
'                 If UTAMA.lblstatus = 0 Then
'                    MsgBox "Data Tidak dapat diubah, karena bukan Administrator", vbCritical, "Error !!"
'                    Exit Sub
'                 Else
                    sql = "update barang set nmbarang='" & UCase(TXTnmbarang) & "',satuan='" & UCase(txtsatuan) & "',keterangan='" & UCase(txtketerangan) & "',kdkategori='" & lblkdkategori & "',BFS=" & ChkBFS.Value & ",Unit_BFS=" & CCur(txtBFS) & ",kd1='" & UCase(TXTKD1) & "',Non_aktif = " & ChkNA.Value & ",tglnon_aktif='" & Format(txttglnon_aktif, "yyyy/MM/dd") & "',merk='" & UCase(txtmerk) & "',kdSAP='" & UCase(txtkdSAP) & "' where kdbarang='" & lblkdbarang & "'"
                    con.Execute (sql)
                 
                    MsgBox "Data Telah di Ubah", vbInformation, "Informasi !"
'                 End If
    
                 
             End If
             
             
             If lblfrm = "BARANG_BR" Then
             Barang_BR.TimerAll.Interval = 10
             Barang_BR.TXTCARI = lblkdbarang
             Else
             Barang.TimerAll.Interval = 10
             End If
             
             Unload Me
        End If
    End If
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

ChkBFS.Value = 0

sql = "Select * from kategoriBRG order by kdkategori"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
CMBKATEGORI.AddItem rs!nmkategori
rs.MoveNext
Loop



txttglnon_aktif = "01/01/1900"
txttglnon_aktif.Enabled = False


Call nul(lblkdbarang)
Call nul(TXTnmbarang)

TimerCMB.Interval = 10


End Sub

Private Sub lblkdpic_Click()

End Sub










Private Sub lblkdbarang_Change()
Call nul(lblkdbarang)
End Sub

Private Sub lblkdbarang_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub lblkdbarang_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub lblkdbarang_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub lblkdbarang_LostFocus()
lblkdbarang = UCase(lblkdbarang)
'TimerQR.Interval = 10
End Sub

Private Sub TimerNO_Timer()

End Sub




Private Sub TimerCMB_Timer()
If LBLKODE = "1" Then
CMBKATEGORI.ListIndex = 0
End If


TimerCMB.Interval = 0
End Sub

Private Sub TimerQR_Timer()
'On Error GoTo hell
'
'   Set cQrCode = New ClassQR
'   Picture1.Picture = cQrCode.GetPictureQrCode(lblkdbarang, Picture1.ScaleWidth, Picture1.ScaleHeight)
'   If Picture1.Picture Is Nothing Then MsgBox "Error!"
'
'   TimerQR.Interval = 0
'
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !"
'TimerQR.Interval = 0

End Sub

Private Sub txtBFS_Change()
Call nul(txtBFS)
End Sub

Private Sub txtBFS_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtBFS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtBFS_KeyPress(KeyAscii As Integer)
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

Private Sub txtBFS_LostFocus()
On Error GoTo hell

txtBFS = FormatNumber(txtBFS, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtBFS.SetFocus

End Sub

Private Sub TXTKD1_Change()
If lblkdkategori = "04" Or lblkdkategori = "10" Then
Call nul(TXTKD1)
End If
End Sub

Private Sub txtkd1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtkd1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtkd1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtkd1_LostFocus()
TXTKD1 = UCase(TXTKD1)
End Sub

Private Sub txtkdSAP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtkdSAP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtkdSAP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtkdSAP_LostFocus()
txtkdSAP = UCase(txtkdSAP)
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

Private Sub txtmerk_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtmerk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtmerk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtmerk_LostFocus()
txtmerk = UCase(txtmerk)
End Sub

Private Sub TXTnmbarang_Change()
Call nul(TXTnmbarang)
End Sub

Private Sub TXTnmbarang_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTnmbarang_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub TXTnmBARANG_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub TXTnmbarang_LostFocus()
TXTnmbarang = UCase(TXTnmbarang)
End Sub


Private Sub txtsatuan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtsatuan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtsatuan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtsatuan_LostFocus()
txtsatuan = UCase(txtsatuan)
End Sub

Private Sub txttglNon_aktif_Change()
Call nul(txttglnon_aktif)
End Sub

Private Sub txttglNon_aktif_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglNon_aktif_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglNon_aktif_KeyPress(KeyAscii As Integer)
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

Private Sub txttglNon_aktif_LostFocus()
On Error GoTo hell

txttglnon_aktif = FormatDateTime(txttglnon_aktif, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglnon_aktif.SetFocus

End Sub

