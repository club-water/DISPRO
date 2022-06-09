VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Piutang_DTU 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPPH23 
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
      Left            =   3375
      TabIndex        =   4
      Text            =   "0"
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Timer TimerNo 
      Left            =   4995
      Top             =   765
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
      Left            =   1305
      TabIndex        =   6
      Top             =   2430
      Width           =   5910
   End
   Begin VB.TextBox txtjmlbayar 
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
      Left            =   1035
      TabIndex        =   3
      Text            =   "0"
      Top             =   2070
      Width           =   1410
   End
   Begin VB.ComboBox CMBjenis 
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
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1350
      Width           =   1275
   End
   Begin VB.TextBox txtpotongan 
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
      Left            =   5805
      TabIndex        =   5
      Text            =   "0"
      Top             =   2070
      Width           =   1410
   End
   Begin VB.TextBox txttglbayar 
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
      Left            =   5175
      TabIndex        =   1
      Top             =   1350
      Width           =   1590
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   9
      Top             =   720
      Width           =   7170
      _Version        =   524288
      _ExtentX        =   12647
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
      TabIndex        =   8
      Top             =   3015
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
      Picture         =   "Piutang_DTU.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   4860
      TabIndex        =   2
      ToolTipText     =   "Simpan"
      Top             =   1665
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
      Picture         =   "Piutang_DTU.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   825
      Left            =   7515
      TabIndex        =   7
      ToolTipText     =   "Simpan"
      Top             =   2115
      Width           =   780
      _ExtentX        =   1376
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Piutang_DTU.frx":9094
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PPH 23 :"
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
      Left            =   2565
      TabIndex        =   25
      Top             =   2115
      Width           =   780
   End
   Begin VB.Label lblkdbyrpiutang 
      Height          =   330
      Left            =   3960
      TabIndex        =   24
      Top             =   4365
      Width           =   2040
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SISA PIUTANG :"
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
      TabIndex        =   23
      Top             =   1035
      Width           =   1230
   End
   Begin VB.Label lblsisa_awal 
      Alignment       =   1  'Right Justify
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
      Left            =   1350
      TabIndex        =   22
      Top             =   990
      Width           =   1590
   End
   Begin VB.Label Label6 
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
      Left            =   90
      TabIndex        =   21
      Top             =   2475
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "JML BAYAR :"
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
      TabIndex        =   20
      Top             =   2115
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "KOLEKTOR :"
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
      TabIndex        =   19
      Top             =   1755
      Width           =   1005
   End
   Begin VB.Label lblkdkolektor 
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
      Left            =   1035
      TabIndex        =   18
      Top             =   1710
      Width           =   870
   End
   Begin VB.Label lblnmkolektor 
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
      Left            =   1935
      TabIndex        =   17
      Top             =   1710
      Width           =   2940
   End
   Begin VB.Label lblurut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2017"
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
      Left            =   1035
      TabIndex        =   16
      Top             =   1350
      Width           =   645
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ANGS KE :"
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
      TabIndex        =   15
      Top             =   1395
      Width           =   780
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   675
      TabIndex        =   14
      Top             =   4500
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembayaran Piutang"
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
      Left            =   540
      TabIndex        =   13
      Top             =   0
      Width           =   5235
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   7515
      Picture         =   "Piutang_DTU.frx":BB01
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS BYR :"
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
      Left            =   1800
      TabIndex        =   12
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "POTONGAN :"
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
      Left            =   4770
      TabIndex        =   11
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL BAYAR :"
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
      Left            =   4185
      TabIndex        =   10
      Top             =   1395
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   3525
      Left            =   0
      Picture         =   "Piutang_DTU.frx":BEC1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8340
   End
End
Attribute VB_Name = "Piutang_DTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim sql As String
Dim color As Long, flag As Byte

Private Sub nomer()
Dim a As Long

On Error GoTo hell

sql = "select isnull(max(urut),0) as urut from byrpiutangsewa where kdpiutang='" & Piutang_D.txtkdPiutang & "' "
Set rs = con.Execute(sql)

'jika masih kosong diberi nomer 00001
If rs.RecordCount <> 0 Then
lblurut = rs!urut + 1
Else
lblurut = 0
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub



Private Sub CMBjenis_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If

End Sub

Private Sub cmdBR_Click()
Kolektor_BR.LBLKODE = "PIUTANG_DTU"
Kolektor_BR.Show vbModal

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

Private Sub cmdsimpan_Click()
On Error GoTo hell

    Call Cek_tglOD
    If CDate(txttglposting) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
        MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
        MousePointer = vbDefault
        Exit Sub
    ElseIf lblnmkolektor = "" Or lblkdkolektor = "" Or (txtjmlbayar = 0 And txtpotongan = 0) Then
    MsgBox "inputan belum lengkap !!", vbCritical, "Error !!"
    Exit Sub
    Else
    
        If CCur(txtjmlbayar) + CCur(txtpotongan) > CCur(lblsisa_awal) Then
            MsgBox "Jumlah Pembayaran Berlebihan !", vbCritical, "Error !!"
            Exit Sub
            
        Else

             If LBLKODE = 1 Then
                 sql = "insert into byrpiutangsewa values ('" & CStr(lblurut) & UCase(Piutang_D.txtkdPiutang) & "'," & lblurut & ",'" & Format(txttglbayar, "yyyy/MM/dd") & "','" & Piutang_D.lblkdcustomer & "','" & lblkdkolektor & "'," & CCur(txtjmlbayar) & "," & CCur(txtPPH23) & "," & CCur(txtpotongan) & ",'" & UCase(txtketerangan) & "','" & Piutang_D.txtkdPiutang & "'," & CMBjenis.ListIndex & ")"
                 con.Execute (sql)
                 MsgBox "Data Telah Tersimpan", vbInformation, "Informasi !"
         
                 Piutang_D.TimerALL.Interval = 10
                 Piutang.TimerALL.Interval = 10
             Else
                 sql = "update byrpiutangsewa set trf=" & CMBjenis.ListIndex & ",kdkolektor='" & lblkdkolektor & "',keterangan='" & UCase(txtketerangan) & "',jmlbayar=" & CCur(txtjmlbayar) & ",potongan=" & CCur(txtpotongan) & ",rpPPH23=" & CCur(txtPPH23) & ",tglbayar='" & Format(txttglbayar, "yyyy/MM/dd") & "' where kdbyrpiutang='" & lblkdbyrpiutang & "'"
                 con.Execute (sql)
                 
              
                 
                 MsgBox "Data Telah di Ubah", vbInformation, "Informasi !"
    
                 Piutang_D.TimerALL.Interval = 10
                 Piutang.TimerALL.Interval = 10
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

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttglbayar = Date


CMBjenis.AddItem "TUNAI"
CMBjenis.AddItem "TRANSFER"
CMBjenis.ListIndex = 0

Call nul(lblkdkolektor)
Call nul(lblnmkolektor)

TimerNO.Interval = 10
End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If

TimerNO.Interval = 0

End Sub

Private Sub txtjmlbayar_Change()
Call nul(txtjmlbayar)
On Error GoTo hell
lblrupiah = CCur(txtunit) * CCur(txtjmlbayar)
lblrupiah = FormatNumber(lblrupiah, 0)

Exit Sub
hell:
lblrupiah = 0

End Sub

Private Sub txtjmlbayar_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtjmlbayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txtjmlbayar_KeyPress(KeyAscii As Integer)
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

Private Sub txtjmlbayar_LostFocus()
On Error GoTo hell

txtjmlbayar = FormatNumber(txtjmlbayar, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtjmlbayar.SetFocus

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

Private Sub txtpotongan_Change()
Call nul(txtpotongan)
End Sub

Private Sub txtpotongan_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtpotongan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txtpotongan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If

End If

End Sub

Private Sub txtpotongan_LostFocus()
On Error GoTo hell

txtpotongan = FormatNumber(txtpotongan, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtpotongan.SetFocus

End Sub


Private Sub lblkdkolektor_Change()
Call nul(lblkdkolektor)
End Sub

Private Sub lblnmkolektor_Change()
Call nul(lblnmkolektor)
End Sub


Private Sub txtPPH23_Change()
Call nul(txtPPH23)
End Sub

Private Sub txtPPH23_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPPH23_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txtPPH23_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If

End If

End Sub

Private Sub txtPPH23_LostFocus()
On Error GoTo hell

txtPPH23 = FormatNumber(txtPPH23, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtPPH23.SetFocus

End Sub

Private Sub txttglbayar_Change()
Call nul(txttglbayar)



End Sub

Private Sub txttglbayar_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglbayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglbayar_KeyPress(KeyAscii As Integer)
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

Private Sub txttglbayar_LostFocus()
On Error GoTo hell

txttglbayar = FormatDateTime(txttglbayar, vbGeneralDate)

If CDate(Piutang_D.lbltglposting) > CDate(txttglbayar) Then
    MsgBox "Tgl Bayar Harus lebih besar dari pada tanggal Posting !", vbCritical, "Error !"
    txttglbayar.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglbayar.SetFocus

End Sub



