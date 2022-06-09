VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form KLAIM_DTU 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5625
      TabIndex        =   0
      Top             =   990
      Width           =   1590
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
      Left            =   4050
      TabIndex        =   2
      Text            =   "0"
      Top             =   1350
      Width           =   1590
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
      Left            =   1305
      TabIndex        =   1
      Text            =   "0"
      Top             =   1350
      Width           =   1590
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
      TabIndex        =   3
      Top             =   1710
      Width           =   5910
   End
   Begin VB.Timer TimerNo 
      Left            =   5985
      Top             =   225
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   6
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
      Left            =   315
      TabIndex        =   5
      Top             =   2520
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
      Picture         =   "KLAIM_DTU.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   825
      Left            =   7515
      TabIndex        =   4
      ToolTipText     =   "Simpan"
      Top             =   1665
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
      Picture         =   "KLAIM_DTU.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
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
      Left            =   4635
      TabIndex        =   17
      Top             =   1035
      Width           =   1185
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
      Left            =   3015
      TabIndex        =   16
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   7515
      Picture         =   "KLAIM_DTU.frx":92CF
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembayaran Klaim"
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
      Left            =   405
      TabIndex        =   15
      Top             =   0
      Width           =   5235
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   675
      TabIndex        =   14
      Top             =   4500
      Width           =   1545
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
      Left            =   3060
      TabIndex        =   13
      Top             =   1035
      Width           =   780
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
      Left            =   3870
      TabIndex        =   12
      Top             =   990
      Width           =   645
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
      Left            =   135
      TabIndex        =   11
      Top             =   1395
      Width           =   1185
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
      TabIndex        =   10
      Top             =   1755
      Width           =   1185
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
      Left            =   1305
      TabIndex        =   9
      Top             =   990
      Width           =   1590
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SISA KLAIM :"
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
      Left            =   45
      TabIndex        =   8
      Top             =   1035
      Width           =   1050
   End
   Begin VB.Label lblkdbyrklaim 
      Height          =   330
      Left            =   3960
      TabIndex        =   7
      Top             =   4365
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   0
      Picture         =   "KLAIM_DTU.frx":968F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8340
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
      TabIndex        =   20
      Top             =   1710
      Width           =   2940
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
      TabIndex        =   19
      Top             =   1710
      Width           =   870
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
      TabIndex        =   18
      Top             =   1755
      Width           =   1005
   End
End
Attribute VB_Name = "KLAIM_DTU"
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

sql = "select isnull(max(urut),0) as urut from byrKlaim where kdKlaim='" & Klaim_D.txtkdklaim & "' "
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

    If txtjmlbayar = 0 And txtpotongan = 0 Then
    MsgBox "inputan belum lengkap !!", vbCritical, "Error !!"
    Exit Sub
    Else
    
        If CCur(txtjmlbayar) + CCur(txtpotongan) > CCur(lblsisa_awal) Then
            MsgBox "Jumlah Pembayaran Berlebihan !", vbCritical, "Error !!"
            Exit Sub
            
        Else

             If LBLKODE = 1 Then
                 sql = "insert into byrKlaim values ('" & CStr(lblurut) & UCase(Klaim_D.txtkdklaim) & "'," & lblurut & ",'" & Format(txttglbayar, "yyyy/MM/dd") & "','1900/01/01','" & Klaim_D.lblkdcustomer & "'," & CCur(txtjmlbayar) & "," & CCur(txtpotongan) & ",'" & UCase(txtketerangan) & "','" & Klaim_D.txtkdklaim & "',getdate(),'" & UTAMA.lblkduser & "')"
                 con.Execute (sql)
                 MsgBox "Data Telah Tersimpan", vbInformation, "Informasi !"
         
                 Klaim_D.TimerALL.Interval = 10
                 Klaim.TimerALL.Interval = 10
             Else
                 sql = "update byrKlaim set keterangan='" & UCase(txtketerangan) & "',jmlbayar=" & CCur(txtjmlbayar) & ",potongan=" & CCur(txtpotongan) & ",tglbayar='" & Format(txttglbayar, "yyyy/MM/dd") & "',tglinput=getdate(), kduser='" & UTAMA.lblkduser & "' where kdbyrKlaim='" & lblkdbyrklaim & "'"
                 con.Execute (sql)
                 
              
                 
                 MsgBox "Data Telah di Ubah", vbInformation, "Informasi !"
    
                 Klaim_D.TimerALL.Interval = 10
                 Klaim.TimerALL.Interval = 10
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



TimerNo.Interval = 10
End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If

TimerNo.Interval = 0

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

'If CDate(Klaim_D.lbltglklaim) > CDate(txttglbayar) Then
'    MsgBox "Tgl Bayar Harus lebih besar dari pada tanggal Klaim !", vbCritical, "Error !"
'    txttglbayar.SetFocus
'End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglbayar.SetFocus

End Sub




