VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Supplier_TU 
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttelp 
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
      TabIndex        =   3
      Top             =   2295
      Width           =   4605
   End
   Begin VB.TextBox TXTnmsupplier 
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
   Begin VB.TextBox txtalamat 
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
      Width           =   4605
   End
   Begin VB.TextBox lblkdsupplier 
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
      Width           =   1680
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
      TabIndex        =   4
      Top             =   2655
      Width           =   4605
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   90
      TabIndex        =   7
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
      Left            =   540
      TabIndex        =   6
      Top             =   3195
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
      Picture         =   "Supplier_TU.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   1095
      Left            =   6075
      TabIndex        =   5
      ToolTipText     =   "Simpan"
      Top             =   2070
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1931
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
      Picture         =   "Supplier_TU.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label lblfrm 
      Height          =   330
      Left            =   2340
      TabIndex        =   15
      Top             =   3915
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   6210
      Picture         =   "Supplier_TU.frx":92CF
      Stretch         =   -1  'True
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER :"
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
      TabIndex        =   14
      Top             =   1620
      Width           =   1320
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
      Left            =   180
      TabIndex        =   13
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   495
      TabIndex        =   12
      Top             =   4185
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT :"
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
      TabIndex        =   11
      Top             =   1980
      Width           =   1320
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
      Left            =   180
      TabIndex        =   10
      Top             =   2700
      Width           =   1320
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NO TELP :"
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
      TabIndex        =   9
      Top             =   2340
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Supplier"
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
      Left            =   675
      TabIndex        =   8
      Top             =   0
      Width           =   3525
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   3105
      Picture         =   "Supplier_TU.frx":968F
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   3660
      Left            =   -45
      Picture         =   "Supplier_TU.frx":A94C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "Supplier_TU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim a As Integer

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




Private Sub cmdsimpan_Click()
On Error GoTo hell

    If TXTnmsupplier = "" Or lblkdsupplier = "" Then
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !!"
    Exit Sub
    Else

         If LBLKODE = 1 Then
             sql = "insert into Supplier  values ('" & UCase(lblkdsupplier) & "','" & UCase(TXTnmsupplier) & "','" & UCase(txtalamat) & "','" & UCase(txttelp) & "','" & UCase(txtketerangan) & "')"
             con.Execute (sql)
             MsgBox "Data Telah Tersimpan", vbInformation, "Informasi !"
     

         Else
             sql = "update supplier set nmsupplier='" & UCase(TXTnmsupplier) & "',alamat='" & UCase(txtalamat) & "',keterangan='" & UCase(txtketerangan) & "',telp='" & UCase(txttelp) & "' where kdsupplier='" & lblkdsupplier & "'"
             con.Execute (sql)
             MsgBox "Data Telah di Ubah", vbInformation, "Informasi !"

         End If
         
         If lblfrm = "SUPPLIER_BR" Then
         Supplier_BR.TimerALL.Interval = 10
         Supplier_BR.txtcari = lblkdsupplier
         Else
         Supplier.TimerALL.Interval = 10
         End If
         
         Unload Me
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

Call nul(lblkdsupplier)
Call nul(TXTnmsupplier)
End Sub


Private Sub lblkdsupplier_Change()
Call nul(lblkdsupplier)
End Sub

Private Sub lblkdsupplier_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub lblkdsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub lblkdsupplier_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub lblkdsupplier_LostFocus()
lblkdsupplier = UCase(lblkdsupplier)
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

Private Sub txtnmsupplier_Change()
Call nul(TXTnmsupplier)
End Sub

Private Sub txtnmsupplier_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnmsupplier_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnmsupplier_LostFocus()
TXTnmsupplier = UCase(TXTnmsupplier)
End Sub


Private Sub txtalamat_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtalamat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtalamat_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtalamat_LostFocus()
txtalamat = UCase(txtalamat)
End Sub

Private Sub txttelp_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttelp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttelp_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txttelp_LostFocus()
txttelp = UCase(txttelp)
End Sub
