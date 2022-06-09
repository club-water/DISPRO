VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form SPIAP_TU 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTnmSP 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1575
      Width           =   4605
   End
   Begin VB.Timer TimerCMB 
      Left            =   4545
      Top             =   1035
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1935
      Width           =   4605
   End
   Begin VB.TextBox txtnoSP 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1215
      Width           =   1365
   End
   Begin VB.ComboBox CMBcabang 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2295
      Width           =   3525
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   1095
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "Simpan"
      Top             =   2160
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
      Picture         =   "SPIAP_TU.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   6
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
      Left            =   675
      TabIndex        =   7
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
      Picture         =   "SPIAP_TU.frx":2A6D
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label lblkdSP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   810
      TabIndex        =   15
      Top             =   2790
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblkdcabang 
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
      Left            =   4995
      TabIndex        =   4
      Top             =   2295
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "STOK POINT :"
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
      Caption         =   "KODE SP :"
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
      Width           =   1005
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   540
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
      Left            =   225
      TabIndex        =   11
      Top             =   1980
      Width           =   1320
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CABANG IAP :"
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
      Top             =   2340
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail SP IAP"
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
      TabIndex        =   9
      Top             =   0
      Width           =   3525
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2835
      Picture         =   "SPIAP_TU.frx":92CF
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   465
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   6255
      Picture         =   "SPIAP_TU.frx":A58C
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Label lblfrm 
      Height          =   285
      Left            =   2925
      TabIndex        =   8
      Top             =   4230
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   3705
      Left            =   0
      Picture         =   "SPIAP_TU.frx":A94C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "SPIAP_TU"
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



Private Sub Cmbcabang_Click()
On Error Resume Next
sql1 = "select * from cabang_IAP where nmCabang='" & CMBcabang.Text & "'"
Set rs1 = con.Execute(sql1)

lblkdcabang = rs1!kdcabang

End Sub

Private Sub Cmbcabang_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub cmdsimpan_Click()
On Error GoTo hell

    If txtnoSP = "" Or TXTnmSP = "" Then
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !!"
    Exit Sub
    Else

         If lblkode = 1 Then
             sql = "insert into SP_IAP values ('" & UCase(txtnoSP) & "','" & UCase(TXTnmSP) & "','" & UCase(txtalamat) & "','" & UCase(lblkdcabang) & "/" & UCase(txtnoSP) & "','" & lblkdcabang & "')"
             con.Execute (sql)
             MsgBox "Data Telah Tersimpan", vbInformation, "Informasi !"
     
             
         Else
             sql = "update SP_IAP set nmSP='" & UCase(TXTnmSP) & "',alamat='" & UCase(txtalamat) & "',noSP='" & UCase(txtnoSP) & "',kdcabang='" & lblkdcabang & "',kdsp='" & UCase(lblkdcabang) & "/" & UCase(txtnoSP) & "' where kdSP='" & lblkdSP & "'"
             con.Execute (sql)
             MsgBox "Data Telah di Ubah", vbInformation, "Informasi !"

             
         End If
         
         
         SPIAP.TimerALL.Interval = 10
               
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

sql = "Select * from Cabang_IAP order by kdcabang"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
CMBcabang.AddItem rs!nmcabang
rs.MoveNext
Loop




TimerCMB.Interval = 10
End Sub

Private Sub lblkdpic_Click()

End Sub











Private Sub TimerNo_Timer()

End Sub




Private Sub TimerCMB_Timer()
If lblkode = "1" Then
CMBcabang.ListIndex = 0
End If


TimerCMB.Interval = 0
End Sub



Private Sub txtnoSP_Change()
Call nul(txtnoSP)
End Sub

Private Sub txtnoSP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoSP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnoSP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnoSP_LostFocus()
txtnoSP = UCase(txtnoSP)
End Sub

Private Sub txtnmSP_Change()
Call nul(TXTnmSP)
End Sub

Private Sub txtnmSP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmSP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnmSP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnmSP_LostFocus()
TXTnmSP = UCase(TXTnmSP)
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

