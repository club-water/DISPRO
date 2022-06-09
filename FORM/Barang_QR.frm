VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Barang_QR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerQR 
      Left            =   180
      Top             =   225
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   2160
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   1260
      Width           =   2385
   End
   Begin Threed.SSCommand cmdT 
      Height          =   915
      Index           =   4
      Left            =   6570
      TabIndex        =   2
      ToolTipText     =   "Cetak"
      Top             =   3780
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Barang_QR.frx":0000
      ButtonStyle     =   4
   End
   Begin VB.Label lblkdbarang 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TMP/P/00000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   855
      TabIndex        =   1
      Top             =   4725
      Width           =   4515
   End
   Begin VB.Image Image1 
      Height          =   5460
      Left            =   0
      Picture         =   "Barang_QR.frx":3A5D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7665
   End
End
Attribute VB_Name = "Barang_QR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim color As Long, flag As Byte

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0


TimerQR.Interval = 10
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub TimerQR_Timer()
On Error GoTo hell

   Set cQrCode = New ClassQR
   Picture1.Picture = cQrCode.GetPictureQrCode(lblkdbarang, Picture1.ScaleWidth, Picture1.ScaleHeight)
   If Picture1.Picture Is Nothing Then MsgBox "Error!"
    
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
TimerQR.Interval = 0

End Sub
