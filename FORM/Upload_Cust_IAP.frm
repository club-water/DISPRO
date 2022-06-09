VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Upload_Cust_IAP 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   330
      Left            =   720
      TabIndex        =   0
      Top             =   1395
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   915
      Left            =   4725
      TabIndex        =   1
      ToolTipText     =   "Dailly Close"
      Top             =   810
      Width           =   870
      _ExtentX        =   1535
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
      Picture         =   "Upload_Cust_IAP.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   2
      Top             =   540
      Width           =   4590
      _Version        =   524288
      _ExtentX        =   8096
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA YG DIUPLOAD DI SERVER :"
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
      Height          =   375
      Left            =   45
      TabIndex        =   5
      Top             =   630
      Width           =   4380
   End
   Begin VB.Label lbltglOD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "D:\DISPRO\UPLOAD\CUSTOMER_IAP.XLSX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   4650
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Customer IAP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   4290
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   4950
      Picture         =   "Upload_Cust_IAP.frx":294E
      Stretch         =   -1  'True
      Top             =   90
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   0
      Picture         =   "Upload_Cust_IAP.frx":2D0E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5625
   End
End
Attribute VB_Name = "Upload_Cust_IAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim i As Integer
Dim ms As VbMsgBoxResult
Dim filename As String
Dim color As Long, flag As Byte


Private Sub bar(a As Long, b As Long)
For i = a To b
Bar1.Value = i
Next
End Sub



Private Sub cmdOK_Click()
On Error GoTo hell

MousePointer = vbHourglass

con.Execute ("exec sp_ins_customer_IAP")
Call bar(0, 100)
MsgBox "Upload data berhasil", vbInformation, "Informasi !"
MousePointer = vbDefault

Unload Me



Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
Bar1.Value = 0
MousePointer = vbDefault
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
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

End Sub


