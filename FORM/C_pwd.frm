VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form C_pwd 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpwd1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   225
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1215
      Width           =   1995
   End
   Begin VB.TextBox txtpwd2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   225
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   1995
   End
   Begin Threed.SSCommand cmdLOGIN 
      Height          =   510
      Left            =   3105
      TabIndex        =   2
      Top             =   2070
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   900
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " &GO"
      ButtonStyle     =   4
      PictureAlignment=   3
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   90
      TabIndex        =   3
      Top             =   675
      Width           =   2940
      _Version        =   524288
      _ExtentX        =   5186
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin VB.Label lblkduser 
      Caption         =   "lblkduser"
      Height          =   375
      Left            =   675
      TabIndex        =   7
      Top             =   3285
      Width           =   870
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   2295
      Picture         =   "C_pwd.frx":0000
      Stretch         =   -1  'True
      Top             =   2115
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Top             =   855
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Left            =   225
      TabIndex        =   4
      Top             =   1800
      Width           =   2085
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2250
      Picture         =   "C_pwd.frx":12BD
      Stretch         =   -1  'True
      Top             =   1215
      Width           =   465
   End
   Begin VB.Image Image4 
      Height          =   435
      Left            =   3105
      Picture         =   "C_pwd.frx":257A
      Stretch         =   -1  'True
      Top             =   180
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "C_pwd.frx":293A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3690
   End
End
Attribute VB_Name = "C_pwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim color As Long, flag As Byte

Private Sub cmdLOGIN_Click()
If txtpwd1 = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Password tidak boleh kosong !!", vbCritical, "Error !"
    Exit Sub
ElseIf txtpwd1 <> txtpwd2 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Password tidak boleh kosong !!", vbCritical, "Error !"
    Exit Sub
ElseIf txtpwd1 = txtpwd2 Then
    con.Execute ("Update user_m set [password]='" & txtpwd1 & "' where kduser='" & lblkduser & "' ")
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Password Berhasil diubah !!", vbInformation, "Info !"
    
    Unload Me
End If
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub txtpwd1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPWD1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub


Private Sub txtpwd1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtpwd2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPWD2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtpwd2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdLOGIN.SetFocus
SendKeys vbCr
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub
