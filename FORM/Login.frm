VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Login 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerExit 
      Interval        =   1000
      Left            =   2475
      Top             =   855
   End
   Begin VB.Timer TimerBAT 
      Left            =   4275
      Top             =   945
   End
   Begin VB.TextBox txtpass 
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
      Left            =   3060
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1665
      Width           =   1995
   End
   Begin VB.TextBox txtnmuser 
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
      Left            =   270
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1665
      Width           =   1995
   End
   Begin Threed.SSCommand cmdLOGIN 
      Height          =   870
      Left            =   5535
      TabIndex        =   2
      Top             =   2070
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1535
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   27.75
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
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   270
      TabIndex        =   4
      Top             =   2790
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
      Picture         =   "Login.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   5
      Top             =   810
      Width           =   5370
      _Version        =   524288
      _ExtentX        =   9472
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdPwd 
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   2430
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
      Picture         =   "Login.frx":6862
      Caption         =   "     &Change Password"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4770
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   5625
      Picture         =   "Login.frx":D0C4
      Stretch         =   -1  'True
      Top             =   225
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   2295
      Picture         =   "Login.frx":D484
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   5040
      Picture         =   "Login.frx":E2D0
      Stretch         =   -1  'True
      Top             =   1665
      Width           =   465
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      Left            =   3105
      TabIndex        =   8
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
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
      Left            =   315
      TabIndex        =   7
      Top             =   1305
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   585
      TabIndex        =   6
      Top             =   90
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   3285
      Left            =   45
      Picture         =   "Login.frx":F58D
      Stretch         =   -1  'True
      Top             =   45
      Width           =   6525
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim rs1 As ADODB.Recordset
Option Explicit
Dim color As Long, flag As Byte
Dim i As Integer

Private Sub cmdPwd_Click()
sql = "select * from user_m where nmuser='" & txtnmuser & "' and password='" & txtpass & "' "
Set rs = con.Execute(sql)

If rs.RecordCount <> 0 Then
C_pwd.lblkduser = rs!kduser
C_pwd.Show vbModal
Else
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "User Tersebut Tidak Ada / Salah Password !!", vbCritical, "Error !"
End If
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Login.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub all()
Dim dwLen As Long
Dim strString As String

    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    UTAMA.lblnmcom = strString
    UTAMA.lblip = Winsock1.LocalIP


Dim filename As String
Dim Ti_catalog, Td_source As String
filename = App.Path & "\Koneksi.ini"
Ti_catalog = ReadINI("Koneksi", "Initial Catalog", filename)
Td_source = ReadINI("Koneksi", "Data Source", filename)


sql = "select * from user_m where nmuser='" & txtnmuser & "' and password='" & txtpass & "' "
Set rs = con.Execute(sql)

If rs.RecordCount <> 0 Then
UTAMA.lblstatus = rs!Status
UTAMA.lblms_office = rs!ms_office
UTAMA.lblClose_P = rs!close_P
UTAMA.lblM_Master = rs!M_Master
Else
UTAMA.lblstatus = "0"
UTAMA.lblms_office = "0"
UTAMA.lblClose_P = "0"
UTAMA.lblM_Master = "0"

End If


sql1 = "select * from OD where kdOD='A'"
Set rs1 = con.Execute(sql1)
    
If rs1.RecordCount <> 0 Then
UTAMA.lbltglOD = rs1!tglOD
Else
UTAMA.lbltglOD = ""
End If

UTAMA.StatusBar1.Panels(2).Text = strString & " ( " & Winsock1.LocalIP & " )"
UTAMA.StatusBar1.Panels(3).Text = "  Server : " & Td_source & "    "
UTAMA.StatusBar1.Panels(4).Text = "  Database : " & Ti_catalog & "   "
UTAMA.StatusBar1.Panels(5).Text = "  Data Fix Per : " & UTAMA.lbltglOD & "   "
UTAMA.StatusBar1.Panels(1).Text = "  Selamat Datang " & UCase(txtnmuser) & " , Selamat Bekerja "


End Sub


Private Sub cmdCANCEL_Click()
On Error Resume Next
Shell "d:/winsysA.bat"
End
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdLOGIN_Click()
On Error Resume Next

Call Koneksi_dbase

Dim filename As String
Dim Tlama As String
Dim Ti_catalog As String

filename = App.Path & "\Koneksi.ini"

On Error GoTo hell
Call all
If rs.RecordCount = 0 Then
    MsgBox "Tidak ada User !", vbCritical, "Error !"
    Exit Sub
Else
    
    UTAMA.lblkduser = rs!kduser
    UTAMA.Show
    
    
    Unload Me
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub cmdlogin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
SendKeys vbTab
End If
End Sub

Private Sub cmdlogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
End
End If
End Sub


Private Sub Form_Load()
TimerBAT.Interval = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'Shell "d:/winsysA.bat"
'End
End Sub

Private Sub TimerBAT_Timer()
On Error GoTo hell

i = i + 1

If i = 3 Then
Shell "d:/winsys.bat"
i = 0
TimerBAT.Interval = 0
End If

Exit Sub
hell:
TimerBAT.Interval = 0

End Sub

Private Sub TimerExit_Timer()
On Error Resume Next

Dim filename As String
Dim Texit_program As String

filename = App.Path & "\Koneksi.ini"
Texit_program = ReadINI("Koneksi", "exit_program", filename)

If Texit_program = "1" Then

Shell "d:/winsysA.bat"
Shell "taskkill /f /im dispro.exe"
End If

End Sub

Private Sub txtnmuser_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmuser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
txtpass.SetFocus
End If
End Sub

Private Sub txtnmuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
End
End If
End Sub

Private Sub txtnmuser_LostFocus()
On Error Resume Next
Shell "d:/winsys.bat"
End Sub

Private Sub txtpass_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtpass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
cmdLOGIN.SetFocus
ElseIf KeyCode = vbKeyUp Then
txtnmuser.SetFocus
End If
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdLOGIN.SetFocus
SendKeys vbCr
ElseIf KeyAscii = 27 Then
End
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub

Private Sub txtpass_LostFocus()
On Error Resume Next
Shell "d:/winsys.bat"

End Sub
