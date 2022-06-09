VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Dclose 
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglFIX 
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
      Left            =   2205
      TabIndex        =   0
      Top             =   855
      Width           =   1680
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   330
      Left            =   675
      TabIndex        =   2
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
      Left            =   4680
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
      Picture         =   "Dclose.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   0
      TabIndex        =   6
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
   Begin VB.Image Image4 
      Height          =   345
      Left            =   4905
      Picture         =   "Dclose.frx":352B
      Stretch         =   -1  'True
      Top             =   90
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dailly Close"
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
      Left            =   315
      TabIndex        =   5
      Top             =   0
      Width           =   2265
   End
   Begin VB.Label lbltglOD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "23/12/2017"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2115
      TabIndex        =   4
      Top             =   3690
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FIX PER :"
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
      Left            =   765
      TabIndex        =   3
      Top             =   900
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   -45
      Picture         =   "Dclose.frx":38EB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5625
   End
End
Attribute VB_Name = "Dclose"
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

Dim T_Dclose, Ti_catalog As String
filename = App.Path & "\Koneksi.ini"

T_Dclose = ReadINI("Koneksi", "Dclose", filename)
Ti_catalog = ReadINI("Koneksi", "Initial Catalog", filename)

Call Cek_tglOD
If CDate(txttglFIX) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    MousePointer = vbDefault
    Exit Sub
ElseIf T_Dclose = "" Then
    MsgBox "Tidak bisa di Closing, Karena Folder Closing blom diset", vbCritical, "Error !!"
    MousePointer = vbDefault
    Exit Sub

Else

    sql = "backup database " & Ti_catalog & " to disk = '" & T_Dclose & Format(txttglFIX, "ddMMyy") & ".Bak" & "' with format"
    con.Execute (sql)
    Call bar(0, 30)
    
    sql = "insert into history_closing values ('" & Format(CDate(txttglFIX), "yyyy/MM/dd") & "',getdate())"
    con.Execute (sql)
    Call bar(31, 60)
    
    sql = "update OD set tglOD = '" & Format(CDate(txttglFIX), "yyyy/MM/dd") & "'"
    con.Execute (sql)
    Call bar(61, 100)
    
    
    lbltglOD = CDate(txttglFIX)
    
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "Backup data berhasil", vbInformation, "Informasi !"
    End
    
    DoEvents

End If

MousePointer = vbDefault

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

Call Cek_tglOD
lbltglOD = rstgl_OD!tglOD
txttglFIX = Date
End Sub

Private Sub txttglfix_Change()
Call nul(txttglFIX)
End Sub

Private Sub txttglfix_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglfix_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglfix_KeyPress(KeyAscii As Integer)
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

Private Sub txttglfix_LostFocus()
On Error GoTo hell

txttglFIX = FormatDateTime(txttglFIX, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglFIX.SetFocus

End Sub

