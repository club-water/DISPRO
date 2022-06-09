VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Planning_kirim_TU 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3555
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglPK 
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
      Left            =   945
      TabIndex        =   0
      Top             =   1845
      Width           =   1590
   End
   Begin VB.TextBox txturaian 
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
      Left            =   945
      TabIndex        =   2
      Top             =   2565
      Width           =   5010
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   915
      Left            =   6210
      TabIndex        =   3
      ToolTipText     =   "Simpan"
      Top             =   1260
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
      Picture         =   "Planning_kirim_TU.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   4
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
      Left            =   495
      TabIndex        =   5
      Top             =   3060
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
      Picture         =   "Planning_kirim_TU.frx":2A6D
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   4770
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   2160
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
      Picture         =   "Planning_kirim_TU.frx":92CF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label lbljmlunit 
      Caption         =   "Label3"
      Height          =   285
      Left            =   6165
      TabIndex        =   19
      Top             =   3690
      Width           =   690
   End
   Begin VB.Label lblketerangan 
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1530
      TabIndex        =   18
      Top             =   1350
      Width           =   4515
   End
   Begin VB.Label lblkdPK 
      Caption         =   "lblkdpk"
      Height          =   285
      Left            =   4860
      TabIndex        =   17
      Top             =   3690
      Width           =   960
   End
   Begin VB.Label lblkdcustomer 
      BackStyle       =   0  'Transparent
      Caption         =   "C00005"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   135
      TabIndex        =   16
      Top             =   810
      Width           =   735
   End
   Begin VB.Label lblalamat 
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   135
      TabIndex        =   15
      Top             =   1080
      Width           =   5820
   End
   Begin VB.Label lblnmkategori 
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   135
      TabIndex        =   14
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label lblnmcustomer 
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   945
      TabIndex        =   13
      Top             =   810
      Width           =   5055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PLANNING KIRIM :"
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
      Height          =   420
      Left            =   90
      TabIndex        =   12
      Top             =   1800
      Width           =   960
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   6255
      Picture         =   "Planning_kirim_TU.frx":BB01
      Stretch         =   -1  'True
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "URAIAN  :"
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
      Top             =   2610
      Width           =   915
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   945
      TabIndex        =   10
      Top             =   4275
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planning Kirim"
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
      Width           =   4605
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SOPIR :"
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
      Height          =   465
      Left            =   135
      TabIndex        =   8
      Top             =   2250
      Width           =   825
   End
   Begin VB.Label lblkdteknisi 
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
      Left            =   945
      TabIndex        =   7
      Top             =   2205
      Width           =   870
   End
   Begin VB.Label lblnmteknisi 
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
      Left            =   1845
      TabIndex        =   6
      Top             =   2205
      Width           =   2940
   End
   Begin VB.Image Image1 
      Height          =   3480
      Left            =   0
      Picture         =   "Planning_kirim_TU.frx":BEC1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "Planning_kirim_TU"
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




Private Sub cmdBR1_Click()
Teknisi_BR.lblkode = "PLANNING_KIRIM_TU"
Teknisi_BR.Show vbModal

End Sub




Private Sub cmdCANCEL_Click()
Unload Me
End Sub




Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub




Private Sub cmdsimpan_Click()
'On Error GoTo hell

    If txttglPK = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !"
    Exit Sub
    Else

         If Planning_kirim.Opt1.Value = True Then
             sql = "insert into planning_kirim  values ('" & UCase(lblkdPK) & "','" & Format(txttglPK, "yyyy/MM/dd") & "','" & lblkdteknisi & "','" & lblkdcustomer & "'," & CInt(lbljmlunit) & ",'" & UCase(txturaian) & "')"
             con.Execute (sql)
             SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
             MsgBox "Data Telah Tersimpan", vbInformation, "Info !"
     
             Planning_kirim.TimerALL.Interval = 10
         Else
             sql = "update planning_kirim set uraian='" & UCase(txturaian) & "',kdteknisi='" & lblkdteknisi & "',tglpk='" & Format(txttglPK, "yyyy/MM/dd") & "'  where kdPK='" & lblkdPK & "'"
             con.Execute (sql)
             SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
             MsgBox "Data Telah di Ubah", vbInformation, "Info !"

             Planning_kirim.TimerALL.Interval = 10

         End If
         
         Unload Me
    End If
'Exit Sub
'hell:
'SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
'MsgBox err.Description, vbCritical, "Error !"

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0


Call nul(txttglPK)
Call nul(lblnmteknisi)
Call nul(lblkdteknisi)
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







Private Sub lblkdteknisi_Change()
Call nul(lblkdteknisi)
End Sub

Private Sub lblnmteknisi_Change()
Call nul(lblnmteknisi)
End Sub


Private Sub txttglPK_Change()
Call nul(txttglPK)
End Sub

Private Sub txttglPK_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglPK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglPK_KeyPress(KeyAscii As Integer)
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

Private Sub txttglPK_LostFocus()
On Error GoTo hell

txttglPK = FormatDateTime(txttglPK, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglPK.SetFocus

End Sub

Private Sub txtUraian_Change()
Call nul(txturaian)
End Sub

Private Sub txtUraian_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtUraian_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txtUraian_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtUraian_LostFocus()
txturaian = UCase(txturaian)
End Sub




