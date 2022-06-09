VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form ACekher_TU 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3525
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlama 
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
      Left            =   4275
      TabIndex        =   4
      Text            =   "0"
      Top             =   2385
      Width           =   870
   End
   Begin VB.Timer TimerNO 
      Left            =   3240
      Top             =   765
   End
   Begin VB.TextBox txtnmarea 
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
      TabIndex        =   1
      Top             =   1575
      Width           =   5010
   End
   Begin VB.TextBox lblkdarea 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      TabIndex        =   0
      Top             =   1215
      Width           =   1320
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   915
      Left            =   6210
      TabIndex        =   5
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
      Picture         =   "ACekher_TU.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   9
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
      TabIndex        =   6
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
      Picture         =   "ACekher_TU.frx":2A6D
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   4860
      TabIndex        =   2
      ToolTipText     =   "Simpan"
      Top             =   1890
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
      Picture         =   "ACekher_TU.frx":92CF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC1 
      Height          =   420
      Left            =   5355
      TabIndex        =   3
      Top             =   1890
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
      Picture         =   "ACekher_TU.frx":BB01
      ButtonStyle     =   4
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "HARI"
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
      Left            =   5175
      TabIndex        =   16
      Top             =   2430
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LAMA CEK:"
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
      Left            =   3375
      TabIndex        =   15
      Top             =   2430
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
      Left            =   1935
      TabIndex        =   7
      Top             =   1935
      Width           =   2940
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
      Left            =   1035
      TabIndex        =   8
      Top             =   1935
      Width           =   870
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DEFAULT CHEKHER :"
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
      TabIndex        =   14
      Top             =   1890
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2385
      Picture         =   "ACekher_TU.frx":E14B
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Area Cheker"
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
      TabIndex        =   13
      Top             =   0
      Width           =   4605
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   945
      TabIndex        =   12
      Top             =   4275
      Width           =   1545
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
      Left            =   225
      TabIndex        =   11
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "AREA  :"
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
      TabIndex        =   10
      Top             =   1620
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   6255
      Picture         =   "ACekher_TU.frx":F408
      Stretch         =   -1  'True
      Top             =   135
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   3480
      Left            =   0
      Picture         =   "ACekher_TU.frx":F7C8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "ACekher_TU"
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


Private Sub nomer()
On Error GoTo hell

sql = "Select isnull(max(right(kdareaC,4)),0) as xx from Area_cheker"
Set rs = con.Execute(sql)


        a = CInt(rs!xx) + 1
                
        Select Case Len(CStr(a))
        Case 1
           lblkdarea = "AC000" & (a)
        Case 2
           lblkdarea = "AC00" & (a)
        Case 3
           lblkdarea = "AC0" & (a)
        Case 4
           lblkdarea = "AC" & (a)
        
        End Select

Exit Sub
hell:
lblkdarea = "AC01"

End Sub


Private Sub cmdBR1_Click()
Teknisi_BR.LBLKODE = "ACEKHER_TU"
Teknisi_BR.Show vbModal

End Sub



Private Sub cmdC1_Click()
lblkdteknisi = ""
lblnmteknisi = ""
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
On Error GoTo hell

    If txtnmarea = "" Or lblkdarea = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !"
    Exit Sub
    Else

         If LBLKODE = 1 Then
             sql = "insert into area_Cheker  values ('" & UCase(lblkdarea) & "','" & UCase(txtnmarea) & "','" & lblkdteknisi & "'," & CInt(txtlama) & ")"
             con.Execute (sql)
             SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
             MsgBox "Data Telah Tersimpan", vbInformation, "Info !"
     
             ACekher.TimerALL.Interval = 10
         Else
             sql = "update area_Cheker set nmareaC='" & UCase(txtnmarea) & "',kdteknisi='" & lblkdteknisi & "',lama_cek=" & CInt(txtlama) & "  where kdareaC='" & lblkdarea & "'"
             con.Execute (sql)
             SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
             MsgBox "Data Telah di Ubah", vbInformation, "Info !"

             ACekher.TimerALL.Interval = 10

         End If
         
         Unload Me
    End If
Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

TimerNO.Interval = 10

If UTAMA.lblstatus = 0 Then
txtlama.Enabled = False
Else
txtlama.Enabled = True
End If

Call nul(lblkdarea)
Call nul(txtnmarea)
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




Private Sub txtalamat_Change()

End Sub



Private Sub lblkdarea_Change()
Call nul(lblkdarea)
End Sub

Private Sub lblkdarea_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub lblkdarea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub lblkdarea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub lblkdarea_LostFocus()
lblkdarea = UCase(lblkdarea)
End Sub







Private Sub TimerNO_Timer()
On Error GoTo hell

If LBLKODE = "1" Then
Call nomer

TimerNO.Interval = 0
End If


Exit Sub
hell:
TimerNO.Interval = 0
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub txtnmarea_Change()
Call nul(txtnmarea)
End Sub

Private Sub txtnmarea_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmarea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txtnmarea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnmarea_LostFocus()
txtnmarea = UCase(txtnmarea)
End Sub




Private Sub txtlama_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtlama_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtlama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii <> vbKeyBack Then

    cekTBL = InStr("1234567890.,-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txtlama_LostFocus()
On Error GoTo hell

txtlama = FormatNumber(txtlama, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtlama.SetFocus

End Sub


Private Sub txtunit_Change()

End Sub
