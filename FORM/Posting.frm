VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Posting 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglposting 
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
      TabIndex        =   2
      Top             =   1395
      Width           =   1590
   End
   Begin VB.TextBox txttahun 
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
      Left            =   3240
      TabIndex        =   1
      Text            =   "2017"
      Top             =   1395
      Width           =   960
   End
   Begin VB.ComboBox CMBBLN 
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
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1395
      Width           =   1005
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
      Left            =   675
      TabIndex        =   5
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
      Picture         =   "Posting.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   780
      Left            =   7515
      TabIndex        =   4
      ToolTipText     =   "Posting"
      Top             =   1755
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "Posting.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Left            =   7515
      TabIndex        =   3
      ToolTipText     =   "Cetak Bentuk List"
      Top             =   945
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
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
      Picture         =   "Posting.frx":972E
      ButtonStyle     =   4
   End
   Begin VB.Label lblbln 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   1980
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL POSTING :"
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
      Left            =   4365
      TabIndex        =   11
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN :"
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
      Left            =   2565
      TabIndex        =   10
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TAGIHAN BLN :"
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
      TabIndex        =   9
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   7515
      Picture         =   "Posting.frx":CAB4
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Tagihan Sewa"
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
      Left            =   540
      TabIndex        =   8
      Top             =   0
      Width           =   5235
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   675
      TabIndex        =   7
      Top             =   3735
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   0
      Picture         =   "Posting.frx":CE74
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8340
   End
End
Attribute VB_Name = "Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim sql As String
Dim color As Long, flag As Byte

Private Sub CMBBLN_Click()
Select Case cmbbln.ListIndex
    Case 0
        lblbln = "I"
    Case 1
        lblbln = "II"
    Case 2
        lblbln = "III"
    Case 3
        lblbln = "IV"
    Case 4
        lblbln = "V"
    Case 5
        lblbln = "VI"
    Case 6
        lblbln = "VII"
    Case 7
        lblbln = "VIII"
    Case 8
        lblbln = "IX"
    Case 9
        lblbln = "X"
    Case 10
        lblbln = "XI"
    Case 11
        lblbln = "XII"
End Select
 
End Sub

Private Sub CMBBLN_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If

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

MousePointer = vbHourglass

Call Cek_tglOD
If CDate(txttglposting) <= rstgl_OD!tglOD And UTAMA.lblstatus = 0 Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
    MsgBox "Data Tidak dapat Di Update, Sudah Fix Per Tgl " & rstgl_OD!tglOD, vbCritical, "Error !"
    MousePointer = vbDefault
    Exit Sub
Else
    ms = MsgBox("Pastikan data yg akan di POSTING BENAR!!...Apakah anda ingin POSTING Kwitansi ?", vbQuestion + vbYesNo, "Pertanyaan !")
    If ms = vbYes Then
        sql1 = "select kdcustomer,sum(unit) as unit from (" & vbCrLf & _
                   "select 'A' as kode,a.kdsewa,b.kdcustomer,a.kdbarang,a.unit from sewa_d a left join  sewa b on a.kdsewa=b.kdsewa where b.tglsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "'" & vbCrLf & _
                   "Union" & vbCrLf & _
                   "select 'B' as kode,a.kdsewa,b.kdcustomer,a.kdbarang,-sum(a.unit) as unit from Rsewa_d a left join Rsewa b on a.kdRsewa =b.kdRsewa" & vbCrLf & _
                   "where b.tglRsewa <='" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' group by a.kdsewa,b.kdcustomer,a.kdbarang" & vbCrLf & _
                   " ) a group by kdcustomer"
                   
        sql = "insert into piutangsewa select a.kdcustomer +'/' + '" & Posting.lblbln & "' + '/' + '" & CStr(Posting.txttahun) & "' as kdpiutang,a.kdcustomer," & CCur(Posting.cmbbln.Text) & " as bln," & CCur(Posting.txttahun) & " as tahun,a.unit,b.hrgsewa as harga,(a.unit * b.hrgsewa) as jmlpiutang,'" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' as tglposting,0,'" & Format(Posting.txttglposting, "yyyy/MM/dd") & "' as tglcetak from  " & vbCrLf & _
              "(" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer where a.unit<>0 "
        
        
        con.Execute (sql)
        
        MsgBox "Tagihan Sewa bulan " & cmbbln.Text & " tahun = " & txttahun & " berhasil di posting", vbInformation, "Info !!"
        Unload Me
    Else
        MousePointer = vbDefault
        Exit Sub
    End If
End If

MousePointer = vbDefault
Exit Sub
hell:
MousePointer = vbDefault
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdT_Click()
Kwitansi_LIST.lblfrm = "POSTING"
Kwitansi_LIST.Show vbModal
End Sub

Private Sub cmdT_KeyPress(KeyAscii As Integer)
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

txttglposting = Date
txttahun = Year(Date)

cmbbln.AddItem "1"
cmbbln.AddItem "2"
cmbbln.AddItem "3"
cmbbln.AddItem "4"
cmbbln.AddItem "5"
cmbbln.AddItem "6"
cmbbln.AddItem "7"
cmbbln.AddItem "8"
cmbbln.AddItem "9"
cmbbln.AddItem "10"
cmbbln.AddItem "11"
cmbbln.AddItem "12"
cmbbln.ListIndex = Month(Date) - 1

End Sub

Private Sub txttahun_Change()
Call nul(txttahun)
End Sub

Private Sub txttahun_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttahun_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttahun_KeyPress(KeyAscii As Integer)
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

Private Sub txttglposting_Change()
Call nul(txttglposting)
End Sub

Private Sub txttglposting_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglposting_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglposting_KeyPress(KeyAscii As Integer)
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

Private Sub txttglposting_LostFocus()
On Error GoTo hell

txttglposting = FormatDateTime(txttglposting, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglposting.SetFocus

End Sub

