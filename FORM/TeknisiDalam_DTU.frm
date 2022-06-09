VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form TeknisiDalam_DTU 
   BorderStyle     =   0  'None
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3225
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtunit 
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
      Left            =   1395
      TabIndex        =   1
      Text            =   "1"
      Top             =   1485
      Width           =   1140
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
      TabIndex        =   2
      Top             =   1845
      Width           =   6675
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   4
      Top             =   720
      Width           =   7980
      _Version        =   524288
      _ExtentX        =   14076
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
      Top             =   2565
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
      Picture         =   "TeknisiDalam_DTU.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   870
      Left            =   8235
      TabIndex        =   3
      ToolTipText     =   "Simpan"
      Top             =   1800
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
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
      Picture         =   "TeknisiDalam_DTU.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   7560
      TabIndex        =   0
      Top             =   1080
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
      Picture         =   "TeknisiDalam_DTU.frx":92CF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label lblkdTD_d 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2340
      TabIndex        =   15
      Top             =   3690
      Width           =   2040
   End
   Begin VB.Label lblnmbarang 
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
      Left            =   2925
      TabIndex        =   14
      Top             =   1125
      Width           =   4650
   End
   Begin VB.Label lblkdbarang 
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
      Left            =   1395
      TabIndex        =   13
      Top             =   1125
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPART :"
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
      TabIndex        =   12
      Top             =   1170
      Width           =   1320
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   675
      TabIndex        =   11
      Top             =   3735
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QTY  :"
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
      TabIndex        =   10
      Top             =   1530
      Width           =   870
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
      Left            =   135
      TabIndex        =   9
      Top             =   1890
      Width           =   1320
   End
   Begin VB.Label lblsatuan 
      BackStyle       =   0  'Transparent
      Caption         =   "SATUAN"
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
      Left            =   2610
      TabIndex        =   8
      Top             =   1530
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SparePart"
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
      TabIndex        =   7
      Top             =   0
      Width           =   3525
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   8325
      Picture         =   "TeknisiDalam_DTU.frx":BB01
      Stretch         =   -1  'True
      Top             =   135
      Width           =   555
   End
   Begin VB.Label lblform 
      Caption         =   "lblform"
      Height          =   330
      Left            =   4725
      TabIndex        =   6
      Top             =   3690
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   3165
      Left            =   0
      Picture         =   "TeknisiDalam_DTU.frx":BEC1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "TeknisiDalam_DTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim sql As String
Dim sqlST, sqlA1, sqlA2 As String
Dim rsST As ADODB.Recordset
Dim sql1 As String
Dim a As Integer
Dim color As Long, flag As Byte


Private Sub cmdBR_Click()
Barang_BR.lblkode = UCase("teknisidalam_dTU")
Barang_BR.Show vbModal
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




Private Sub set_cmbbrg()
On Error GoTo hell

sql = "Select * from kategoriBRG order by kdkategori"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
cmbkategori.AddItem rs!nmkategori
rs.MoveNext
Loop

If lblkode = "1" Then
cmbkategori.ListIndex = 0
End If


        
 
Exit Sub
hell:
MsgBox err.Description

End Sub


Private Sub cmdsimpan_Click()
On Error GoTo hell

MousePointer = vbHourglass

    If lblnmbarang = "" Or lblkdbarang = "" Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Inputan Belum Lengkap !", vbCritical, "Error !"
        Exit Sub
    ElseIf txtunit = 0 Then
        SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
        MsgBox "Inputan Qty Tidak Boleh 0 !", vbCritical, "Error !"
        Exit Sub
    Else
         
         If lblkode = 1 Then
         
                If lblform = "TEKNISIDALAM_D" Then
                   sql = "insert into teknisidalam_d values ('" & UCase(lblkdbarang) & "_" & UCase(TeknisiDalam_D.txtkdTD) & "','" & UCase(TeknisiDalam_D.txtkdTD) & "','" & UCase(lblkdbarang) & "'," & CCur(txtunit) & ",'" & UCase(txtketerangan) & "',getdate(),'" & UTAMA.lblkduser & "')"
                   con.Execute (sql)
                   SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
                   MsgBox "Data Telah Tersimpan", vbInformation, "Info !"
        
                   TeknisiDalam_D.TimerAll.Interval = 10
                ElseIf lblform = "TEKNISILUAR_D" Then
                   sql = "insert into teknisiluar_d values ('" & UCase(lblkdbarang) & "_" & UCase(TeknisiLuar_D.txtkdTL) & "','" & UCase(TeknisiLuar_D.txtkdTL) & "','" & UCase(lblkdbarang) & "'," & CCur(txtunit) & ",'" & UCase(txtketerangan) & "',getdate(),'" & UTAMA.lblkduser & "')"
                   con.Execute (sql)
                   SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
                   MsgBox "Data Telah Tersimpan", vbInformation, "Info !"
        
                   TeknisiLuar_D.TimerAll.Interval = 10
                
                End If
         Else
                If lblform = "TEKNISIDALAM_D" Then
                    sql = "update teknisidalam_d set qty=" & CCur(txtunit) & ",keterangan='" & UCase(txtketerangan) & "' where kdTD_D='" & lblkdTD_d & "'"
                    con.Execute (sql)
                    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
                    MsgBox "Data Telah di Ubah", vbInformation, "Info !"
        
                    TeknisiDalam_D.TimerAll.Interval = 10
                ElseIf lblform = "TEKNISILUAR_D" Then
                    sql = "update teknisiluar_d set qty=" & CCur(txtunit) & ",keterangan='" & UCase(txtketerangan) & "' where kdTL_D='" & lblkdTD_d & "'"
                    con.Execute (sql)
                    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
                    MsgBox "Data Telah di Ubah", vbInformation, "Info !"
        
                    TeknisiLuar_D.TimerAll.Interval = 10
                
                End If

         End If
         
         MousePointer = vbDefault
         Unload Me
    End If
Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"
MousePointer = vbDefault

End Sub

Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

Call nul(lblnmbarang)
Call nul(lblkdbarang)
End Sub





Private Sub lblkdbarang_Change()
Call nul(lblkdbarang)
End Sub

Private Sub lblnmbarang_Change()
Call nul(lblnmbarang)
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



Private Sub txtunit_Change()
Call nul(txtunit)
End Sub

Private Sub txtunit_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtunit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtunit_KeyPress(KeyAscii As Integer)
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

Private Sub txtunit_LostFocus()
On Error GoTo hell

txtunit = FormatNumber(txtunit, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtunit.SetFocus

End Sub



