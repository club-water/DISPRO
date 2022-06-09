VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form Kwitansi_cetak 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17100
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   17100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerPDF 
      Left            =   16200
      Top             =   3465
   End
   Begin VB.Timer Timerxls 
      Left            =   16470
      Top             =   4875
   End
   Begin VB.Timer Timerrtf 
      Left            =   16245
      Top             =   4095
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   0
      Top             =   675
      Width           =   15630
      _Version        =   524288
      _ExtentX        =   27570
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   930
      Left            =   15750
      TabIndex        =   3
      Top             =   5805
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1640
      _Version        =   262144
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Kwitansi_cetak.frx":0000
      AutoSize        =   1
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   900
      Left            =   15750
      TabIndex        =   4
      Top             =   6750
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1588
      _Version        =   262144
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Kwitansi_cetak.frx":49F2
      AutoSize        =   1
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   870
      Left            =   16155
      TabIndex        =   5
      Top             =   1485
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      _Version        =   262144
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Kwitansi_cetak.frx":93E4
      AutoSize        =   1
      Alignment       =   5
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   13635
      TabIndex        =   6
      Top             =   855
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   262144
      ForeColor       =   16711680
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Kwitansi_cetak.frx":3CCB2
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   9060
      Left            =   450
      TabIndex        =   2
      Top             =   765
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   15981
      SectionData     =   "Kwitansi_cetak.frx":43514
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   9945
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
      Picture         =   "Kwitansi_cetak.frx":43550
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Kwitansi"
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
      Left            =   1215
      TabIndex        =   1
      Top             =   0
      Width           =   7395
   End
   Begin VB.Image Image1 
      Height          =   10365
      Left            =   0
      Picture         =   "Kwitansi_cetak.frx":49DB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17070
   End
End
Attribute VB_Name = "Kwitansi_cetak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim kode As Integer
Dim rsmax As ADODB.Recordset

Dim color As Long, flag As Byte

Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdfs_Click()
AR_Kwitansi.Show vbModal
End Sub

Private Sub cmdfs_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdrtf_Click()
Timerrtf.Interval = 10
End Sub

Private Sub cmdrtf_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdxls_Click()
Timerxls.Interval = 10
End Sub


Private Sub cmdxls_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdPdf_Click()
TimerPDF.Interval = 10
End Sub

Private Sub cmdPdf_KeyPress(KeyAscii As Integer)
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



Private Sub hps()
'On Error GoTo hell
'kode = 3
'Call max
'    ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
'    If ms = vbYes Then
'        sql = "delete from gudang where kdgudang='" & rs!kdgudang & "' "
'        con.Execute (sql)
'
'        TimerAll.Interval = 10
'    Else
'        Exit Sub
'    End If
'
'
'Exit Sub
'hell:
'MsgBox err.Description
End Sub


Private Sub all()
If txtcari = "" Then
sql = "select * from Gudang  order by nmgudang"
Else
sql = "select * from gudang where " & kategori & " like '%" & txtcari & "%' order by nmgudang"
End If

Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs

Call LG
End Sub





Private Sub Form_Load()
GradientForm Me, 0
End Sub

Private Sub TimerPDF_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF

out2 = out2 + 1

pdf.FileName = App.Path & "\outfile" & CStr(out2) & ".pdf"
pdf.Export AR_Kwitansi.Pages

Call EX_PDF(Me)
TimerPDF.Interval = 0

Exit Sub
hell:
TimerPDF.Interval = 0
If out2 < 10 Then
cmdPdf_Click
End If

End Sub

Private Sub Timerrtf_Timer()
On Error GoTo hell
Dim rtf As New ActiveReportsRTFExport.ARExportRTF
out = out + 1


rtf.FileName = App.Path & "\outfile" & CStr(out) & ".rtf"
rtf.Export ARV1.Pages

Call EX_WORD(Me)
Timerrtf.Interval = 0

Exit Sub
hell:
Timerrtf.Interval = 0
If out < 10 Then
cmdrtf_Click
End If
End Sub

Private Sub Timerxls_Timer()
On Error GoTo hell
Dim xls As New ActiveReportsExcelExport.ARExportExcel

out1 = out1 + 1

xls.FileName = App.Path & "\outfile" & CStr(out1) & ".xls"
xls.Export ARV1.Pages

Call EX_EXEL(Me)
Timerxls.Interval = 0

Exit Sub
hell:
Timerxls.Interval = 0
If out1 < 10 Then
cmdxls_Click
End If
End Sub

