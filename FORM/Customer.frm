VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Customer 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   18720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtR 
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
      Left            =   15570
      TabIndex        =   15
      Text            =   "100"
      Top             =   315
      Width           =   735
   End
   Begin VB.CheckBox ChkR 
      BackColor       =   &H00000000&
      Caption         =   "TAMPILKAN :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   13995
      MaskColor       =   &H00000000&
      TabIndex        =   14
      Top             =   315
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.Timer TimerG 
      Left            =   6165
      Top             =   4815
   End
   Begin VB.Timer TimerAll 
      Left            =   5625
      Top             =   4815
   End
   Begin VB.TextBox TXTCARI 
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
      Height          =   300
      Left            =   3375
      TabIndex        =   8
      Top             =   9720
      Width           =   2850
   End
   Begin VB.ComboBox CMBCARI 
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
      Height          =   345
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   9720
      Width           =   1860
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   270
      TabIndex        =   9
      Top             =   675
      Width           =   17340
      _Version        =   524288
      _ExtentX        =   30586
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   0
      Left            =   17775
      TabIndex        =   0
      ToolTipText     =   "Tambah"
      Top             =   1080
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16744576
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
      Picture         =   "Customer.frx":0000
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   1
      Left            =   17775
      TabIndex        =   1
      ToolTipText     =   "Ubah"
      Top             =   1890
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
      Picture         =   "Customer.frx":2C74
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   2
      Left            =   17775
      TabIndex        =   2
      ToolTipText     =   "Hapus"
      Top             =   2700
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
      Picture         =   "Customer.frx":5E71
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   3
      Left            =   17775
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   4320
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
      Picture         =   "Customer.frx":8F0A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   4
      Left            =   17775
      TabIndex        =   6
      ToolTipText     =   "Cari Data"
      Top             =   3510
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
      Picture         =   "Customer.frx":C086
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   8205
      Left            =   225
      TabIndex        =   5
      Top             =   945
      Width           =   17295
      _cx             =   30506
      _cy             =   14473
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   36
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Customer.frx":EFAC
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   4
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin Threed.SSCommand cmdT 
      Height          =   825
      Index           =   5
      Left            =   17775
      TabIndex        =   4
      ToolTipText     =   "Cetak Bentuk List"
      Top             =   5175
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1455
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
      Picture         =   "Customer.frx":F317
      ButtonStyle     =   4
   End
   Begin Threed.SSOption Opt1 
      Height          =   330
      Left            =   450
      TabIndex        =   17
      Top             =   675
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ALL Customer"
   End
   Begin Threed.SSOption Opt2 
      Height          =   330
      Left            =   1980
      TabIndex        =   18
      Top             =   675
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Support Dispenser dan Showcase"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RECORD"
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
      Left            =   16335
      TabIndex        =   16
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label lblpos 
      Caption         =   "lblpos"
      Height          =   195
      Left            =   225
      TabIndex        =   13
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA TIDAK ADA"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11790
      TabIndex        =   12
      Top             =   9900
      Width           =   2220
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   12600
      Picture         =   "Customer.frx":1269D
      Stretch         =   -1  'True
      Top             =   9405
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Customer"
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
      Left            =   1125
      TabIndex        =   11
      Top             =   0
      Width           =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   1260
      Top             =   9315
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Pencarian"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1485
      TabIndex        =   10
      Top             =   9360
      Width           =   4560
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6255
      Picture         =   "Customer.frx":18EEF
      Stretch         =   -1  'True
      Top             =   9675
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   17685
      Picture         =   "Customer.frx":25D9F
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   10305
      Left            =   0
      Picture         =   "Customer.frx":2615F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18645
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim kode As Integer
Dim rsmax As ADODB.Recordset
Dim sql1 As String

Dim color As Long, flag As Byte

Private Sub ChkR_Click()
TimerALL.Interval = 10

If ChkR.Value = 0 Then
txtR.Enabled = False
Else
txtR.Enabled = True
End If

End Sub

Private Sub ChkR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
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



'untuk set cursor pada saat dihapus
Private Sub max()
If rs.AbsolutePosition = 1 Then
lblpos = 1
Else
lblpos = CLng(rs.AbsolutePosition) - 1
End If
End Sub


Private Sub tbl()
If rs.RecordCount = 0 Then
    cmdT(1).Enabled = False
    cmdT(2).Enabled = False
    datagrid1.Enabled = False
    img1.Visible = True
    lbl1.Visible = True
Else
    cmdT(1).Enabled = True
    cmdT(2).Enabled = True
    datagrid1.Enabled = True
    img1.Visible = False
    lbl1.Visible = False
End If
End Sub


Private Sub LG()
On Error GoTo hell

Call tbl

Exit Sub
hell:
End Sub

Private Sub tbh()
Customer_TU.LBLKODE = 1
Customer_TU.Show vbModal
End Sub

Private Sub ubh()
On Error Resume Next

Customer_TU.LBLKODE = 2
lblpos = rs.AbsolutePosition
kode = 2

Customer_TU.lbltgldibuat = rs!tgldibuat
Customer_TU.lblkdcustomer = rs!kdcustomer
Customer_TU.TXTnmcustomer = rs!nmcustomer
Customer_TU.txtalamat = rs!alamat
Customer_TU.txtalamat_TGH = rs!alamat_tgh
Customer_TU.txthrgSewa = Format(rs!hrgSewa, "#,###0")
Customer_TU.txttelp = rs!telp
Customer_TU.txtketerangan = rs!keterangan
If rs!keterangan = "" Then
Customer_TU.Chkket.Value = 0
Else
Customer_TU.Chkket.Value = 1
End If


Customer_TU.lblkdwilayah = rs!kdwilayah
'Customer_TU.lblnmwilayah = rs!nmwilayah
Customer_TU.txtCP = rs!CP
Customer_TU.lblkdSP = rs!kdsp
Customer_TU.txtkdcustomer_IAP = rs!kdcustomer_IAP
Customer_TU.CMBbank.Text = rs!kdbank
Customer_TU.lblkdkolektor = rs!kdkolektor
Customer_TU.lblkdarea = rs!kdarea
Customer_TU.CMbJNSBYR.Text = rs!jnsbayar
Customer_TU.txtnospk = rs!noSPK
Customer_TU.txttglspk1 = rs!tglSPK1
Customer_TU.txttglspk2 = rs!tglSPK2
Customer_TU.txtcup = FormatNumber(rs!target_cup, 0)
Customer_TU.txtbtl = FormatNumber(rs!target_btl, 0)
Customer_TU.txtgln = FormatNumber(rs!target_gln, 0)
Customer_TU.ChkNA.Value = rs!non_aktif
Customer_TU.txtnoNPWP = rs!NPWP
Customer_TU.txtnmNPWP = rs!nmNPWP
Customer_TU.txtalamatNPWP = rs!alamatNPWP
Customer_TU.lblkdareaC = rs!kdareaC
Customer_TU.lblkdteknisi = rs!kdteknisi
Customer_TU.lblkdPIC = rs!kdpic
Customer_TU.chkpph23.Value = rs!pph23



If rs!pkp = 0 Then
Customer_TU.OPT1.Value = False
Customer_TU.Opt2.Value = True
Else
Customer_TU.OPT1.Value = True
Customer_TU.Opt2.Value = False
End If

Customer_TU.Show vbModal




End Sub

Private Sub hps()
On Error GoTo hell
kode = 3
Call max
    ms = MsgBox("Apakah anda ingin Menghapus data ini ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql = "delete from customer where kdcustomer='" & rs!kdcustomer & "' "
        con.Execute (sql)
        
        TimerALL.Interval = 10
    Else
        Exit Sub
    End If


Exit Sub
hell:
MsgBox err.Description
End Sub


Private Sub all()
MousePointer = vbHourglass


If OPT1.Value = True Then
    If ChkR.Value = 0 Then
        If TXTCARI = "" Then
        sql = "select a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP order by a.kdcustomer desc"
        Else
        sql = "select a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP where " & kategori & " like '%" & TXTCARI & "%' order by a.kdcustomer desc"
        End If
    Else
        If TXTCARI = "" Then
        sql = "select TOP " & CLng(txtR) & " a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP order by a.kdcustomer desc"
        Else
        sql = "select TOP " & CLng(txtR) & " a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP where " & kategori & " like '%" & TXTCARI & "%' order by a.kdcustomer desc"
        End If
    End If
Else
    If ChkR.Value = 0 Then
        If TXTCARI = "" Then
        sql1 = "select a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP "
        Else
        sql1 = "select a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP where " & kategori & " like '%" & TXTCARI & "%'"
        End If
        
        sql = "select * from (" & sql1 & ") x where kdcustomer in (select kdcustomer from rekap_pjm_sewa where kdkategori between '04' and '10') order by kdcustomer desc"
        
    Else
        If TXTCARI = "" Then
        sql1 = "select  a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP"
        Else
        sql1 = "select  a.*,b.nmareaC,c.nmteknisi,N_aktif = case when a.non_aktif=1 then 'X' else '' end,isnull(d.nmcustomer_IAP,'') as nmcustomer_IAP from Customer a left join area_cheker b on a.kdareaC=b.kdareaC left join teknisi c on a.kdteknisi=c.kdteknisi left join customer_IAP d on  convert(varchar,a.kdSP) + '/' + a.kdcustomer_IAP=d.PK_CUST_IAP where " & kategori & " like '%" & TXTCARI & "%'"
        End If
        
        sql = "select TOP " & CLng(txtR) & " * from (" & sql1 & ") x where kdcustomer in (select kdcustomer from rekap_pjm_sewa where kdkategori between '04' and '10') order by kdcustomer desc"
    End If
    
    
End If


Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

Call LG

For i = 1 To (datagrid1.Rows - 1)
For j = 1 To (datagrid1.Cols - 1)


If datagrid1.TextMatrix(i, 34) = "X" Then
datagrid1.Cell(flexcpForeColor, i, j) = vbRed
End If

If datagrid1.TextMatrix(i, 31) = 1 And datagrid1.TextMatrix(i, 21) = 0 Then
datagrid1.Cell(flexcpForeColor, i, j) = &HFF00FF
End If


If datagrid1.TextMatrix(i, 35) = "" Then
datagrid1.Cell(flexcpBackColor, i, 1) = vbYellow
End If


Next
Next

MousePointer = vbDefault
End Sub

Private Sub CMBCARI_Click()
If CMBCARI.ListIndex = 0 Then
kategori = "a.nmCustomer"
ElseIf CMBCARI.ListIndex = 1 Then
kategori = "a.kdCustomer"
ElseIf CMBCARI.ListIndex = 2 Then
kategori = "a.alamat"
ElseIf CMBCARI.ListIndex = 3 Then
kategori = "b.nmareaC"
ElseIf CMBCARI.ListIndex = 4 Then
kategori = "c.nmteknisi"
ElseIf CMBCARI.ListIndex = 5 Then
kategori = "a.non_aktif"

End If

TimerALL.Interval = 10
End Sub

Private Sub CMBCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
Call tbh
ElseIf Index = 1 Then
     If rs.RecordCount <> 0 Then
     Call ubh
     End If
ElseIf Index = 2 Then
     If rs.RecordCount <> 0 Then
     Call hps
     End If
ElseIf Index = 3 Then
TXTCARI = ""
Call all
ElseIf Index = 4 Then
TXTCARI = ""
    If TXTCARI.Enabled = True Then
    Me.Height = Me.Height - 1170

    TXTCARI.Enabled = False
    CMBCARI.Enabled = False
    Else
    Me.Height = Me.Height + 1170

    TXTCARI.Enabled = True
    CMBCARI.Enabled = True
    End If
ElseIf Index = 5 Then
List_Customer.Show vbModal

End If
End Sub



Private Sub cmdT_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
 TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub

Private Sub datagrid1_Click()
TimerG.Interval = 10
End Sub

Private Sub datagrid1_DblClick()


 If rs.RecordCount <> 0 Then
 Call ubh
 End If

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyLeft Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyRight Then
cmdT(0).SetFocus
ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
 Call tbh
ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
 If rs.RecordCount <> 0 Then
 Call ubh
 End If
ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
 If rs.RecordCount <> 0 Then
 Call hps
 End If
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
ElseIf KeyAscii = Asc("k") Or KeyAscii = Asc("K") Then
 CMBCARI.SetFocus
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()

GradientForm Me, 0

Me.Height = Me.Height - 1170


CMBCARI.AddItem "NAMA CUSTOMER"
CMBCARI.AddItem "KODE"
CMBCARI.AddItem "ALAMAT"
CMBCARI.AddItem "AREA CHEKER"
CMBCARI.AddItem "CEKHER"
CMBCARI.AddItem "NON AKTIF"

CMBCARI.ListIndex = 0

OPT1.Value = True


TimerALL.Interval = 10
End Sub

Private Sub OPT1_Click(Value As Integer)
TimerALL.Interval = 10
End Sub

Private Sub Opt2_Click(Value As Integer)
TimerALL.Interval = 10
End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
Call all

If kode = 2 Or kode = 3 Then
rs.AbsolutePosition = lblpos
End If

TimerALL.Interval = 0

MousePointer = vbDefault

End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerALL.Interval = 10
End Sub

Private Sub TXTCARI_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTCARI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    TimerG.Interval = 10
    Else
    SendKeys vbTab
    End If
ElseIf KeyAscii = 27 Then
Unload Me
'ElseIf KeyAscii = Asc("t") Or KeyAscii = Asc("T") Then
' Call tbh
'ElseIf KeyAscii = Asc("u") Or KeyAscii = Asc("U") Then
' If rs.RecordCount <> 0 Then
' Call ubh
' End If
'ElseIf KeyAscii = Asc("h") Or KeyAscii = Asc("H") Then
' If rs.RecordCount <> 0 Then
' Call hps
' End If
'ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
' Call all
End If
End Sub




Private Sub txtR_Change()
Call nul(txtR)
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii <> vbKeyBack Then

    cekTBL = InStr("1234567890.,", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txtR_LostFocus()
On Error GoTo hell

txtR = FormatNumber(txtR, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtR.SetFocus

End Sub


