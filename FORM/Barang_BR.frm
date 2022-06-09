VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Barang_BR 
   BorderStyle     =   0  'None
   ClientHeight    =   10305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   19305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglnon_aktif 
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
      Left            =   16650
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtcari 
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
      Left            =   450
      TabIndex        =   0
      Top             =   1485
      Width           =   2490
   End
   Begin VB.Timer TimerALL 
      Left            =   6120
      Top             =   1665
   End
   Begin VB.Timer TimerG 
      Left            =   5580
      Top             =   1665
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   225
      TabIndex        =   5
      Top             =   855
      Width           =   17835
      _Version        =   524288
      _ExtentX        =   31459
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1305
      TabIndex        =   7
      Top             =   9765
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
      Picture         =   "Barang_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   2925
      TabIndex        =   2
      ToolTipText     =   "Tambah Barang Baru"
      Top             =   1485
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
      Picture         =   "Barang_BR.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   7665
      Left            =   450
      TabIndex        =   1
      Top             =   1935
      Width           =   17655
      _cx             =   31141
      _cy             =   13520
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
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Barang_BR.frx":902B
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
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
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "PER TGL NON AKTIF :"
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
      Left            =   14940
      TabIndex        =   9
      Top             =   1485
      Width           =   1725
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1530
      Picture         =   "Barang_BR.frx":9188
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   18270
      Picture         =   "Barang_BR.frx":16038
      Stretch         =   -1  'True
      Top             =   360
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barang"
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
      Left            =   1035
      TabIndex        =   6
      Top             =   90
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data :"
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
      Left            =   495
      TabIndex        =   4
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   9765
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   10275
      Left            =   45
      Picture         =   "Barang_BR.frx":163F8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19230
   End
End
Attribute VB_Name = "Barang_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim sqlX1, sql1, sqlA, sqlA1 As String
Dim i, j As Integer

Private Sub cmdBR_Click()
Barang_TU.LBLKODE = 1
Barang_TU.lblfrm = "BARANG_BR"
Barang_TU.Show vbModal
End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
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




Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub LG()
On Error GoTo hell




Exit Sub
hell:

End Sub

Private Sub all()
'On Error GoTo hell

MousePointer = vbHourglass

'hanya utk TEKNISIDALAM_D
If LBLKODE = "TEKNISIDALAM_D" Then
sqlA1 = "select '' as status,kdgudang as kdcustomer,kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - Repair) as qty" & vbCrLf & _
        "from RKP_stok where tgl <= '" & Format(Date, "yyyy/MM/dd") & "' and kdgudang in ('GD2') group by kdbarang,kdgudang"
ElseIf LBLKODE = "PINJAM_DTU" Then
sqlA1 = "select '' as status,kdgudang as kdcustomer,kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - Repair) as qty" & vbCrLf & _
        "from RKP_stok where tgl <= '" & Format(Date, "yyyy/MM/dd") & "' and kdgudang in ('" & Pinjam_D.lblkdgudang & "') group by kdbarang,kdgudang"
ElseIf LBLKODE = "SEWA_DTU" Then
sqlA1 = "select '' as status,kdgudang as kdcustomer,kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - Repair) as qty" & vbCrLf & _
        "from RKP_stok where tgl <= '" & Format(Date, "yyyy/MM/dd") & "' and kdgudang in ('" & Sewa_d.lblkdgudang & "') group by kdbarang,kdgudang"
ElseIf LBLKODE = "FREE_DTU" Then
sqlA1 = "select '' as status,kdgudang as kdcustomer,kdbarang,sum(U_beli + U_Rpinjam + U_Rsewa + M_unit - U_free - U_pinjam - U_sewa  - K_unit - Repair) as qty" & vbCrLf & _
        "from RKP_stok where tgl <= '" & Format(Date, "yyyy/MM/dd") & "' and kdgudang in ('" & Free_D.lblkdgudang & "') group by kdbarang,kdgudang"


End If


sqlA = "select a.status,a.kdcustomer,b.nmgudang as nmcustomer,'-' as alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdsap,c.kdkategori,sum(qty) as qty " & vbCrLf & _
       "from (" & sqlA1 & ") a left join gudang b on a.kdcustomer=b.kdgudang left join barang c on a.kdbarang=c.kdbarang where c.kdkategori in ('04','05','06','07','08','09','10') group by a.status,a.kdcustomer,b.nmgudang,a.kdbarang,c.nmbarang,c.kd1,c.kdsap,c.kdkategori"

'------------


If LBLKODE = "PO_D" Then
    If txtcari = "" Then
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori>'03'  and a.non_aktif=0  "
    Else
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori>'03' and a.non_aktif=0 and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%' or a.merk like '%" & txtcari & "%')"
    End If
ElseIf LBLKODE = "TEKNISIDALAM_DTU" Then
    If txtcari = "" Then
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori='03'  and a.non_aktif=0  "
    Else
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori='03' and a.non_aktif=0 and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%' or a.merk like '%" & txtcari & "%')"
    End If
ElseIf LBLKODE = "TEKNISIDALAM_D" Then
  

    If txtcari = "" Then
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori between '04' and '10' and a.kdbarang in (select kdbarang from (" & sqlA & ") x WHERE qty<>0 ) and a.non_aktif=0 "
    Else
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori between '04' and '10' and a.kdbarang in (select kdbarang from (" & sqlA & ") x WHERE qty<>0 ) and a.non_aktif=0 and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%' or a.merk like '%" & txtcari & "%')"
    End If

ElseIf LBLKODE = "PINJAM_DTU" Or LBLKODE = "SEWA_DTU" Or LBLKODE = "FREE_DTU" Then
  
    If txtcari = "" Then
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori between '04' and '10' and a.kdbarang in (select kdbarang from (" & sqlA & ") x WHERE qty<>0 ) and a.non_aktif=0 "
    Else
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori between '04' and '10' and a.kdbarang in (select kdbarang from (" & sqlA & ") x WHERE qty<>0 ) and a.non_aktif=0 and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%' or a.merk like '%" & txtcari & "%')"
    End If

ElseIf LBLKODE = "TEKNISILUAR_D" Then
    If txtcari = "" Then
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori >'03' and kdbarang in (select kdbarang from rekap_pjm_sewa where kdcustomer='" & TeknisiLuar_D.lblkdcustomer & "') "
    Else
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.kdkategori>'03' and kdbarang in (select kdbarang from rekap_pjm_sewa where kdcustomer='" & TeknisiLuar_D.lblkdcustomer & "') and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%' or a.merk like '%" & txtcari & "%')"
    End If


Else

    If txtcari = "" Then
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.non_aktif=0  "
    Else
    sql1 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.non_aktif=0 and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%') "
    End If
    
End If

If txtcari = "" Then
    sql2 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from Barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.non_aktif=1 and a.tglnon_aktif > '" & Format(txttglnon_aktif, "yyyy/MM/dd") & "'  "
Else
    sql2 = "select a.kdbarang,a.kd1,a.kdSAP,a.nmbarang,b.nmkategori,a.merk,a.satuan,N_aktif = case when a.non_aktif=1 then 'X' else '' end,a.non_aktif,a.tglnon_aktif from barang a left join kategoribrg b on a.kdkategori=b.kdkategori where a.non_aktif=1 and a.tglnon_aktif > '" & Format(txttglnon_aktif, "yyyy/MM/dd") & "' and (a.kdbarang like '%" & txtcari & "%' or a.nmbarang like '%" & txtcari & "%' or a.kd1 like '%" & txtcari & "%' or a.kdsap like '%" & txtcari & "%' or a.merk like '%" & txtcari & "%') "
End If



sql = "select * from (" & sql1 & " union all " & sql2 & ") x order by kdbarang"

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
Call LG

For i = 1 To (datagrid1.Rows - 1)
For j = 1 To (datagrid1.Cols - 1)


If datagrid1.TextMatrix(i, 8) = 1 Then
datagrid1.Cell(flexcpForeColor, i, j) = vbRed
End If


Next
Next

MousePointer = vbDefault

'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell
If LBLKODE = UCase("PObeli_DTU") Then
PObeli_DTU.lblkdbarang = rs!kdbarang
PObeli_DTU.lblnmbarang = rs!nmbarang
PObeli_DTU.lblsatuan = rs!satuan

ElseIf LBLKODE = UCase("po_DTU") Then
PO_DTU.lblkdbarang = rs!kdbarang
PO_DTU.lblnmbarang = rs!nmbarang
PO_DTU.lblsatuan = rs!satuan

ElseIf LBLKODE = UCase("HRGSEWA") Then
hrgSewa_TU.lblkdbarang = rs!kdbarang
hrgSewa_TU.lblnmbarang = rs!nmbarang

ElseIf LBLKODE = UCase("3A4") Then
Cetak_3A4.lblkdbarang = rs!kdbarang
Cetak_3A4.lblnmbarang = rs!nmbarang

ElseIf LBLKODE = UCase("4A4") Then
Cetak_4A4.lblkdbarang = rs!kdbarang
Cetak_4A4.lblnmbarang = rs!nmbarang


ElseIf LBLKODE = UCase("1A2") Then
Cetak_1A2.lblkdbarang = rs!kdbarang
Cetak_1A2.lblnmbarang = rs!nmbarang

ElseIf LBLKODE = UCase("6A3") Then
Cetak_6A3.lblkdbarang = rs!kdbarang
Cetak_6A3.lblnmbarang = rs!nmbarang

ElseIf LBLKODE = UCase("PO_D") Then
PO_D.lblkdbarang = rs!kdbarang
PO_D.lblnmbarang = rs!nmbarang


ElseIf LBLKODE = "6A4_01" Then
Cetak_6A4.lblkdbarang1 = rs!kdbarang
ElseIf LBLKODE = "6A4_02" Then
Cetak_6A4.lblkdbarang2 = rs!kdbarang
ElseIf LBLKODE = "6A4_03" Then
Cetak_6A4.lblkdbarang3 = rs!kdbarang
ElseIf LBLKODE = "6A4_04" Then
Cetak_6A4.lblkdbarang4 = rs!kdbarang
ElseIf LBLKODE = "6A4_05" Then
Cetak_6A4.lblkdbarang5 = rs!kdbarang
ElseIf LBLKODE = "6A4_06" Then
Cetak_6A4.lblkdbarang6 = rs!kdbarang
ElseIf LBLKODE = "6A4_07" Then
Cetak_6A4.lblkdbarang7 = rs!kdbarang
ElseIf LBLKODE = "6A4_08" Then
Cetak_6A4.lblkdbarang8 = rs!kdbarang
ElseIf LBLKODE = "6A4_09" Then
Cetak_6A4.lblkdbarang9 = rs!kdbarang
ElseIf LBLKODE = "6A4_10" Then
Cetak_6A4.lblkdbarang10 = rs!kdbarang
ElseIf LBLKODE = "6A4_11" Then
Cetak_6A4.lblkdbarang11 = rs!kdbarang
ElseIf LBLKODE = "6A4_12" Then
Cetak_6A4.lblkdbarang12 = rs!kdbarang
ElseIf LBLKODE = "6A4_13" Then
Cetak_6A4.lblkdbarang13 = rs!kdbarang
ElseIf LBLKODE = "6A4_14" Then
Cetak_6A4.lblkdbarang14 = rs!kdbarang
ElseIf LBLKODE = "6A4_15" Then
Cetak_6A4.lblkdbarang15 = rs!kdbarang
ElseIf LBLKODE = "6A4_16" Then
Cetak_6A4.lblkdbarang16 = rs!kdbarang
ElseIf LBLKODE = "6A4_17" Then
Cetak_6A4.lblkdbarang17 = rs!kdbarang
ElseIf LBLKODE = "6A4_18" Then
Cetak_6A4.lblkdbarang18 = rs!kdbarang
ElseIf LBLKODE = "6A4_19" Then
Cetak_6A4.lblkdbarang19 = rs!kdbarang
ElseIf LBLKODE = "6A4_20" Then
Cetak_6A4.lblkdbarang20 = rs!kdbarang
ElseIf LBLKODE = "TEKNISIDALAM_D" Then
TeknisiDalam_D.lblkdbarang = rs!kdbarang
TeknisiDalam_D.lblkd1 = rs!kd1
TeknisiDalam_D.lblkdsap = rs!kdSAP
TeknisiDalam_D.lblnmkategori = rs!nmkategori
TeknisiDalam_D.lblmerk = rs!merk
ElseIf LBLKODE = UCase("TEKNISIDALAM_DTU") Then
TeknisiDalam_DTU.lblkdbarang = rs!kdbarang
TeknisiDalam_DTU.lblnmbarang = rs!nmbarang
ElseIf LBLKODE = "TEKNISILUAR_D" Then
TeknisiLuar_D.lblkdbarang = rs!kdbarang
TeknisiLuar_D.lblkd1 = rs!kd1
TeknisiLuar_D.lblkdsap = rs!kdSAP
TeknisiLuar_D.lblnmkategori = rs!nmkategori
TeknisiLuar_D.lblmerk = rs!merk
ElseIf LBLKODE = UCase("PINJAM_DTU") Then
Pinjam_DTU.lblkdbarang = rs!kdbarang
Pinjam_DTU.lblnmbarang = rs!nmbarang
ElseIf LBLKODE = UCase("SEWA_DTU") Then
Sewa_DTU.lblkdbarang = rs!kdbarang
Sewa_DTU.lblnmbarang = rs!nmbarang


End If
Unload Me

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyUp Then

    If rs.AbsolutePosition = 1 Then
    txtcari.SetFocus
    End If

ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

On Error GoTo hell

If KeyAscii = 13 Then
    
    If LBLKODE = UCase("PObeli_DTU") Then
    PObeli_DTU.lblkdbarang = rs!kdbarang
    PObeli_DTU.lblnmbarang = rs!nmbarang
    PObeli_DTU.lblsatuan = rs!satuan
    ElseIf LBLKODE = UCase("po_DTU") Then
    PO_DTU.lblkdbarang = rs!kdbarang
    PO_DTU.lblnmbarang = rs!nmbarang
    PO_DTU.lblsatuan = rs!satuan
    ElseIf LBLKODE = UCase("HRGSEWA") Then
    hrgSewa_TU.lblkdbarang = rs!kdbarang
    hrgSewa_TU.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = UCase("3A4") Then
    Cetak_3A4.lblkdbarang = rs!kdbarang
    Cetak_3A4.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = UCase("4A4") Then
    Cetak_4A4.lblkdbarang = rs!kdbarang
    Cetak_4A4.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = UCase("1A2") Then
    Cetak_1A2.lblkdbarang = rs!kdbarang
    Cetak_1A2.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = UCase("6A3") Then
    Cetak_6A3.lblkdbarang = rs!kdbarang
    Cetak_6A3.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = UCase("PO_D") Then
    PO_D.lblkdbarang = rs!kdbarang
    PO_D.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = "6A4_01" Then
    Cetak_6A4.lblkdbarang1 = rs!kdbarang
    ElseIf LBLKODE = "6A4_02" Then
    Cetak_6A4.lblkdbarang2 = rs!kdbarang
    ElseIf LBLKODE = "6A4_03" Then
    Cetak_6A4.lblkdbarang3 = rs!kdbarang
    ElseIf LBLKODE = "6A4_04" Then
    Cetak_6A4.lblkdbarang4 = rs!kdbarang
    ElseIf LBLKODE = "6A4_05" Then
    Cetak_6A4.lblkdbarang5 = rs!kdbarang
    ElseIf LBLKODE = "6A4_06" Then
    Cetak_6A4.lblkdbarang6 = rs!kdbarang
    ElseIf LBLKODE = "6A4_07" Then
    Cetak_6A4.lblkdbarang7 = rs!kdbarang
    ElseIf LBLKODE = "6A4_08" Then
    Cetak_6A4.lblkdbarang8 = rs!kdbarang
    ElseIf LBLKODE = "6A4_09" Then
    Cetak_6A4.lblkdbarang9 = rs!kdbarang
    ElseIf LBLKODE = "6A4_10" Then
    Cetak_6A4.lblkdbarang10 = rs!kdbarang
    ElseIf LBLKODE = "6A4_11" Then
    Cetak_6A4.lblkdbarang11 = rs!kdbarang
    ElseIf LBLKODE = "6A4_12" Then
    Cetak_6A4.lblkdbarang12 = rs!kdbarang
    ElseIf LBLKODE = "6A4_13" Then
    Cetak_6A4.lblkdbarang13 = rs!kdbarang
    ElseIf LBLKODE = "6A4_14" Then
    Cetak_6A4.lblkdbarang14 = rs!kdbarang
    ElseIf LBLKODE = "6A4_15" Then
    Cetak_6A4.lblkdbarang15 = rs!kdbarang
    ElseIf LBLKODE = "6A4_16" Then
    Cetak_6A4.lblkdbarang16 = rs!kdbarang
    ElseIf LBLKODE = "6A4_17" Then
    Cetak_6A4.lblkdbarang17 = rs!kdbarang
    ElseIf LBLKODE = "6A4_18" Then
    Cetak_6A4.lblkdbarang18 = rs!kdbarang
    ElseIf LBLKODE = "6A4_19" Then
    Cetak_6A4.lblkdbarang19 = rs!kdbarang
    ElseIf LBLKODE = "6A4_20" Then
    Cetak_6A4.lblkdbarang20 = rs!kdbarang
    ElseIf LBLKODE = "TEKNISIDALAM_D" Then
    TeknisiDalam_D.lblkdbarang = rs!kdbarang
    TeknisiDalam_D.lblkd1 = rs!kd1
    TeknisiDalam_D.lblkdsap = rs!kdSAP
    TeknisiDalam_D.lblnmkategori = rs!nmkategori
    TeknisiDalam_D.lblmerk = rs!merk
    ElseIf LBLKODE = UCase("TEKNISIDALAM_DTU") Then
    TeknisiDalam_DTU.lblkdbarang = rs!kdbarang
    TeknisiDalam_DTU.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = "TEKNISILUAR_D" Then
    TeknisiLuar_D.lblkdbarang = rs!kdbarang
    TeknisiLuar_D.lblkd1 = rs!kd1
    TeknisiLuar_D.lblkdsap = rs!kdSAP
    TeknisiLuar_D.lblnmkategori = rs!nmkategori
    TeknisiLuar_D.lblmerk = rs!merk
    ElseIf LBLKODE = UCase("PINJAM_DTU") Then
    Pinjam_DTU.lblkdbarang = rs!kdbarang
    Pinjam_DTU.lblnmbarang = rs!nmbarang
    ElseIf LBLKODE = UCase("SEWA_DTU") Then
    Sewa_DTU.lblkdbarang = rs!kdbarang
    Sewa_DTU.lblnmbarang = rs!nmbarang

    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0

txttglnon_aktif = Date

TimerALL.Interval = 10
End Sub





Private Sub TimerAll_Timer()
On Error Resume Next

Call all

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
    SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub




Private Sub txttglNon_aktif_Change()
Call nul(txttglnon_aktif)
End Sub

Private Sub txttglNon_aktif_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglNon_aktif_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglNon_aktif_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
TimerALL.Interval = 10
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttglNon_aktif_LostFocus()
On Error GoTo hell

txttglnon_aktif = FormatDateTime(txttglnon_aktif, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglnon_aktif.SetFocus

End Sub


