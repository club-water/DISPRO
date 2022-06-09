VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Customer_TU 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglsegel 
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
      Left            =   12150
      TabIndex        =   35
      Top             =   5355
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox chkpph23 
      BackColor       =   &H00000000&
      Caption         =   "PPH 23"
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
      Height          =   285
      Left            =   7155
      TabIndex        =   34
      Top             =   4545
      Width           =   915
   End
   Begin VB.CheckBox Chkket 
      BackColor       =   &H00000000&
      Caption         =   "DOWNLINE"
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
      Left            =   2880
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   4680
      Width           =   1320
   End
   Begin VB.TextBox txtalamat_TGH 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   2295
      Width           =   5505
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D3 
      Height          =   30
      Left            =   7110
      TabIndex        =   81
      Top             =   3330
      Width           =   6945
      _Version        =   524288
      _ExtentX        =   12250
      _ExtentY        =   53
      _StockProps     =   8
   End
   Begin VB.TextBox txtalamatNPWP 
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
      Height          =   915
      Left            =   8955
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   2250
      Width           =   5100
   End
   Begin VB.TextBox txtnmNPWP 
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
      Left            =   8955
      TabIndex        =   26
      Top             =   1890
      Width           =   5100
   End
   Begin VB.TextBox txtnoNPWP 
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
      Left            =   8955
      TabIndex        =   25
      Top             =   1530
      Width           =   5100
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   7035
      Left            =   7065
      TabIndex        =   77
      Top             =   720
      Width           =   15
      _Version        =   524288
      _ExtentX        =   26
      _ExtentY        =   12409
      _StockProps     =   8
   End
   Begin VB.CheckBox ChkNA 
      BackColor       =   &H00000000&
      Caption         =   "NON AKTIF"
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
      Height          =   330
      Left            =   12825
      TabIndex        =   36
      Top             =   6255
      Width           =   1230
   End
   Begin VB.TextBox txtgln 
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
      Left            =   15840
      TabIndex        =   33
      Text            =   "0"
      Top             =   3015
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtbtl 
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
      Left            =   15840
      TabIndex        =   32
      Text            =   "0"
      Top             =   2655
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtcup 
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
      Left            =   15840
      TabIndex        =   31
      Text            =   "0"
      Top             =   2295
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txttglspk2 
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
      Left            =   10305
      TabIndex        =   30
      Text            =   "01/01/1900"
      Top             =   3870
      Width           =   1590
   End
   Begin VB.TextBox txttglspk1 
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
      Left            =   8370
      TabIndex        =   29
      Text            =   "01/01/1900"
      Top             =   3870
      Width           =   1590
   End
   Begin VB.TextBox txtnospk 
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
      Left            =   8370
      TabIndex        =   28
      Top             =   3510
      Width           =   5640
   End
   Begin VB.ComboBox CMbJNSBYR 
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
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   6345
      Width           =   1500
   End
   Begin VB.Timer TimerCMB 
      Left            =   3780
      Top             =   810
   End
   Begin VB.ComboBox CMBbank 
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
      Left            =   5310
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6345
      Width           =   1095
   End
   Begin VB.TextBox txtkdcustomer_IAP 
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
      Left            =   1440
      TabIndex        =   11
      Top             =   4680
      Width           =   1275
   End
   Begin VB.TextBox txtCP 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   3015
      Width           =   1950
   End
   Begin VB.TextBox txthrgSewa 
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
      Left            =   4635
      TabIndex        =   6
      Text            =   "37000"
      Top             =   3015
      Width           =   1410
   End
   Begin VB.Timer TimerNO 
      Left            =   2745
      Top             =   765
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
      Left            =   1440
      TabIndex        =   17
      Top             =   10980
      Visible         =   0   'False
      Width           =   5505
   End
   Begin VB.TextBox lblkdcustomer 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1215
      Width           =   1680
   End
   Begin VB.TextBox txtalamat 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1935
      Width           =   5505
   End
   Begin VB.TextBox TXTnmcustomer 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1575
      Width           =   5505
   End
   Begin VB.TextBox txttelp 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2655
      Width           =   5505
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   780
      Left            =   14175
      TabIndex        =   37
      ToolTipText     =   "Simpan"
      Top             =   6345
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
      Picture         =   "Customer_TU.frx":0000
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   540
      TabIndex        =   41
      Top             =   720
      Width           =   13560
      _Version        =   524288
      _ExtentX        =   23918
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   11070
      TabIndex        =   39
      Top             =   7380
      Width           =   3030
      _ExtentX        =   5345
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
      Picture         =   "Customer_TU.frx":2A6D
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdhrg 
      Height          =   915
      Left            =   9315
      TabIndex        =   40
      ToolTipText     =   "Tambah"
      Top             =   6570
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1614
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
      Picture         =   "Customer_TU.frx":92CF
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR 
      Height          =   420
      Left            =   13455
      TabIndex        =   9
      ToolTipText     =   "Simpan"
      Top             =   10980
      Visible         =   0   'False
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
      Picture         =   "Customer_TU.frx":BF43
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   6435
      TabIndex        =   10
      Top             =   3915
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
      Picture         =   "Customer_TU.frx":E775
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR2 
      Height          =   420
      Left            =   6390
      TabIndex        =   14
      Top             =   5895
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
      Picture         =   "Customer_TU.frx":10FA7
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR3 
      Height          =   420
      Left            =   7065
      TabIndex        =   18
      ToolTipText     =   "Simpan"
      Top             =   10980
      Visible         =   0   'False
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
      Picture         =   "Customer_TU.frx":137D9
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSOption Opt1 
      Height          =   330
      Left            =   12105
      TabIndex        =   23
      Top             =   1170
      Width           =   690
      _ExtentX        =   1217
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
      Caption         =   "PKP"
   End
   Begin Threed.SSOption Opt2 
      Height          =   330
      Left            =   12915
      TabIndex        =   24
      Top             =   1170
      Width           =   1140
      _ExtentX        =   2011
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
      Caption         =   "NON PKP"
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D4 
      Height          =   30
      Left            =   1035
      TabIndex        =   85
      Top             =   6795
      Width           =   6045
      _Version        =   524288
      _ExtentX        =   10663
      _ExtentY        =   53
      _StockProps     =   8
   End
   Begin Threed.SSCommand cmdBR4 
      Height          =   420
      Left            =   5985
      TabIndex        =   19
      Top             =   6885
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
      Picture         =   "Customer_TU.frx":1600B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR5 
      Height          =   420
      Left            =   5985
      TabIndex        =   21
      Top             =   7245
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
      Picture         =   "Customer_TU.frx":1883D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC4 
      Height          =   420
      Left            =   6480
      TabIndex        =   20
      Top             =   6885
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
      Picture         =   "Customer_TU.frx":1B06F
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC5 
      Height          =   420
      Left            =   6480
      TabIndex        =   22
      Top             =   7245
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
      Picture         =   "Customer_TU.frx":1D6B9
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR6 
      Height          =   420
      Left            =   5940
      TabIndex        =   7
      Top             =   3330
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
      Picture         =   "Customer_TU.frx":1FD03
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC6 
      Height          =   420
      Left            =   6435
      TabIndex        =   8
      Top             =   3330
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
      Picture         =   "Customer_TU.frx":22535
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   2670
      Left            =   855
      TabIndex        =   97
      Top             =   7920
      Width           =   13065
      _cx             =   23045
      _cy             =   4710
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
      BackColorAlternate=   12632256
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
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
      FormatString    =   $"Customer_TU.frx":24B7F
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
      Editable        =   2
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
   Begin PVLINE3DLib.PVLine3D PVLine3D5 
      Height          =   30
      Left            =   1620
      TabIndex        =   98
      Top             =   3870
      Width           =   5415
      _Version        =   524288
      _ExtentX        =   9551
      _ExtentY        =   53
      _StockProps     =   8
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D6 
      Height          =   30
      Left            =   1440
      TabIndex        =   100
      Top             =   5850
      Width           =   5640
      _Version        =   524288
      _ExtentX        =   9948
      _ExtentY        =   53
      _StockProps     =   8
   End
   Begin Threed.SSCommand cmdIAP 
      Height          =   330
      Left            =   4275
      TabIndex        =   13
      ToolTipText     =   "Pilih Dari List Customer IAP"
      Top             =   4680
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
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
      Picture         =   "Customer_TU.frx":24CDB
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D7 
      Height          =   30
      Left            =   7110
      TabIndex        =   106
      Top             =   4410
      Width           =   6945
      _Version        =   524288
      _ExtentX        =   12250
      _ExtentY        =   53
      _StockProps     =   8
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   0
      Left            =   14175
      TabIndex        =   38
      ToolTipText     =   "Cek Omset"
      Top             =   7155
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
      Picture         =   "Customer_TU.frx":25075
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdT 
      Height          =   780
      Index           =   1
      Left            =   14175
      TabIndex        =   108
      ToolTipText     =   "Cetak"
      Top             =   7965
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
      Picture         =   "Customer_TU.frx":296EC
      ButtonStyle     =   4
   End
   Begin VB.Label lblkdcustomer_IAP 
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
      Left            =   4680
      TabIndex        =   107
      Top             =   4725
      Width           =   1725
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT IAP :"
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
      Left            =   180
      TabIndex        =   105
      Top             =   5445
      Width           =   1275
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA IAP :"
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
      Left            =   270
      TabIndex        =   104
      Top             =   5130
      Width           =   1005
   End
   Begin VB.Label lblnmCustomer_IAP 
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
      Left            =   1440
      TabIndex        =   103
      Top             =   5040
      Width           =   5550
   End
   Begin VB.Label lblalamat_IAP 
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
      Left            =   1440
      TabIndex        =   102
      Top             =   5400
      Width           =   5550
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "PENAGIHAN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   270
      TabIndex        =   101
      Top             =   5715
      Width           =   1635
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER IAP :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   225
      TabIndex        =   99
      Top             =   3735
      Width           =   1635
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "HRG SEWA :"
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
      Left            =   3600
      TabIndex        =   96
      Top             =   3060
      Width           =   1230
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "PIC MARKETING :"
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
      Left            =   90
      TabIndex        =   95
      Top             =   3420
      Width           =   1365
   End
   Begin VB.Label lblkdPIC 
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
      Left            =   1440
      TabIndex        =   94
      Top             =   3375
      Width           =   870
   End
   Begin VB.Label lblnmPIC 
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
      Left            =   2340
      TabIndex        =   93
      Top             =   3375
      Width           =   3615
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "CHEKER :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   270
      TabIndex        =   92
      Top             =   6660
      Width           =   1635
   End
   Begin VB.Label lblnmareaC 
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
      Left            =   2520
      TabIndex        =   91
      Top             =   6930
      Width           =   3480
   End
   Begin VB.Label lblkdareaC 
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
      Left            =   1485
      TabIndex        =   90
      Top             =   6930
      Width           =   1005
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "AREA CHEKER :"
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
      Left            =   270
      TabIndex        =   89
      Top             =   6975
      Width           =   1185
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
      Left            =   2385
      TabIndex        =   88
      Top             =   7290
      Width           =   3615
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
      Left            =   1485
      TabIndex        =   87
      Top             =   7290
      Width           =   870
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   " CHEKER :"
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
      Left            =   315
      TabIndex        =   86
      Top             =   7335
      Width           =   1095
   End
   Begin VB.Label lbltgldibuat 
      BackStyle       =   0  'Transparent
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
      Left            =   8055
      TabIndex        =   84
      Top             =   6300
      Width           =   1635
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "DIBUAT :"
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
      Left            =   7290
      TabIndex        =   83
      Top             =   6300
      Width           =   870
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT TAGIH :"
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
      TabIndex        =   82
      Top             =   2340
      Width           =   1320
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT NPWP / KTP :"
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
      Left            =   7155
      TabIndex        =   80
      Top             =   2295
      Width           =   1770
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA NPWP / KTP :"
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
      Left            =   7155
      TabIndex        =   79
      Top             =   1935
      Width           =   1770
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "NPWP / KTP :"
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
      Left            =   7155
      TabIndex        =   78
      Top             =   1575
      Width           =   1185
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "GALON :"
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
      Left            =   15210
      TabIndex        =   76
      Top             =   3060
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "BOTOL :"
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
      Left            =   15255
      TabIndex        =   75
      Top             =   2700
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "CUP :"
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
      Left            =   15435
      TabIndex        =   74
      Top             =   2340
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "TARGET OMSET :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   15390
      TabIndex        =   73
      Top             =   1980
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "SD"
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
      Left            =   10035
      TabIndex        =   72
      Top             =   3915
      Width           =   240
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL SPK :"
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
      Left            =   7200
      TabIndex        =   71
      Top             =   3915
      Width           =   870
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "NO SPK :"
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
      Left            =   7200
      TabIndex        =   70
      Top             =   3555
      Width           =   1140
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS PEMBAYARAN :"
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
      Left            =   270
      TabIndex        =   69
      Top             =   6390
      Width           =   1680
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "AREA TAGIH :"
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
      Left            =   945
      TabIndex        =   68
      Top             =   11070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LBLKDAREA 
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
      Left            =   2115
      TabIndex        =   67
      Top             =   11025
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label LBLNMAREA 
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
      Left            =   3015
      TabIndex        =   66
      Top             =   11070
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "KOLEKTOR :"
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
      Left            =   315
      TabIndex        =   65
      Top             =   5985
      Width           =   960
   End
   Begin VB.Label lblkdkolektor 
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
      Left            =   1440
      TabIndex        =   64
      Top             =   5940
      Width           =   870
   End
   Begin VB.Label lblnmkolektor 
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
      Left            =   2340
      TabIndex        =   63
      Top             =   5940
      Width           =   4020
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFER KE BANK :"
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
      Left            =   3645
      TabIndex        =   62
      Top             =   6390
      Width           =   1680
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "KD CUST IAP :"
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
      TabIndex        =   61
      Top             =   4725
      Width           =   1230
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "CABANG IAP :"
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
      Left            =   270
      TabIndex        =   60
      Top             =   4365
      Width           =   1230
   End
   Begin VB.Label lblnmcabang 
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
      Left            =   1440
      TabIndex        =   59
      Top             =   4320
      Width           =   4965
   End
   Begin VB.Label lblkdSP 
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
      Left            =   7470
      TabIndex        =   58
      Top             =   6885
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SP IAP :"
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
      Left            =   270
      TabIndex        =   57
      Top             =   4005
      Width           =   825
   End
   Begin VB.Label lblnosp 
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
      Left            =   1440
      TabIndex        =   56
      Top             =   3960
      Width           =   870
   End
   Begin VB.Label lblnmsp 
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
      Left            =   2340
      TabIndex        =   55
      Top             =   3960
      Width           =   4065
   End
   Begin VB.Label lblfrm 
      Height          =   330
      Left            =   12105
      TabIndex        =   54
      Top             =   6840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblnmwilayah 
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
      Left            =   9360
      TabIndex        =   53
      Top             =   11025
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.Label lblkdwilayah 
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
      Left            =   8460
      TabIndex        =   52
      Top             =   11025
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "WILAYAH :"
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
      Left            =   7650
      TabIndex        =   51
      Top             =   11070
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CP :"
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
      TabIndex        =   50
      Top             =   3060
      Width           =   1230
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "HRG SEWA :"
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
      Left            =   3510
      TabIndex        =   49
      Top             =   2700
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   3150
      Picture         =   "Customer_TU.frx":2D149
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Customer"
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
      Left            =   900
      TabIndex        =   48
      Top             =   45
      Width           =   4605
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NO TELP :"
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
      TabIndex        =   47
      Top             =   2700
      Width           =   1230
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
      Left            =   180
      TabIndex        =   46
      Top             =   11070
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT  :"
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
      TabIndex        =   45
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   10305
      TabIndex        =   44
      Top             =   6885
      Visible         =   0   'False
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
      TabIndex        =   43
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER :"
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
      TabIndex        =   42
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   435
      Left            =   14175
      Picture         =   "Customer_TU.frx":2E406
      Stretch         =   -1  'True
      Top             =   495
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   10905
      Left            =   0
      Picture         =   "Customer_TU.frx":2E7C6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "Customer_TU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim a, i As Integer
Dim rsA As ADODB.Recordset
Dim rsK As ADODB.Recordset
Dim index_PKP As Integer
Dim rsTGLSERVER As ADODB.Recordset
Dim rsAreaC As ADODB.Recordset
Dim rswil As ADODB.Recordset
Dim rsteknisi As ADODB.Recordset
Dim rsPIC As ADODB.Recordset
Dim rsIAP As ADODB.Recordset
Dim ket_downline As String
Dim rsX As ADODB.Recordset
Dim color As Long, flag As Byte
Dim bln As String
Dim ms As String

Private Sub cetak_S_Penarikan()

'Unload AR_S_PENARIKAN
'
'sqlX = "exec sp_pjm_swa_Mcust @kdcustomer='" & lblkdcustomer & "'"
'Set rsX = con.Execute(sqlX)
'
'With AR_S_PENARIKAN.DC1
'.ConnectionString = koneksi
'.Source = sqlX
'End With
'
'With AR_S_PENARIKAN
'.fldunit.DataField = "unit"
'.fldkdbarang.DataField = "kdbarang"
'.fldnmkategori.DataField = "nmkategori"
'.fldmerk.DataField = "merk"
'.fldketerangan.DataField = "keterangan"
'
'.lblnmcustomer = lblkdcustomer & " - " & TXTnmcustomer
'.lblalamat = txtalamat
'.lblCP = txtCP & " ( " & txttelp & " )"
'
'If CInt(Month(Date)) = 1 Then
'bln = "Januari"
'ElseIf CInt(Month(Date)) = 2 Then
'bln = "Februari"
'ElseIf CInt(Month(Date)) = 3 Then
'bln = "Maret"
'ElseIf CInt(Month(Date)) = 4 Then
'bln = "April"
'ElseIf CInt(Month(Date)) = 5 Then
'bln = "Mei"
'ElseIf CInt(Month(Date)) = 6 Then
'bln = "Juni"
'ElseIf CInt(Month(Date)) = 7 Then
'bln = "Juli"
'ElseIf CInt(Month(Date)) = 8 Then
'bln = "Agustus"
'ElseIf CInt(Month(Date)) = 9 Then
'bln = "September"
'ElseIf CInt(Month(Date)) = 10 Then
'bln = "Oktober"
'ElseIf CInt(Month(Date)) = 11 Then
'bln = "November"
'ElseIf CInt(Month(Date)) = 12 Then
'bln = "Desember"
'End If
'
'.lbltgl = Format(Date, "dd") & " " & bln & " " & Format(Date, "yyyy")
'
'
'.lblNB = "NB : " & ms
'
'
'
'AR_S_PENARIKAN.Show vbModal
'
'End With

End Sub

Private Sub data_IAP()
sqlIAP = "select * from customer_IAP where kdsp='" & lblkdSP & "' and kdcustomer_IAP='" & txtkdcustomer_IAP & "'"
Set rsIAP = con.Execute(sqlIAP)

If rsIAP.RecordCount <> 0 Then
lblnmcustomer_IAP = rsIAP!nmcustomer_IAP
lblalamat_IAP = rsIAP!alamat_IAP

lblnmcustomer_IAP.BackColor = vbWhite
lblalamat_IAP.BackColor = vbWhite
Else
lblnmcustomer_IAP = ""
lblalamat_IAP = ""
lblnmcustomer_IAP.BackColor = vbRed
lblalamat_IAP.BackColor = vbRed

End If




End Sub

Private Sub ChkNA_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub chksegel_Click()
If chksegel.Value = 0 Then
    txttglsegel.Enabled = False
    
Else
    txttglsegel.Enabled = True
    txttglsegel = Date
    
End If
End Sub

Private Sub chksegel_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 27 Then
Unload Me
End If

End Sub

Private Sub CMBbank_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub


Private Sub CMbJNSBYR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub cmdBR_Click()
Wilayah_BR.LBLKODE = "CUSTOMER_TU"
Wilayah_BR.Show vbModal
End Sub

Private Sub cmdBR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR1_Click()
SPIAP_BR.LBLKODE = "CUSTOMER_TU"
SPIAP_BR.Show vbModal
End Sub

Private Sub cmdBR1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub cmdBR2_Click()
Kolektor_BR.LBLKODE = "CUSTOMER_TU"
Kolektor_BR.Show vbModal
End Sub

Private Sub cmdBR2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR3_Click()
ATagih_BR.LBLKODE = "CUSTOMER_TU"
ATagih_BR.Show vbModal
End Sub

Private Sub cmdBR3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub



Private Sub cmdBR4_Click()
ACekher_BR.LBLKODE = "CUSTOMER_TU"
ACekher_BR.Show vbModal

End Sub

Private Sub cmdBR4_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me

End Sub

Private Sub cmdBR5_Click()

Teknisi_BR.LBLKODE = "CUSTOMER_TU"
Teknisi_BR.Show vbModal

End Sub

Private Sub cmdBR5_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR6_Click()
PIC_BR.LBLKODE = "CUSTOMER_TU"
PIC_BR.Show vbModal

End Sub

Private Sub cmdBR6_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC4_Click()
lblkdareaC = ""
End Sub

Private Sub cmdC5_Click()
lblkdteknisi = ""
End Sub

Private Sub cmdC6_Click()
lblkdPIC = ""
End Sub


Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If Index = 0 Then
LIST_Omset_IAP.lblkdcustomer_IAP = lblkdcustomer_IAP
LIST_Omset_IAP.lblnmcustomer_IAP = lblnmcustomer_IAP
LIST_Omset_IAP.lblalamat_IAP = lblalamat_IAP
LIST_Omset_IAP.lblnmsp = lblnmsp
LIST_Omset_IAP.lblkdcustomer = lblkdcustomer
LIST_Omset_IAP.Show vbModal

Else

Form_S_PENARIKAN.Show vbModal

'ms = InputBox("Input Note !", "NB : ", "Mohon Konfirmasi Terlebih Dahulu Sebelum Melakukan Penarikan")
'Call cetak_S_Penarikan
'
End If
End Sub

Private Sub cmdhrg_Click()
hrgSewa.Show vbModal
End Sub

Private Sub cmdhrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdIAP_Click()
Customer_IAP_BR.LBLKODE = "CUSTOMER_TU"
Customer_IAP_BR.TXTCARI = TXTnmcustomer
Customer_IAP_BR.Show vbModal

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
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



Private Sub nomer()
On Error GoTo hell

sql = "Select isnull(max(right(kdcustomer,5)),0) as xx from customer"
Set rs = con.Execute(sql)


        a = CInt(rs!xx) + 1
                
        Select Case Len(CStr(a))
        Case 1
            lblkdcustomer = "C0000" & (a)
        Case 2
            lblkdcustomer = "C000" & (a)
        Case 3
            lblkdcustomer = "C00" & (a)
        Case 4
            lblkdcustomer = "C0" & (a)
        Case 5
            lblkdcustomer = "C" & (a)
        
        End Select

Exit Sub
hell:
lblkdcustomer = "C00001"

End Sub






Private Sub cmdsimpan_Click()
On Error GoTo hell

    If TXTnmcustomer = "" Or lblkdcustomer = "" Or txtalamat = "" Or lblnosp = "" Or txtkdcustomer_IAP = "" Or txtnmNPWP = "" Or txtnoNPWP = "" Or txtalamat_TGH = "" Or txtalamatNPWP = "" Then
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !"
    Exit Sub
    Else
                
        If Chkket.Value = 1 Then
        ket_downline = "DOWNLINE"
        Else
        ket_downline = ""
        End If
        
         

         If LBLKODE = 1 Then
         Call nomer
         
         Set rsTGLSERVER = con.Execute("select getdate() as tglserver ")
             sql = "insert into Customer  values ('" & UCase(lblkdcustomer) & "','" & UCase(TXTnmcustomer) & "','" & UCase(txtalamat) & "','" & UCase(txttelp) & "','" & UCase(ket_downline) & "'," & CCur(txthrgSewa) & ",'" & UCase(txtCP) & "','" & lblkdwilayah & "','" & UCase(txtkdcustomer_IAP) & "','" & lblkdSP & "','" & CMBbank.Text & "','" & lblkdkolektor & "','" & lblkdarea & "','" & CMbJNSBYR.Text & "','" & UCase(txtnospk) & "','" & Format(txttglspk1, "yyyy/MM/dd") & "','" & Format(txttglspk2, "yyyy/MM/dd") & "'," & CCur(txtcup) & "," & CCur(txtbtl) & "," & CCur(txtgln) & "," & ChkNA.Value & ",'" & UCase(txtalamat_TGH) & "'," & index_PKP & ",'" & UCase(txtnoNPWP) & "','" & UCase(txtnmNPWP) & "','" & UCase(txtalamatNPWP) & "','" & Format(rsTGLSERVER!tglserver, "yyyy/MM/dd") & "','" & lblkdareaC & "','" & lblkdteknisi & "','" & lblkdPIC & "'," & chkpph23.Value & ")"
             con.Execute (sql)
             
             lbltgldibuat = Format(rsTGLSERVER!tglserver, "dd/MM/yyyy")
             SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
             MsgBox "Data Telah Tersimpan", vbInformation, "Info !"
             
             
         Else
             sql = "update Customer set nmcustomer='" & UCase(TXTnmcustomer) & "',alamat='" & UCase(txtalamat) & "',keterangan='" & UCase(ket_downline) & "',telp='" & UCase(txttelp) & "',hrgsewa=" & CCur(txthrgSewa) & ",kdwilayah='" & UCase(lblkdwilayah) & "',cp='" & UCase(txtCP) & "',kdSP='" & lblkdSP & "',kdcustomer_IAP='" & UCase(txtkdcustomer_IAP) & "',kdbank='" & CMBbank.Text & "',kdkolektor='" & lblkdkolektor & "',kdarea='" & lblkdarea & "',jnsbayar='" & CMbJNSBYR.Text & "',noSPK='" & UCase(txtnospk) & "',tglSPK1='" & Format(txttglspk1, "yyyy/MM/dd") & "',tglSPK2='" & Format(txttglspk2, "yyyy/MM/dd") & "',target_cup=" & CCur(txtcup) & ",target_btl=" & CCur(txtbtl) & ",target_gln=" & CCur(txtgln) & ",non_aktif=" & ChkNA.Value & ",alamat_tgh='" & UCase(txtalamat_TGH) & "',pkp=" & index_PKP & ",npwp='" & UCase(txtnoNPWP) & "',nmNPWP='" & UCase(txtnmNPWP) & "',alamatnpwp='" & UCase(txtalamatNPWP) & "',kdareaC='" & lblkdareaC & "'" & vbCrLf & _
                   ",kdteknisi='" & lblkdteknisi & "',kdpic='" & lblkdPIC & "',pph23=" & chkpph23.Value & " where kdcustomer='" & lblkdcustomer & "'"
             con.Execute (sql)
             
             SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
             MsgBox "Data Telah di Ubah", vbInformation, "Info !"

             
         End If
         
         If lblfrm = "CUSTOMER_BR" Then
         Customer_br.TimerALL.Interval = 10
         Customer_br.TXTCARI = lblkdcustomer
         Else
         Customer.TimerALL.Interval = 10
         End If
                  
         
         Unload Me
    End If
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0


sql = "Select * from BANK order by kdbank"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
CMBbank.AddItem rs!kdbank
rs.MoveNext
Loop

Opt1.Value = True
index_PKP = 1

CMbJNSBYR.AddItem "TUNAI"
CMbJNSBYR.AddItem "TRANSFER"
CMbJNSBYR.ListIndex = 0

Call nul(lblkdcustomer)
Call nul(TXTnmcustomer)
Call nul(txtalamat)
Call nul(txtalamat_TGH)
Call nul(txtkdcustomer_IAP)
Call nul(lblnmcabang)
Call nul(lblnmsp)
Call nul(lblnosp)
Call nul(lblkdwilayah)
Call nul(lblnmwilayah)
Call nul(txtnoNPWP)
Call nul(txtnmNPWP)
Call nul(txtalamatNPWP)


If UTAMA.lblM_Master = 0 Then
cmdsimpan.Enabled = False
Else
cmdsimpan.Enabled = True
End If

TimerNO.Interval = 10

TimerCMB.Interval = 10
End Sub


Private Sub lblkdarea_Change()
sqlA = "select * from area_tagih where kdarea='" & lblkdarea & "'"
Set rsA = con.Execute(sqlA)

If rsA.RecordCount <> 0 Then
LBLNMAREA = rsA!nmarea
Else
LBLNMAREA = ""
End If

End Sub

Private Sub lblkdareaC_Change()
sqlAreaC = "select a.*,isnull(b.nmteknisi,'') as nmteknisi from area_cheker a left join teknisi b on a.kdteknisi=b.kdteknisi where a.kdareaC='" & lblkdareaC & "'"
Set rsAreaC = con.Execute(sqlAreaC)


If rsAreaC.RecordCount <> 0 Then
lblnmareaC = rsAreaC!nmareaC
lblkdteknisi = rsAreaC!kdteknisi
Else
lblnmareaC = ""
lblkdteknisi = ""
End If

End Sub

Private Sub lblkdcustomer_Change()
Call nul(lblkdcustomer)

On Error GoTo hell

Set datagrid1.DataSource = con.Execute("exec sp_pjm_swa_Mcust @kdcustomer='" & lblkdcustomer & "'")

For i = 1 To (datagrid1.Rows - 1)

datagrid1.TextMatrix(i, 0) = i

Next

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub LBLKDCUSTOMER_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub LBLKDCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub LBLKDCUSTOMER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub LBLKDCUSTOMER_LostFocus()
lblkdcustomer = UCase(lblkdcustomer)
End Sub



Private Sub lblkdsupplier_Change()

End Sub

Private Sub lblkdkolektor_Change()
sqlK = "select * from kolektor where kdkolektor='" & lblkdkolektor & "'"
Set rsK = con.Execute(sqlK)

If rsK.RecordCount <> 0 Then
lblnmkolektor = rsK!nmkolektor
Else
lblnmkolektor = ""
End If

End Sub

Private Sub lblkdPIC_Change()
sqlpic = "select * from PIC_Marketing where kdpic='" & lblkdPIC & "'"
Set rsPIC = con.Execute(sqlpic)

If rsPIC.RecordCount <> 0 Then
lblnmPIC = rsPIC!nmpic
Else
lblnmPIC = ""
End If

End Sub

Private Sub lblkdSP_Change()
On Error Resume Next
sqlSP = "select a.*,b.nmcabang from SP_IAP a left join cabang_IAP b on a.kdcabang=b.kdcabang where a.kdsp ='" & lblkdSP & "' "
Set rsSP = con.Execute(sqlSP)

If rsSP.RecordCount <> 0 Then
lblnosp = rsSP!noSP
lblnmsp = rsSP!nmsp
lblnmcabang = rsSP!nmcabang
lblkdcustomer_IAP = lblkdSP & "/" & txtkdcustomer_IAP
Else
lblnosp = ""
lblnmsp = ""
lblnmcabang = ""
lblkdcustomer_IAP = ""
End If

Call data_IAP
End Sub

Private Sub lblkdteknisi_Change()
sqlteknisi = "select * from teknisi where kdteknisi='" & lblkdteknisi & "'"
Set rsteknisi = con.Execute(sqlteknisi)

If rsteknisi.RecordCount <> 0 Then
lblnmteknisi = rsteknisi!nmteknisi
Else
lblnmteknisi = ""
End If
End Sub

Private Sub lblkdwilayah_Change()
Call nul(lblkdwilayah)

sqlwil = "select * from wilayah where kdwilayah='" & lblkdwilayah & "'"
Set rswil = con.Execute(sqlwil)


If rswil.RecordCount <> 0 Then
lblnmwilayah = rswil!nmwilayah

Else
lblnmwilayah = ""

End If

End Sub

Private Sub lblnmcabang_Change()
Call nul(lblnmcabang)
End Sub

Private Sub lblnmsp_Change()
Call nul(lblnmsp)
End Sub

Private Sub lblnmwilayah_Change()
Call nul(lblnmwilayah)
End Sub

Private Sub lblnosp_Change()
Call nul(lblnosp)
End Sub

Private Sub Text3_Change()

End Sub

Private Sub OPT1_Click(Value As Integer)
index_PKP = 1
End Sub

Private Sub OPT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Opt2_Click(Value As Integer)
index_PKP = 0
End Sub

Private Sub Opt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub TimerCMB_Timer()
If LBLKODE = "1" Then
CMBbank.ListIndex = 0
End If


TimerCMB.Interval = 0

End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If

TimerNO.Interval = 0
End Sub






Private Sub txtAlamatNPWP_Change()
Call nul(txtalamatNPWP)
End Sub

Private Sub txtAlamatNPWP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtAlamatNPWP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"

End If

End Sub

Private Sub txtAlamatNPWP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtAlamatNPWP_LostFocus()
txtalamatNPWP = UCase(txtalamatNPWP)
End Sub

Private Sub txtbtl_Change()
Call nul(txtbtl)
End Sub

Private Sub txtbtl_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtbtl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtbtl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
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

Private Sub txtbtl_LostFocus()
On Error GoTo hell

txtbtl = FormatNumber(txtbtl, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtbtl.SetFocus

End Sub

Private Sub txtcup_Change()
Call nul(txtcup)
End Sub

Private Sub txtcup_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtcup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtcup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
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

Private Sub txtcup_LostFocus()
On Error GoTo hell

txtcup = FormatNumber(txtcup, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtcup.SetFocus

End Sub

Private Sub txtgln_Change()
Call nul(txtgln)
End Sub

Private Sub txtgln_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtgln_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtgln_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
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

Private Sub txtgln_LostFocus()
On Error GoTo hell

txtgln = FormatNumber(txtgln, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txtgln.SetFocus

End Sub

Private Sub txthrgsewa_Change()
Call nul(txthrgSewa)
End Sub

Private Sub txthrgsewa_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txthrgsewa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txthrgsewa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
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

Private Sub txthrgsewa_LostFocus()
On Error GoTo hell

txthrgSewa = FormatNumber(txthrgSewa, 0)


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"
txthrgSewa.SetFocus

End Sub

Private Sub txtkdcustomer_IAP_Change()
Call nul(txtkdcustomer_IAP)
Call data_IAP
lblkdcustomer_IAP = lblkdSP & "/" & txtkdcustomer_IAP
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

Private Sub txtnmcustomer_Change()
Call nul(TXTnmcustomer)
End Sub

Private Sub txtnmcustomer_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmcustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnmcustomer_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnmcustomer_LostFocus()
TXTnmcustomer = UCase(TXTnmcustomer)
End Sub

Private Sub txtalamat_Change()
Call nul(txtalamat)

If LBLKODE = 1 Then
txtalamat_TGH = UCase(txtalamat)
End If
End Sub

Private Sub txtalamat_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtalamat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtalamat_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtalamat_LostFocus()
txtalamat = UCase(txtalamat)
End Sub


Private Sub txtalamat_TGH_Change()
Call nul(txtalamat_TGH)
End Sub

Private Sub txtalamat_TGH_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtalamat_TGH_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtalamat_TGH_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtalamat_TGH_LostFocus()
txtalamat_TGH = UCase(txtalamat_TGH)
End Sub


Private Sub txtnmNPWP_Change()
Call nul(txtnmNPWP)
End Sub

Private Sub txtnmNPWP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmNPWP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"

End If

End Sub

Private Sub txtnmNPWP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnmNPWP_LostFocus()
txtnmNPWP = UCase(txtnmNPWP)
End Sub

Private Sub txtnoNPWP_Change()
Call nul(txtnoNPWP)

End Sub

Private Sub txtnoNPWP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoNPWP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtnoNPWP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-.", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txtnoNPWP_LostFocus()
On Error GoTo hell

txtnoNPWP = UCase(txtnoNPWP)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txtnoNPWP.SetFocus

End Sub

Private Sub txtnoSPK_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnoSPK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtnoSPK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtnoSPK_LostFocus()
txtnospk = UCase(txtnospk)
End Sub

Private Sub txttelp_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttelp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttelp_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txttelp_LostFocus()
txttelp = UCase(txttelp)
End Sub

Private Sub txtCP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtCP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtCP_LostFocus()
txtCP = UCase(txtCP)
End Sub


Private Sub txtkdcustomer_IAP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtkdcustomer_IAP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txtkdcustomer_IAP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub txtkdcustomer_IAP_LostFocus()
txtkdcustomer_IAP = UCase(txtkdcustomer_IAP)
End Sub



Private Sub txttglsegel_Change()
Call nul(txttglsegel)
End Sub

Private Sub txttglsegel_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglsegel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttglsegel_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890-/", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If

End If

End Sub

Private Sub txttglsegel_LostFocus()
On Error GoTo hell

txttglsegel = FormatDateTime(txttglsegel, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttglsegel.SetFocus
End Sub

Private Sub txttglspk1_Change()
Call nul(txttglspk1)
End Sub

Private Sub txttglSPK1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglSPK1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglspk1_KeyPress(KeyAscii As Integer)
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

Private Sub txttglspk1_LostFocus()
On Error GoTo hell

txttglspk1 = FormatDateTime(txttglspk1, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglspk1.SetFocus

End Sub


Private Sub txttglSPK2_Change()
Call nul(txttglspk2)
End Sub


Private Sub txttglSPK2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglSPK2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglSPK2_KeyPress(KeyAscii As Integer)
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

Private Sub txttglSPK2_LostFocus()
On Error GoTo hell

txttglspk2 = FormatDateTime(txttglspk2, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglspk2.SetFocus

End Sub


