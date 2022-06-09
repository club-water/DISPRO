VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_6A4 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18750
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   18750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk1 
      Caption         =   "Isi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9945
      TabIndex        =   44
      Top             =   3060
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Timer Timerxls 
      Left            =   14400
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13950
      Top             =   2295
   End
   Begin VB.Timer TimerPdf 
      Left            =   14895
      Top             =   2295
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   15885
      TabIndex        =   41
      Top             =   3060
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   262144
      ForeColor       =   255
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
      Picture         =   "Cetak_6A4.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   7230
      Left            =   360
      TabIndex        =   43
      Top             =   2970
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   12753
      SectionData     =   "Cetak_6A4.frx":6862
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   45
      Top             =   810
      Width           =   17205
      _Version        =   524288
      _ExtentX        =   30348
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   17730
      TabIndex        =   40
      ToolTipText     =   "Simpan"
      Top             =   1170
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Cetak_6A4.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   42
      ToolTipText     =   "Simpan"
      Top             =   2970
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
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
      Picture         =   "Cetak_6A4.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   46
      Top             =   10440
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
      Picture         =   "Cetak_6A4.frx":D33B
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBr1 
      Height          =   420
      Left            =   2430
      TabIndex        =   0
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":13B9D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC1 
      Height          =   420
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":163CF
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR2 
      Height          =   420
      Left            =   5895
      TabIndex        =   2
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":18A19
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC2 
      Height          =   420
      Left            =   6345
      TabIndex        =   3
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":1B24B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr3 
      Height          =   420
      Left            =   9315
      TabIndex        =   4
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":1D895
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC3 
      Height          =   420
      Left            =   9765
      TabIndex        =   5
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":200C7
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr4 
      Height          =   420
      Left            =   12780
      TabIndex        =   6
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":22711
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC4 
      Height          =   420
      Left            =   13230
      TabIndex        =   7
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   900
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
      Picture         =   "Cetak_6A4.frx":24F43
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr5 
      Height          =   420
      Left            =   2430
      TabIndex        =   8
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":2758D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc5 
      Height          =   420
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":29DBF
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr6 
      Height          =   420
      Left            =   5895
      TabIndex        =   10
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":2C409
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc6 
      Height          =   420
      Left            =   6345
      TabIndex        =   11
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":2EC3B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr7 
      Height          =   420
      Left            =   9315
      TabIndex        =   12
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":31285
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc7 
      Height          =   420
      Left            =   9765
      TabIndex        =   13
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":33AB7
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr8 
      Height          =   420
      Left            =   12780
      TabIndex        =   14
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":36101
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc8 
      Height          =   420
      Left            =   13230
      TabIndex        =   15
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1305
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
      Picture         =   "Cetak_6A4.frx":38933
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr9 
      Height          =   420
      Left            =   2430
      TabIndex        =   16
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":3AF7D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc9 
      Height          =   420
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":3D7AF
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr10 
      Height          =   420
      Left            =   5895
      TabIndex        =   18
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":3FDF9
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc10 
      Height          =   420
      Left            =   6345
      TabIndex        =   19
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":4262B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr11 
      Height          =   420
      Left            =   9315
      TabIndex        =   20
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":44C75
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc11 
      Height          =   420
      Left            =   9765
      TabIndex        =   21
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":474A7
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr12 
      Height          =   420
      Left            =   12780
      TabIndex        =   22
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":49AF1
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdc12 
      Height          =   420
      Left            =   13230
      TabIndex        =   23
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   1710
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
      Picture         =   "Cetak_6A4.frx":4C323
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR13 
      Height          =   420
      Left            =   2430
      TabIndex        =   24
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":4E96D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC13 
      Height          =   420
      Left            =   2880
      TabIndex        =   25
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":5119F
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR14 
      Height          =   420
      Left            =   5895
      TabIndex        =   26
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":537E9
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC14 
      Height          =   420
      Left            =   6345
      TabIndex        =   27
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":5601B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR15 
      Height          =   420
      Left            =   9315
      TabIndex        =   28
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":58665
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC15 
      Height          =   420
      Left            =   9765
      TabIndex        =   29
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":5AE97
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR16 
      Height          =   420
      Left            =   12780
      TabIndex        =   30
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":5D4E1
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC16 
      Height          =   420
      Left            =   13230
      TabIndex        =   31
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2115
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
      Picture         =   "Cetak_6A4.frx":5FD13
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR17 
      Height          =   420
      Left            =   2430
      TabIndex        =   32
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":6235D
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC17 
      Height          =   420
      Left            =   2880
      TabIndex        =   33
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":64B8F
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR18 
      Height          =   420
      Left            =   5895
      TabIndex        =   34
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":671D9
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC18 
      Height          =   420
      Left            =   6345
      TabIndex        =   35
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":69A0B
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR19 
      Height          =   420
      Left            =   9315
      TabIndex        =   36
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":6C055
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC19 
      Height          =   420
      Left            =   9765
      TabIndex        =   37
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":6E887
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdBR20 
      Height          =   420
      Left            =   12780
      TabIndex        =   38
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":70ED1
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC20 
      Height          =   420
      Left            =   13230
      TabIndex        =   39
      ToolTipText     =   "kosongi barang untuk menampilkan semuanya"
      Top             =   2520
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
      Picture         =   "Cetak_6A4.frx":73703
      ButtonStyle     =   4
   End
   Begin VB.Label lblkdbarang17 
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
      TabIndex        =   87
      Top             =   2565
      Width           =   1410
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 17 :"
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
      Left            =   405
      TabIndex        =   86
      Top             =   2610
      Width           =   1005
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 18 :"
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
      Left            =   3870
      TabIndex        =   85
      Top             =   2610
      Width           =   645
   End
   Begin VB.Label lblkdbarang18 
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
      Left            =   4500
      TabIndex        =   84
      Top             =   2565
      Width           =   1410
   End
   Begin VB.Label lblkdbarang20 
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
      Left            =   11385
      TabIndex        =   83
      Top             =   2565
      Width           =   1410
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 20 :"
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
      Left            =   10755
      TabIndex        =   82
      Top             =   2610
      Width           =   645
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 19 :"
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
      TabIndex        =   81
      Top             =   2610
      Width           =   1005
   End
   Begin VB.Label lblkdbarang19 
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
      Left            =   7920
      TabIndex        =   80
      Top             =   2565
      Width           =   1410
   End
   Begin VB.Label lblkdbarang15 
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
      Left            =   7920
      TabIndex        =   79
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 15 :"
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
      TabIndex        =   78
      Top             =   2205
      Width           =   1005
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 16 :"
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
      Left            =   10755
      TabIndex        =   77
      Top             =   2205
      Width           =   645
   End
   Begin VB.Label lblkdbarang16 
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
      Left            =   11385
      TabIndex        =   76
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label lblkdbarang14 
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
      Left            =   4500
      TabIndex        =   75
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 14 :"
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
      Left            =   3870
      TabIndex        =   74
      Top             =   2205
      Width           =   645
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 13 :"
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
      Left            =   405
      TabIndex        =   73
      Top             =   2205
      Width           =   1005
   End
   Begin VB.Label lblkdbarang13 
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
      TabIndex        =   72
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label lblkdbarang9 
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
      TabIndex        =   71
      Top             =   1755
      Width           =   1410
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 9 :"
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
      Left            =   405
      TabIndex        =   70
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 10 :"
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
      Left            =   3870
      TabIndex        =   69
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label lblkdbarang10 
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
      Left            =   4500
      TabIndex        =   68
      Top             =   1755
      Width           =   1410
   End
   Begin VB.Label lblkdbarang12 
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
      Left            =   11385
      TabIndex        =   67
      Top             =   1755
      Width           =   1410
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 12 :"
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
      Left            =   10755
      TabIndex        =   66
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 11 :"
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
      TabIndex        =   65
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label lblkdbarang11 
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
      Left            =   7920
      TabIndex        =   64
      Top             =   1755
      Width           =   1410
   End
   Begin VB.Label lblkdbarang5 
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
      TabIndex        =   63
      Top             =   1350
      Width           =   1410
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 5 :"
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
      Left            =   405
      TabIndex        =   62
      Top             =   1395
      Width           =   1005
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 6 :"
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
      Left            =   3870
      TabIndex        =   61
      Top             =   1395
      Width           =   645
   End
   Begin VB.Label lblkdbarang6 
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
      Left            =   4500
      TabIndex        =   60
      Top             =   1350
      Width           =   1410
   End
   Begin VB.Label lblkdbarang8 
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
      Left            =   11385
      TabIndex        =   59
      Top             =   1350
      Width           =   1410
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 8 :"
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
      Left            =   10755
      TabIndex        =   58
      Top             =   1395
      Width           =   645
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 7 :"
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
      TabIndex        =   57
      Top             =   1395
      Width           =   1005
   End
   Begin VB.Label lblkdbarang7 
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
      Left            =   7920
      TabIndex        =   56
      Top             =   1350
      Width           =   1410
   End
   Begin VB.Label lblkdbarang3 
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
      Left            =   7920
      TabIndex        =   55
      Top             =   945
      Width           =   1410
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 3 :"
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
      TabIndex        =   54
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 4 :"
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
      Left            =   10755
      TabIndex        =   53
      Top             =   990
      Width           =   645
   End
   Begin VB.Label lblkdbarang4 
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
      Left            =   11385
      TabIndex        =   52
      Top             =   945
      Width           =   1410
   End
   Begin VB.Label lblkdbarang2 
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
      Left            =   4500
      TabIndex        =   51
      Top             =   945
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 2 :"
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
      Left            =   3870
      TabIndex        =   50
      Top             =   990
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "BRG 1 :"
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
      Left            =   405
      TabIndex        =   49
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label lblkdbarang1 
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
      TabIndex        =   48
      Top             =   945
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak QR Code"
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
      Left            =   1260
      TabIndex        =   47
      Top             =   90
      Width           =   5505
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_6A4.frx":75D4D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_6A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte


Private Sub cmdBR1_Click()
Barang_BR.LBLKODE = UCase("6A4_01")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR2_Click()
Barang_BR.LBLKODE = UCase("6A4_02")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR3_Click()
Barang_BR.LBLKODE = UCase("6A4_03")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR4_Click()
Barang_BR.LBLKODE = UCase("6A4_04")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR4_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR5_Click()
Barang_BR.LBLKODE = UCase("6A4_05")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR5_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR6_Click()
Barang_BR.LBLKODE = UCase("6A4_06")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR6_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR7_Click()
Barang_BR.LBLKODE = UCase("6A4_07")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR7_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR8_Click()
Barang_BR.LBLKODE = UCase("6A4_08")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR8_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR9_Click()
Barang_BR.LBLKODE = UCase("6A4_09")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR9_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR10_Click()
Barang_BR.LBLKODE = UCase("6A4_10")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR10_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR11_Click()
Barang_BR.LBLKODE = UCase("6A4_11")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR11_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR12_Click()
Barang_BR.LBLKODE = UCase("6A4_12")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR12_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR13_Click()
Barang_BR.LBLKODE = UCase("6A4_13")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR13_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR14_Click()
Barang_BR.LBLKODE = UCase("6A4_14")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR14_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR15_Click()
Barang_BR.LBLKODE = UCase("6A4_15")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR15_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR16_Click()
Barang_BR.LBLKODE = UCase("6A4_16")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR16_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR17_Click()
Barang_BR.LBLKODE = UCase("6A4_17")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR17_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR18_Click()
Barang_BR.LBLKODE = UCase("6A4_18")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR18_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR19_Click()
Barang_BR.LBLKODE = UCase("6A4_19")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR19_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdBR20_Click()
Barang_BR.LBLKODE = UCase("6A4_20")
Barang_BR.Show vbModal
End Sub

Private Sub cmdBR20_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub cmdC1_Click()
lblkdbarang1 = ""
End Sub

Private Sub cmdC1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC13_Click()
lblkdbarang13 = ""
End Sub

Private Sub cmdC13_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC14_Click()
lblkdbarang14 = ""
End Sub

Private Sub cmdC14_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC15_Click()
lblkdbarang15 = ""
End Sub

Private Sub cmdC15_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC16_Click()
lblkdbarang16 = ""
End Sub

Private Sub cmdC16_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC17_Click()
lblkdbarang17 = ""
End Sub

Private Sub cmdC17_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC18_Click()
lblkdbarang18 = ""
End Sub

Private Sub cmdC18_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC19_Click()
lblkdbarang19 = ""
End Sub

Private Sub cmdC19_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC20_Click()
lblkdbarang20 = ""
End Sub

Private Sub cmdC20_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub





Private Sub cmdC2_Click()
lblkdbarang2 = ""
End Sub

Private Sub cmdC2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC3_Click()
lblkdbarang3 = ""
End Sub

Private Sub cmdC3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC4_Click()
lblkdbarang4 = ""
End Sub

Private Sub cmdC4_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC5_Click()
lblkdbarang5 = ""
End Sub

Private Sub cmdC5_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC6_Click()
lblkdbarang6 = ""
End Sub

Private Sub cmdC6_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC7_Click()
lblkdbarang7 = ""
End Sub

Private Sub cmdC7_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC8_Click()
lblkdbarang8 = ""
End Sub

Private Sub cmdC8_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC9_Click()
lblkdbarang9 = ""
End Sub

Private Sub cmdC9_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC10_Click()
lblkdbarang10 = ""
End Sub

Private Sub cmdC10_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC11_Click()
lblkdbarang11 = ""
End Sub

Private Sub cmdC11_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdC12_Click()
lblkdbarang12 = ""
End Sub

Private Sub cmdC12_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdCLR_Click()
lblkdbarang = ""
lblnmbarang = ""
End Sub

Private Sub cmdCLR_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub Sawal()
End Sub


Private Sub total()

sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(Rpjm) as rpjm from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as rpjm from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as rpjm from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"


sql2 = "select '1' AS KODE,a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,a.pjm,a.rpjm,(a.pjm-a.Rpjm) as sisa from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang where a.pjm-a.rpjm <>0 "



sqlT = "select kode,sum(pjm) as pjm,sum(Rpjm) as Rpjm, sum(sisa) as sisa from (" & sql2 & ") a group by kode"
Set rs = con.Execute(sqlT)

End Sub




 





Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
Call Cetak
End Sub

Private Sub CHK1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    If Chk1.Value = 1 Then
    Chk1.Value = 0
    Else
    Chk1.Value = 1
    End If
    
    Call Cetak
        
ElseIf KeyAscii = 27 Then
Unload Me
End If

End Sub


Private Sub Cetak()
On Error GoTo hell

Unload AR_6A4


With AR_6A4

Set cQrCode = New ClassQR
If lblkdbarang1 <> "" Then
.Image1.Picture = cQrCode.GetPictureQrCode(lblkdbarang1, 140, 140)
.Image1.Visible = True
Else
.Image1.Visible = False
End If

If lblkdbarang2 <> "" Then
.Image2.Picture = cQrCode.GetPictureQrCode(lblkdbarang2, 140, 140)
.Image2.Visible = True
Else
.Image2.Visible = False
End If

If lblkdbarang3 <> "" Then
.Image3.Picture = cQrCode.GetPictureQrCode(lblkdbarang3, 140, 140)
.Image3.Visible = True
Else
.Image3.Visible = False
End If

If lblkdbarang4 <> "" Then
.Image4.Picture = cQrCode.GetPictureQrCode(lblkdbarang4, 140, 140)
.Image4.Visible = True
Else
.Image4.Visible = False
End If



If lblkdbarang5 <> "" Then
.Image5.Picture = cQrCode.GetPictureQrCode(lblkdbarang5, 140, 140)
.Image5.Visible = True
Else
.Image5.Visible = False
End If



If lblkdbarang6 <> "" Then
.Image6.Picture = cQrCode.GetPictureQrCode(lblkdbarang6, 140, 140)
.Image6.Visible = True
Else
.Image6.Visible = False
End If


If lblkdbarang7 <> "" Then
.Image7.Picture = cQrCode.GetPictureQrCode(lblkdbarang7, 140, 140)
.Image7.Visible = True
Else
.Image7.Visible = False
End If


If lblkdbarang8 <> "" Then
.Image8.Picture = cQrCode.GetPictureQrCode(lblkdbarang8, 140, 140)
.Image8.Visible = True
Else
.Image8.Visible = False
End If

If lblkdbarang9 <> "" Then
.Image9.Picture = cQrCode.GetPictureQrCode(lblkdbarang9, 140, 140)
.Image9.Visible = True
Else
.Image9.Visible = False
End If


If lblkdbarang10 <> "" Then
.Image10.Picture = cQrCode.GetPictureQrCode(lblkdbarang10, 140, 140)
.Image10.Visible = True
Else
.Image10.Visible = False
End If


If lblkdbarang11 <> "" Then
.Image11.Picture = cQrCode.GetPictureQrCode(lblkdbarang11, 140, 140)
.Image11.Visible = True
Else
.Image11.Visible = False
End If


If lblkdbarang12 <> "" Then
.Image12.Picture = cQrCode.GetPictureQrCode(lblkdbarang12, 140, 140)
.Image12.Visible = True
Else
.Image12.Visible = False
End If



If lblkdbarang13 <> "" Then
.Image13.Picture = cQrCode.GetPictureQrCode(lblkdbarang13, 140, 140)
.Image13.Visible = True
Else
.Image13.Visible = False
End If



If lblkdbarang14 <> "" Then
.Image14.Picture = cQrCode.GetPictureQrCode(lblkdbarang14, 140, 140)
.Image14.Visible = True
Else
.Image14.Visible = False
End If


If lblkdbarang15 <> "" Then
.Image15.Picture = cQrCode.GetPictureQrCode(lblkdbarang15, 140, 140)
.Image15.Visible = True
Else
.Image15.Visible = False
End If


If lblkdbarang16 <> "" Then
.Image16.Picture = cQrCode.GetPictureQrCode(lblkdbarang16, 140, 140)
.Image16.Visible = True
Else
.Image16.Visible = False
End If


If lblkdbarang17 <> "" Then
.Image17.Picture = cQrCode.GetPictureQrCode(lblkdbarang17, 140, 140)
.Image17.Visible = True
Else
.Image17.Visible = False
End If



If lblkdbarang18 <> "" Then
.Image18.Picture = cQrCode.GetPictureQrCode(lblkdbarang18, 140, 140)
.Image18.Visible = True
Else
.Image18.Visible = False
End If


If lblkdbarang19 <> "" Then
.Image19.Picture = cQrCode.GetPictureQrCode(lblkdbarang19, 140, 140)
.Image19.Visible = True
Else
.Image19.Visible = False
End If


If lblkdbarang20 <> "" Then
.Image20.Picture = cQrCode.GetPictureQrCode(lblkdbarang20, 140, 140)
.Image20.Visible = True
Else
.Image20.Visible = False
End If


.fldkdbarang1.Text = lblkdbarang1
.fldkdbarang2.Text = lblkdbarang2
.fldkdbarang3.Text = lblkdbarang3
.fldkdbarang4.Text = lblkdbarang4
.fldkdbarang5.Text = lblkdbarang5
.fldkdbarang6.Text = lblkdbarang6
.fldkdbarang7.Text = lblkdbarang7
.fldkdbarang8.Text = lblkdbarang8
.fldkdbarang9.Text = lblkdbarang9
.fldkdbarang10.Text = lblkdbarang10
.fldkdbarang11.Text = lblkdbarang11
.fldkdbarang12.Text = lblkdbarang12
.fldkdbarang13.Text = lblkdbarang13
.fldkdbarang14.Text = lblkdbarang14
.fldkdbarang15.Text = lblkdbarang15
.fldkdbarang16.Text = lblkdbarang16
.fldkdbarang17.Text = lblkdbarang17
.fldkdbarang18.Text = lblkdbarang18
.fldkdbarang19.Text = lblkdbarang19
.fldkdbarang20.Text = lblkdbarang20


Set Me.ARV1.ReportSource = AR_6A4

End With

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub





Private Sub cmdBRKr_Click()
Karyawan_BR.LBLKODE = "LAD"
Karyawan_BR.Show vbModal

End Sub

Private Sub cmdBRKr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub





Private Sub cmdfs_Click()
AR_6A4.Zoom = 110
AR_6A4.Show vbModal
End Sub

Private Sub cmdfs_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_Click()
Call Cetak
ARV1.ToolbarVisible = False
ARV1.ToolbarVisible = True
End Sub

Private Sub cmdGO_Click()
Call Cetak
End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdPDF_Click()
TimerPdf.Interval = 10
End Sub

Private Sub cmdPDF_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdrtf_Click()
TimerRtf.Interval = 10
End Sub

Private Sub cmdrtf_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
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



Private Sub Form_Load()
GradientForm Me, 0

txttgl1 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lblkdbarang1_Change()
On Error Resume Next

Dim a1 As Long
Dim a2 As Long
Dim a3 As Long
Dim a4 As Long
Dim a5 As Long
Dim a6 As Long
Dim a7 As Long
Dim a8 As Long
Dim a9 As Long
Dim a10 As Long
Dim a11 As Long
Dim a12 As Long
Dim a13 As Long
Dim a14 As Long
Dim a15 As Long
Dim a16 As Long
Dim a17 As Long
Dim a18 As Long
Dim a19 As Long


If Left(UCase(lblkdbarang1), 6) = "TMP/P/" And Len(lblkdbarang1) = 11 Then

    a1 = CLng(Right(lblkdbarang1, 5)) + 1
    a2 = CLng(Right(lblkdbarang1, 5)) + 2
    a3 = CLng(Right(lblkdbarang1, 5)) + 3
    a4 = CLng(Right(lblkdbarang1, 5)) + 4
    a5 = CLng(Right(lblkdbarang1, 5)) + 5
    a6 = CLng(Right(lblkdbarang1, 5)) + 6
    a7 = CLng(Right(lblkdbarang1, 5)) + 7
    a8 = CLng(Right(lblkdbarang1, 5)) + 8
    a9 = CLng(Right(lblkdbarang1, 5)) + 9
    a10 = CLng(Right(lblkdbarang1, 5)) + 10
    a11 = CLng(Right(lblkdbarang1, 5)) + 11
    a12 = CLng(Right(lblkdbarang1, 5)) + 12
    a13 = CLng(Right(lblkdbarang1, 5)) + 13
    a14 = CLng(Right(lblkdbarang1, 5)) + 14
    a15 = CLng(Right(lblkdbarang1, 5)) + 15
    a16 = CLng(Right(lblkdbarang1, 5)) + 16
    a17 = CLng(Right(lblkdbarang1, 5)) + 17
    a18 = CLng(Right(lblkdbarang1, 5)) + 18
    a19 = CLng(Right(lblkdbarang1, 5)) + 19
    
    
    Select Case Len(CStr(a1))
    Case 1
        lblkdbarang2 = "TMP/P/0000" & (a1)
    Case 2
        lblkdbarang2 = "TMP/P/000" & (a1)
    Case 3
        lblkdbarang2 = "TMP/P/00" & (a1)
    Case 4
        lblkdbarang2 = "TMP/P/0" & (a1)
    Case 5
        lblkdbarang2 = "TMP/P/" & (a1)
    
    End Select
    
    
    Select Case Len(CStr(a2))
    Case 1
        lblkdbarang3 = "TMP/P/0000" & (a2)
    Case 2
        lblkdbarang3 = "TMP/P/000" & (a2)
    Case 3
        lblkdbarang3 = "TMP/P/00" & (a2)
    Case 4
        lblkdbarang3 = "TMP/P/0" & (a2)
    Case 5
        lblkdbarang3 = "TMP/P/" & (a2)
    
    End Select

    Select Case Len(CStr(a3))
    Case 1
        lblkdbarang4 = "TMP/P/0000" & (a3)
    Case 2
        lblkdbarang4 = "TMP/P/000" & (a3)
    Case 3
        lblkdbarang4 = "TMP/P/00" & (a3)
    Case 4
        lblkdbarang4 = "TMP/P/0" & (a3)
    Case 5
        lblkdbarang4 = "TMP/P/" & (a3)
    
    End Select
    
    Select Case Len(CStr(a4))
    Case 1
        lblkdbarang5 = "TMP/P/0000" & (a4)
    Case 2
        lblkdbarang5 = "TMP/P/000" & (a4)
    Case 3
        lblkdbarang5 = "TMP/P/00" & (a4)
    Case 4
        lblkdbarang5 = "TMP/P/0" & (a4)
    Case 5
        lblkdbarang5 = "TMP/P/" & (a4)
    
    End Select

    Select Case Len(CStr(a5))
    Case 1
        lblkdbarang6 = "TMP/P/0000" & (a5)
    Case 2
        lblkdbarang6 = "TMP/P/000" & (a5)
    Case 3
        lblkdbarang6 = "TMP/P/00" & (a5)
    Case 4
        lblkdbarang6 = "TMP/P/0" & (a5)
    Case 5
        lblkdbarang6 = "TMP/P/" & (a5)
    
    End Select
    
    
    Select Case Len(CStr(a6))
    Case 1
        lblkdbarang7 = "TMP/P/0000" & (a6)
    Case 2
        lblkdbarang7 = "TMP/P/000" & (a6)
    Case 3
        lblkdbarang7 = "TMP/P/00" & (a6)
    Case 4
        lblkdbarang7 = "TMP/P/0" & (a6)
    Case 5
        lblkdbarang7 = "TMP/P/" & (a6)
    
    End Select
    
    Select Case Len(CStr(a7))
    Case 1
        lblkdbarang8 = "TMP/P/0000" & (a7)
    Case 2
        lblkdbarang8 = "TMP/P/000" & (a7)
    Case 3
        lblkdbarang8 = "TMP/P/00" & (a7)
    Case 4
        lblkdbarang8 = "TMP/P/0" & (a7)
    Case 5
        lblkdbarang8 = "TMP/P/" & (a7)
    
    End Select
    
    
    Select Case Len(CStr(a8))
    Case 1
        lblkdbarang9 = "TMP/P/0000" & (a8)
    Case 2
        lblkdbarang9 = "TMP/P/000" & (a8)
    Case 3
        lblkdbarang9 = "TMP/P/00" & (a8)
    Case 4
        lblkdbarang9 = "TMP/P/0" & (a8)
    Case 5
        lblkdbarang9 = "TMP/P/" & (a8)
    
    End Select
    
    
    Select Case Len(CStr(a9))
    Case 1
        lblkdbarang10 = "TMP/P/0000" & (a9)
    Case 2
        lblkdbarang10 = "TMP/P/000" & (a9)
    Case 3
        lblkdbarang10 = "TMP/P/00" & (a9)
    Case 4
        lblkdbarang10 = "TMP/P/0" & (a9)
    Case 5
        lblkdbarang10 = "TMP/P/" & (a9)
    
    End Select
    
    Select Case Len(CStr(a10))
    Case 1
        lblkdbarang11 = "TMP/P/0000" & (a10)
    Case 2
        lblkdbarang11 = "TMP/P/000" & (a10)
    Case 3
        lblkdbarang11 = "TMP/P/00" & (a10)
    Case 4
        lblkdbarang11 = "TMP/P/0" & (a10)
    Case 5
        lblkdbarang11 = "TMP/P/" & (a10)
    
    End Select
    
    Select Case Len(CStr(a11))
    Case 1
        lblkdbarang12 = "TMP/P/0000" & (a11)
    Case 2
        lblkdbarang12 = "TMP/P/000" & (a11)
    Case 3
        lblkdbarang12 = "TMP/P/00" & (a11)
    Case 4
        lblkdbarang12 = "TMP/P/0" & (a11)
    Case 5
        lblkdbarang12 = "TMP/P/" & (a11)
    
    End Select
    
    Select Case Len(CStr(a12))
    Case 1
        lblkdbarang13 = "TMP/P/0000" & (a12)
    Case 2
        lblkdbarang13 = "TMP/P/000" & (a12)
    Case 3
        lblkdbarang13 = "TMP/P/00" & (a12)
    Case 4
        lblkdbarang13 = "TMP/P/0" & (a12)
    Case 5
        lblkdbarang13 = "TMP/P/" & (a12)
    
    End Select
    
    Select Case Len(CStr(a13))
    Case 1
        lblkdbarang14 = "TMP/P/0000" & (a13)
    Case 2
        lblkdbarang14 = "TMP/P/000" & (a13)
    Case 3
        lblkdbarang14 = "TMP/P/00" & (a13)
    Case 4
        lblkdbarang14 = "TMP/P/0" & (a13)
    Case 5
        lblkdbarang14 = "TMP/P/" & (a13)
    
    End Select
    
    Select Case Len(CStr(a14))
    Case 1
        lblkdbarang15 = "TMP/P/0000" & (a14)
    Case 2
        lblkdbarang15 = "TMP/P/000" & (a14)
    Case 3
        lblkdbarang15 = "TMP/P/00" & (a14)
    Case 4
        lblkdbarang15 = "TMP/P/0" & (a14)
    Case 5
        lblkdbarang15 = "TMP/P/" & (a14)
    
    End Select
    
    
    Select Case Len(CStr(a15))
    Case 1
        lblkdbarang16 = "TMP/P/0000" & (a15)
    Case 2
        lblkdbarang16 = "TMP/P/000" & (a15)
    Case 3
        lblkdbarang16 = "TMP/P/00" & (a15)
    Case 4
        lblkdbarang16 = "TMP/P/0" & (a15)
    Case 5
        lblkdbarang16 = "TMP/P/" & (a15)
    
    End Select
    
    Select Case Len(CStr(a16))
    Case 1
        lblkdbarang17 = "TMP/P/0000" & (a16)
    Case 2
        lblkdbarang17 = "TMP/P/000" & (a16)
    Case 3
        lblkdbarang17 = "TMP/P/00" & (a16)
    Case 4
        lblkdbarang17 = "TMP/P/0" & (a16)
    Case 5
        lblkdbarang17 = "TMP/P/" & (a16)
    
    End Select
    
    Select Case Len(CStr(a17))
    Case 1
        lblkdbarang18 = "TMP/P/0000" & (a17)
    Case 2
        lblkdbarang18 = "TMP/P/000" & (a17)
    Case 3
        lblkdbarang18 = "TMP/P/00" & (a17)
    Case 4
        lblkdbarang18 = "TMP/P/0" & (a17)
    Case 5
        lblkdbarang18 = "TMP/P/" & (a17)
    
    End Select
    
    
    Select Case Len(CStr(a18))
    Case 1
        lblkdbarang19 = "TMP/P/0000" & (a18)
    Case 2
        lblkdbarang19 = "TMP/P/000" & (a18)
    Case 3
        lblkdbarang19 = "TMP/P/00" & (a18)
    Case 4
        lblkdbarang19 = "TMP/P/0" & (a18)
    Case 5
        lblkdbarang19 = "TMP/P/" & (a18)
    
    End Select
    
    Select Case Len(CStr(a19))
    Case 1
        lblkdbarang20 = "TMP/P/0000" & (a19)
    Case 2
        lblkdbarang20 = "TMP/P/000" & (a19)
    Case 3
        lblkdbarang20 = "TMP/P/00" & (a19)
    Case 4
        lblkdbarang20 = "TMP/P/0" & (a19)
    Case 5
        lblkdbarang20 = "TMP/P/" & (a19)
    
    End Select
    
    
End If
End Sub

Private Sub TimerPDF_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF

out2 = out2 + 1

Call save_out
pdf.filename = alamat_save & "\outfile" & CStr(out2) & ".pdf"
pdf.Export ARV1.Pages

Call EX_PDF(Me)
TimerPdf.Interval = 0

Exit Sub
hell:
TimerPdf.Interval = 0
If out2 < 10 Then
cmdPDF_Click
End If

End Sub

Private Sub Timerrtf_Timer()
On Error GoTo hell
Dim rtf As New ActiveReportsRTFExport.ARExportRTF
out = out + 1

Call save_out
rtf.filename = alamat_save & "\outfile" & CStr(out) & ".rtf"
rtf.Export ARV1.Pages

Call EX_WORD(Me)
TimerRtf.Interval = 0

Exit Sub
hell:
TimerRtf.Interval = 0
If out < 10 Then
cmdrtf_Click
End If
End Sub

Private Sub Timerxls_Timer()
On Error GoTo hell
Dim xls As New ActiveReportsExcelExport.ARExportExcel



out1 = out1 + 1

Call save_out
xls.filename = alamat_save & "\outfile" & CStr(out1) & ".xls"
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















