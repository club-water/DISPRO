VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Grafik_D1 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerH 
      Left            =   9855
      Top             =   5130
   End
   Begin VB.Timer TimerALL 
      Left            =   9270
      Top             =   5130
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D2 
      Height          =   30
      Left            =   3510
      TabIndex        =   31
      Top             =   1260
      Width           =   7215
      _Version        =   524288
      _ExtentX        =   12726
      _ExtentY        =   53
      _StockProps     =   8
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   3750
      Left            =   5850
      TabIndex        =   22
      Top             =   1260
      Width           =   30
      _Version        =   524288
      _ExtentX        =   53
      _ExtentY        =   6615
      _StockProps     =   8
      ForeColor       =   65280
      ShadowColor     =   65280
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D4 
      Height          =   3750
      Left            =   8505
      TabIndex        =   32
      Top             =   1260
      Width           =   30
      _Version        =   524288
      _ExtentX        =   53
      _ExtentY        =   6615
      _StockProps     =   8
      ForeColor       =   65280
      ShadowColor     =   65280
   End
   Begin MSComCtl2.DTPicker DTPdata 
      Height          =   330
      Left            =   9180
      TabIndex        =   43
      Top             =   45
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16761024
      CheckBox        =   -1  'True
      CustomFormat    =   "dd / MM / yyyy"
      Format          =   92274689
      CurrentDate     =   43923
   End
   Begin VB.Label lbldata_perTgl 
      Caption         =   "lbldata_perTgl"
      Height          =   285
      Left            =   3825
      TabIndex        =   55
      Top             =   5625
      Width           =   1095
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   180
      TabIndex        =   54
      Top             =   4320
      Width           =   2085
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "RATA Qyt/Hari"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   180
      TabIndex        =   53
      Top             =   4680
      Width           =   2670
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3120
      TabIndex        =   52
      Top             =   4275
      Width           =   240
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3120
      TabIndex        =   51
      Top             =   4680
      Width           =   240
   End
   Begin VB.Label lbljmlQty 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3570
      TabIndex        =   50
      Top             =   4275
      Width           =   1995
   End
   Begin VB.Label lblQTY_per_Hr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3570
      TabIndex        =   49
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lbljmlQTY1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5955
      TabIndex        =   48
      Top             =   4275
      Width           =   2490
   End
   Begin VB.Label lblQTY_per_Hr1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6225
      TabIndex        =   47
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lblH8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8835
      TabIndex        =   46
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lblH7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8610
      TabIndex        =   45
      Top             =   4275
      Width           =   2445
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Per Tanggal :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   7740
      TabIndex        =   44
      Top             =   45
      Width           =   1500
   End
   Begin VB.Label lblnmrute 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planning"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   465
      Left            =   135
      TabIndex        =   42
      Top             =   720
      Width           =   2850
   End
   Begin VB.Label lblnmteknisi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planning"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   135
      TabIndex        =   41
      Top             =   90
      Width           =   10860
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   8325
      TabIndex        =   40
      Top             =   720
      Width           =   2850
   End
   Begin VB.Label lblh1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8820
      TabIndex        =   39
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label LBLH2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8820
      TabIndex        =   38
      Top             =   1845
      Width           =   1995
   End
   Begin VB.Label LBLH3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8820
      TabIndex        =   37
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label lblh4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8820
      TabIndex        =   36
      Top             =   3060
      Width           =   1995
   End
   Begin VB.Label lblh5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8595
      TabIndex        =   35
      Top             =   3465
      Width           =   2445
   End
   Begin VB.Label lblh6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8820
      TabIndex        =   34
      Top             =   3870
      Width           =   1995
   End
   Begin VB.Label lblkdteknisi 
      Caption         =   "lblkdteknisi"
      Height          =   330
      Left            =   2025
      TabIndex        =   33
      Top             =   5670
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblcust_per_hr1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6210
      TabIndex        =   30
      Top             =   3870
      Width           =   1995
   End
   Begin VB.Label lbljmlcustomer1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5940
      TabIndex        =   29
      Top             =   3465
      Width           =   2490
   End
   Begin VB.Label lblhr_kjgn1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6210
      TabIndex        =   28
      Top             =   3060
      Width           =   1995
   End
   Begin VB.Label lbloff1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6210
      TabIndex        =   27
      Top             =   2655
      Width           =   1995
   End
   Begin VB.Label lbljmlhari1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6210
      TabIndex        =   26
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label lbltglakhir1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "01/08/2021"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6210
      TabIndex        =   25
      Top             =   1845
      Width           =   1995
   End
   Begin VB.Label lbltglawal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20/05/2021"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6210
      TabIndex        =   24
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Realisasi"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   5715
      TabIndex        =   23
      Top             =   720
      Width           =   2850
   End
   Begin VB.Label lblcust_per_hr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   21
      Top             =   3870
      Width           =   1995
   End
   Begin VB.Label lbljmlcustomer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   20
      Top             =   3465
      Width           =   1995
   End
   Begin VB.Label lblhr_kjgn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   19
      Top             =   3060
      Width           =   1995
   End
   Begin VB.Label lbloff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   18
      Top             =   2655
      Width           =   1995
   End
   Begin VB.Label lbljmlhari 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   17
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label lbltglakhir 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "01/08/2021"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   16
      Top             =   1845
      Width           =   1995
   End
   Begin VB.Label lbltglawal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20/05/2021"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3555
      TabIndex        =   15
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   14
      Top             =   3870
      Width           =   240
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   13
      Top             =   3465
      Width           =   240
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   12
      Top             =   3060
      Width           =   240
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   11
      Top             =   2610
      Width           =   240
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   10
      Top             =   2205
      Width           =   240
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   9
      Top             =   1845
      Width           =   240
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3105
      TabIndex        =   8
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "RATA Jml Cust/Hari"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   7
      Top             =   3870
      Width           =   2670
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   6
      Top             =   3465
      Width           =   2085
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Kunjungan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   5
      Top             =   3060
      Width           =   2085
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Off Kunjungan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   4
      Top             =   2655
      Width           =   1995
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lama"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   3
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Akhir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   2
      Top             =   1845
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Awal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planning"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   3060
      TabIndex        =   0
      Top             =   720
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   5145
      Left            =   0
      Picture         =   "Grafik_D1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11355
   End
End
Attribute VB_Name = "Grafik_D1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsO As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsQ As ADODB.Recordset
Dim rsR As ADODB.Recordset
Dim sqlQ1, sqlQ, sqlR1, sqlR As String
Dim x, x1, y, y1, H7 As Currency

Private Sub all()
MousePointer = vbHourglass

sql = "select * from V_rekap_plan_vs_real where nmrute='" & Grafik_Kunjungan_Cheker.txtperiode & "' and kdteknisi='" & lblkdteknisi & "'"
Set rs = con.Execute(sql)


If rs.RecordCount <> 0 Then

sqlO1 = "select '1' as kode,tgloff from OFF_kunjungan_all union all select '1' as kode,tgloFF_C as tgloff from OFF_kunjungan_C where kdteknisi='" & lblkdteknisi & "'"
sqlO2 = "select kode,tgloff from (" & sqlO1 & ") x where tgloff between '" & Format(rs!tglawal, "yyyy/MM/dd") & "' and '" & Format(rs!tglakhir, "yyyy/MM/dd") & "'  group by kode,tgloff"
sqlO = "select kode,count(tgloff) as jmloff from (" & sqlO2 & ") x group by kode"
Set rsO = con.Execute(sqlO)


sqlP1 = "select '1' as kode,tgloff from OFF_kunjungan_all union all select '1' as kode,tgloFF_C as tgloff from OFF_kunjungan_C where kdteknisi='" & lblkdteknisi & "'"
sqlP2 = "select kode,tgloff from (" & sqlP1 & ") x where tgloff between '" & Format(rs!tglawal1, "yyyy/MM/dd") & "' and '" & Format(rs!tglakhir1, "yyyy/MM/dd") & "'  group by kode,tgloff"
sqlP = "select kode,count(tgloff) as jmloff from (" & sqlP2 & ") x group by kode"
Set rsP = con.Execute(sqlP)

x = (CCur(rs!jmlcustomer1) / CCur(rs!jmlcustomer)) * 100

x1 = 100 - FormatNumber(x, 1)


sqlQ1 = "select '1' as kode,a.kdcustomer,b.qty from route_plan a left join (select kdcustomer,sum(unit-Runit)as qty from V_brg_split where tgl <= '" & Format(lbldata_perTgl, "yyyy/MM/dd") & "' group by kdcustomer) b on a.kdcustomer=b.kdcustomer where a.nmrute='" & lblnmrute & "' and kdteknisi='" & lblkdteknisi & "' "
sqlQ = "select kode,sum(qty) as qty from (" & sqlQ1 & ") Q group by kode"
Set rsQ = con.Execute(sqlQ)

sqlR1 = "select '1' as kode,kdcustomer,kdbarang,1 as qty1 from real_cek where nmrute='" & lblnmrute & "' and kdteknisi='" & lblkdteknisi & "' and tglcek <= getdate() "
sqlR = "select kode,sum(qty1) as qty1 from (" & sqlR1 & ") R group by kode"
Set rsR = con.Execute(sqlR)

y = (CCur(rsR!qty1) / CCur(rsQ!qty)) * 100

y1 = 100 - FormatNumber(y, 1)



lbltglawal = rs!tglawal
lbltglawal1 = rs!tglawal1
lbltglakhir = rs!tglakhir
lbltglakhir1 = rs!tglakhir1
lbljmlhari = FormatNumber(rs!jmlhari, 0)
lbljmlhari1 = FormatNumber(rs!jmlhari1, 0)
lbljmlcustomer = FormatNumber(rs!jmlcustomer, 0) & " (100%)"
lbljmlcustomer1 = FormatNumber(rs!jmlcustomer1, 0) & " (" & FormatNumber(x, 1) & "%)"
lbloff = rsO!jmloff

    If rsP.RecordCount <> 0 Then
    lbloff1 = rsP!jmloff
    Else
    lbloff1 = 0
    End If

lblh1 = FormatNumber(rs!h1, 0)
LBLH2 = FormatNumber(rs!h2, 0)
LBLH3 = FormatNumber(rs!h3, 0)
lblh5 = FormatNumber(rs!h5, 0) & " (" & FormatNumber(x1, 1) & "%)"

lbljmlQty = FormatNumber(rsQ!qty, 0) & " (100%)"
lbljmlQTY1 = FormatNumber(rsR!qty1, 0) & " (" & FormatNumber(y, 1) & "%)"

lblQTY_per_Hr = FormatNumber(CCur(rsQ!qty) / CCur(lbljmlhari), 0)
lblQTY_per_Hr1 = FormatNumber(CCur(rsR!qty1) / CCur(lbljmlhari1), 0)

lblH7 = CCur(rsR!qty1) - CCur(rsQ!qty) & " (" & FormatNumber(y1, 1) & "%)"
Else
lbltglawal = "-"
lbltglawal1 = "-"
lbltglakhir = "-"
lbltglakhir1 = "-"
lbljmlhari = 0
lbljmlhari1 = 0
lbljmlcustomer = 0
lbljmlcustomer1 = 0
lbloff = 0
lbloff1 = 0
lblh1 = 0
LBLH2 = 0
LBLH3 = 0
lblh5 = 0


lbljmlQty = 0
lblH7 = 0

lblQTY_per_Hr = 0
lblQTY_per_Hr1 = 0
End If

MousePointer = vbDefault

End Sub


Private Sub all1()
MousePointer = vbHourglass

sql = "exec SP_analisa_cheker @tgl1='" & Format(DTPdata, "yyyy/MM/dd") & "',@nmrute='" & Grafik_Kunjungan_Cheker.txtperiode & "',@kdteknisi='" & lblkdteknisi & "' "
Set rs = con.Execute(sql)


If rs.RecordCount <> 0 Then

sqlO1 = "select '1' as kode,tgloff from OFF_kunjungan_all union all select '1' as kode,tgloFF_C as tgloff from OFF_kunjungan_C where kdteknisi='" & lblkdteknisi & "'"
sqlO2 = "select kode,tgloff from (" & sqlO1 & ") x where tgloff between '" & Format(rs!tglawal, "yyyy/MM/dd") & "' and '" & Format(rs!tglakhir, "yyyy/MM/dd") & "'  group by kode,tgloff"
sqlO = "select kode,count(tgloff) as jmloff from (" & sqlO2 & ") x group by kode"
Set rsO = con.Execute(sqlO)


sqlP1 = "select '1' as kode,tgloff from OFF_kunjungan_all union all select '1' as kode,tgloFF_C as tgloff from OFF_kunjungan_C where kdteknisi='" & lblkdteknisi & "'"
sqlP2 = "select kode,tgloff from (" & sqlP1 & ") x where tgloff between '" & Format(rs!tglawal1, "yyyy/MM/dd") & "' and '" & Format(rs!tglakhir1, "yyyy/MM/dd") & "'  group by kode,tgloff"
sqlP = "select kode,count(tgloff) as jmloff from (" & sqlP2 & ") x group by kode"
Set rsP = con.Execute(sqlP)

x = (CCur(rs!jmlcustomer1) / CCur(rs!jmlcustomer)) * 100

x1 = 100 - FormatNumber(x, 1)



sqlQ1 = "select '1' as kode,a.kdcustomer,b.qty from route_plan a left join (select kdcustomer,sum(unit-Runit)as qty from V_brg_split where tgl <= '" & Format(lbldata_perTgl, "yyyy/MM/dd") & "' group by kdcustomer) b on a.kdcustomer=b.kdcustomer where a.nmrute='" & lblnmrute & "' and kdteknisi='" & lblkdteknisi & "' and tglplan <= '" & Format(DTPdata, "yyyy/MM/dd") & "' "
sqlQ = "select kode,sum(qty) as qty from (" & sqlQ1 & ") Q group by kode"
Set rsQ = con.Execute(sqlQ)

Text1 = sqlQ

sqlR1 = "select '1' as kode,kdcustomer,kdbarang,1 as qty1 from real_cek where nmrute='" & lblnmrute & "' and kdteknisi='" & lblkdteknisi & "' and tglcek <= '" & Format(DTPdata, "yyyy/MM/dd") & "' "
sqlR = "select kode,sum(qty1) as qty1 from (" & sqlR1 & ") R group by kode"
Set rsR = con.Execute(sqlR)

y = (CCur(rsR!qty1) / CCur(rsQ!qty)) * 100

y1 = 100 - FormatNumber(y, 1)


lbltglawal = rs!tglawal
lbltglawal1 = rs!tglawal1
lbltglakhir = rs!tglakhir
lbltglakhir1 = rs!tglakhir1
lbljmlhari = FormatNumber(rs!jmlhari, 0)
lbljmlhari1 = FormatNumber(rs!jmlhari1, 0)
lbljmlcustomer = FormatNumber(rs!jmlcustomer, 0) & " (100%)"
lbljmlcustomer1 = FormatNumber(rs!jmlcustomer1, 0) & " (" & FormatNumber(x, 1) & "%)"
lbloff = rsO!jmloff
lbloff1 = rsP!jmloff
lblh1 = FormatNumber(rs!h1, 0)
LBLH2 = FormatNumber(rs!h2, 0)
LBLH3 = FormatNumber(rs!h3, 0)
lblh5 = FormatNumber(rs!h5, 0) & " (" & FormatNumber(x1, 1) & "%)"

Text1 = sqlQ

lbljmlQty = FormatNumber(rsQ!qty, 0) & " (100%)"
lbljmlQTY1 = FormatNumber(rsR!qty1, 0) & " (" & FormatNumber(y, 1) & "%)"

lblQTY_per_Hr = FormatNumber(CCur(rsQ!qty) / CCur(lbljmlhari), 0)
lblQTY_per_Hr1 = FormatNumber(CCur(rsR!qty1) / CCur(lbljmlhari1), 0)

lblH7 = CCur(rsR!qty1) - CCur(rsQ!qty) & " (" & FormatNumber(y1, 1) & "%)"

Else
lbltglawal = "-"
lbltglawal1 = "-"
lbltglakhir = "-"
lbltglakhir1 = "-"
lbljmlhari = 0
lbljmlhari1 = 0
lbljmlcustomer = 0
lbljmlcustomer1 = 0
lbloff = 0
lbloff1 = 0
lblh1 = 0
LBLH2 = 0
LBLH3 = 0
lblh5 = 0

lbljmlQty = 0
lblH7 = 0

lblQTY_per_Hr = 0
lblQTY_per_Hr1 = 0

End If

MousePointer = vbDefault
End Sub





Private Sub DTPdata_Change()
TimerALL.Interval = 10
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
DTPdata.Value = Date
DTPdata.Value = Null

TimerALL.Interval = 10

End Sub

Private Sub TimerAll_Timer()
On Error Resume Next
If IsNull(DTPdata.Value) Then
Call all
Else
Call all1
End If
TimerH.Interval = 10

TimerALL.Interval = 0


MousePointer = vbDefault

'Exit Sub
'hell:
'MsgBox err.Description
'TimerALL.Interval = 0
'MousePointer = vbDefault
End Sub

Private Sub TimerH_Timer()
On Error Resume Next

If lblh1 > 0 Then
lblh1.ForeColor = vbRed
ElseIf lblh1 < 0 Then
lblh1.ForeColor = vbBlue
Else
lblh1.ForeColor = vbWhite
End If

If LBLH2 > 0 Then
LBLH2.ForeColor = vbRed
ElseIf lblh1 < 0 Then
LBLH2.ForeColor = vbBlue
Else
LBLH2.ForeColor = vbWhite
End If

If LBLH3 > 0 Then
LBLH3.ForeColor = vbRed
ElseIf LBLH3 < 0 Then
LBLH3.ForeColor = vbBlue
Else
LBLH3.ForeColor = vbWhite
End If


lblhr_kjgn = CLng(lbljmlhari) - CLng(lbloff)
lblhr_kjgn1 = CLng(lbljmlhari1) - CLng(lbloff1)

lblh4 = CLng(lblhr_kjgn1) - CLng(lblhr_kjgn)
If lblh4 > 0 Then
lblh4.ForeColor = vbRed
ElseIf lblh4 < 0 Then
lblh4.ForeColor = vbBlue
Else
lblh4.ForeColor = vbWhite
End If


If rs!h5 > 0 Then
lblh5.ForeColor = vbBlue
ElseIf rs!h5 < 0 Then
lblh5.ForeColor = vbRed
Else
lblh5.ForeColor = vbWhite
End If

lblcust_per_hr = FormatNumber(CLng(rs!jmlcustomer) / CLng(lblhr_kjgn), 0)
lblcust_per_hr1 = FormatNumber(CLng(rs!jmlcustomer1) / CLng(lblhr_kjgn1), 0)

lblh6 = CLng(lblcust_per_hr1) - CLng(lblcust_per_hr)
If lblh6 > 0 Then
lblh6.ForeColor = vbBlue
ElseIf lblh6 < 0 Then
lblh6.ForeColor = vbRed
Else
lblh6.ForeColor = vbWhite
End If

H7 = CLng(rsR!qty1) - CLng(rsQ!qty)
If H7 > 0 Then
lblH7.ForeColor = vbBlue
ElseIf H7 < 0 Then
lblH7.ForeColor = vbRed
Else
lblH7.ForeColor = vbWhite
End If

lblH8 = lblQTY_per_Hr1 - lblQTY_per_Hr



TimerH.Interval = 0
End Sub
