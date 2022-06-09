VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form fixrute_S 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   2475
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttglK8 
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
      Left            =   1935
      TabIndex        =   13
      Top             =   1260
      Width           =   1590
   End
   Begin VB.CheckBox ChkK8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 8 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   12
      Top             =   1260
      Width           =   1770
   End
   Begin VB.TextBox txttglK9 
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
      Left            =   5535
      TabIndex        =   15
      Top             =   1260
      Width           =   1590
   End
   Begin VB.CheckBox ChkK9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 9 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3690
      TabIndex        =   14
      Top             =   1260
      Width           =   1770
   End
   Begin VB.TextBox txttglK10 
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
      Left            =   9135
      TabIndex        =   17
      Top             =   1260
      Width           =   1590
   End
   Begin VB.CheckBox ChkK10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 10 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7290
      TabIndex        =   16
      Top             =   1260
      Width           =   1770
   End
   Begin VB.TextBox txttglK5 
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
      Left            =   1935
      TabIndex        =   7
      Top             =   855
      Width           =   1590
   End
   Begin VB.CheckBox ChkK5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 5 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   6
      Top             =   855
      Width           =   1770
   End
   Begin VB.TextBox txttglK6 
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
      Left            =   5535
      TabIndex        =   9
      Top             =   855
      Width           =   1590
   End
   Begin VB.CheckBox ChkK6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 6 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3690
      TabIndex        =   8
      Top             =   855
      Width           =   1770
   End
   Begin VB.TextBox txttglK7 
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
      Left            =   9135
      TabIndex        =   11
      Top             =   855
      Width           =   1590
   End
   Begin VB.CheckBox ChkK7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 7 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7290
      TabIndex        =   10
      Top             =   855
      Width           =   1770
   End
   Begin VB.Timer TimerALL 
      Left            =   3150
      Top             =   2790
   End
   Begin VB.CheckBox ChkK4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 4 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7290
      TabIndex        =   4
      Top             =   450
      Width           =   1770
   End
   Begin VB.TextBox txttglK4 
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
      Left            =   9135
      TabIndex        =   5
      Top             =   450
      Width           =   1590
   End
   Begin VB.CheckBox ChkK3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 3 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3690
      TabIndex        =   2
      Top             =   450
      Width           =   1770
   End
   Begin VB.TextBox txttglK3 
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
      Left            =   5535
      TabIndex        =   3
      Top             =   450
      Width           =   1590
   End
   Begin VB.CheckBox chkK2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Caption         =   "KUNJUNGAN 2 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   1770
   End
   Begin VB.TextBox txttglK2 
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
      Left            =   1935
      TabIndex        =   1
      Top             =   450
      Width           =   1590
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   645
      Left            =   8370
      TabIndex        =   18
      ToolTipText     =   "Update"
      Top             =   1665
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1138
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "fixrute_S.frx":0000
      Caption         =   "  &Update"
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label lblkdcustomer 
      Caption         =   "lblkdcustomer"
      Height          =   285
      Left            =   1170
      TabIndex        =   21
      Top             =   3330
      Width           =   1725
   End
   Begin VB.Label lblidrute 
      Caption         =   "lblidrute"
      Height          =   330
      Left            =   1170
      TabIndex        =   20
      Top             =   2925
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunjungan Tanggal Lain :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   540
      TabIndex        =   19
      Top             =   90
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   2445
      Left            =   0
      Picture         =   "fixrute_S.frx":3CF1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "fixrute_S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim color As Long, flag As Byte
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs7 As ADODB.Recordset
Dim rs8 As ADODB.Recordset
Dim rs9 As ADODB.Recordset
Dim rs10 As ADODB.Recordset


Private Sub ALL()
sql2 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/02" & "'"
Set rs2 = con.Execute(sql2)

If rs2.RecordCount <> 0 Then
chkK2.Value = 1
txttglK2 = rs2!tglrute_S
txttglK2.Enabled = True
Else
chkK2.Value = 0
txttglK2 = "01/01/1900"
txttglK2.Enabled = False
End If

sql3 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/03" & "'"
Set rs3 = con.Execute(sql3)

If rs3.RecordCount <> 0 Then
ChkK3.Value = 1
txttglK3 = rs3!tglrute_S
txttglK3.Enabled = True
Else
ChkK3.Value = 0
txttglK3 = "01/01/1900"
txttglK3.Enabled = False
End If

sql4 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/04" & "'"
Set rs4 = con.Execute(sql4)

If rs4.RecordCount <> 0 Then
ChkK4.Value = 1
txttglK4 = rs4!tglrute_S
txttglK4.Enabled = True
Else
ChkK4.Value = 0
txttglK4 = "01/01/1900"
txttglK4.Enabled = False
End If

sql5 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/05" & "'"
Set rs5 = con.Execute(sql5)

If rs5.RecordCount <> 0 Then
ChkK5.Value = 1
txttglK5 = rs5!tglrute_S
txttglK5.Enabled = True
Else
ChkK5.Value = 0
txttglK5 = "01/01/1900"
txttglK5.Enabled = False
End If

sql6 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/06" & "'"
Set rs6 = con.Execute(sql6)

If rs6.RecordCount <> 0 Then
ChkK6.Value = 1
txttglK6 = rs6!tglrute_S
txttglK6.Enabled = True
Else
ChkK6.Value = 0
txttglK6 = "01/01/1900"
txttglK6.Enabled = False
End If

sql7 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/07" & "'"
Set rs7 = con.Execute(sql7)

If rs7.RecordCount <> 0 Then
ChkK7.Value = 1
txttglK7 = rs7!tglrute_S
txttglK7.Enabled = True
Else
ChkK7.Value = 0
txttglK7 = "01/01/1900"
txttglK7.Enabled = False
End If

sql8 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/08" & "'"
Set rs8 = con.Execute(sql8)

If rs8.RecordCount <> 0 Then
chkK8.Value = 1
txttglK8 = rs8!tglrute_S
txttglK8.Enabled = True
Else
chkK8.Value = 0
txttglK8 = "01/01/1900"
txttglK8.Enabled = False
End If

sql9 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/09" & "'"
Set rs9 = con.Execute(sql9)

If rs9.RecordCount <> 0 Then
ChkK9.Value = 1
txttglK9 = rs9!tglrute_S
txttglK9.Enabled = True
Else
ChkK9.Value = 0
txttglK9 = "01/01/1900"
txttglK9.Enabled = False
End If

sql10 = "select * from route_plan_s where idrute_s ='" & lblidrute & "/10" & "'"
Set rs10 = con.Execute(sql10)

If rs10.RecordCount <> 0 Then
ChkK10.Value = 1
txttglK10 = rs10!tglrute_S
txttglK10.Enabled = True
Else
ChkK10.Value = 0
txttglK10 = "01/01/1900"
txttglK10.Enabled = False
End If



End Sub


Private Sub chkK2_Click()
If chkK2.Value = 0 Then
txttglK2 = "01/01/1900"
txttglK2.Enabled = False
Else
txttglK2 = Date
txttglK2.Enabled = True
End If
End Sub

Private Sub chkK3_Click()
If ChkK3.Value = 0 Then
txttglK3 = "01/01/1900"
txttglK3.Enabled = False
Else
txttglK3 = Date
txttglK3.Enabled = True
End If
End Sub

Private Sub chkK4_Click()
If ChkK4.Value = 0 Then
txttglK4 = "01/01/1900"
txttglK4.Enabled = False
Else
txttglK4 = Date
txttglK4.Enabled = True
End If
End Sub

Private Sub chkK5_Click()
If ChkK5.Value = 0 Then
txttglK5 = "01/01/1900"
txttglK5.Enabled = False
Else
txttglK5 = Date
txttglK5.Enabled = True
End If
End Sub

Private Sub chkK6_Click()
If ChkK6.Value = 0 Then
txttglK6 = "01/01/1900"
txttglK6.Enabled = False
Else
txttglK6 = Date
txttglK6.Enabled = True
End If
End Sub

Private Sub chkK7_Click()
If ChkK7.Value = 0 Then
txttglK7 = "01/01/1900"
txttglK7.Enabled = False
Else
txttglK7 = Date
txttglK7.Enabled = True
End If
End Sub

Private Sub chkK8_Click()
If chkK8.Value = 0 Then
txttglK8 = "01/01/1900"
txttglK8.Enabled = False
Else
txttglK8 = Date
txttglK8.Enabled = True
End If
End Sub

Private Sub chkK9_Click()
If ChkK9.Value = 0 Then
txttglK9 = "01/01/1900"
txttglK9.Enabled = False
Else
txttglK9 = Date
txttglK9.Enabled = True
End If
End Sub

Private Sub chkK10_Click()
If ChkK10.Value = 0 Then
txttglK10 = "01/01/1900"
txttglK10.Enabled = False
Else
txttglK10 = Date
txttglK10.Enabled = True
End If
End Sub




Private Sub cmdsimpan_Click()
If chkK2.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/02" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/02" & "','" & Format(txttglK2, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/02" & "'")
End If

If ChkK3.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/03" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/03" & "','" & Format(txttglK3, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/03" & "'")
End If

If ChkK4.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/04" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/04" & "','" & Format(txttglK4, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/04" & "'")
End If


If ChkK5.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/05" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/05" & "','" & Format(txttglK5, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/05" & "'")
End If

If ChkK6.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/06" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/06" & "','" & Format(txttglK6, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/06" & "'")
End If

If ChkK7.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/07" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/07" & "','" & Format(txttglK7, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/07" & "'")
End If

If chkK8.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/08" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/08" & "','" & Format(txttglK8, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/08" & "'")
End If

If ChkK9.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/09" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/09" & "','" & Format(txttglK9, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/09" & "'")
End If

If ChkK10.Value = 1 Then
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/10" & "'")
con.Execute ("insert into route_plan_S values ('" & lblidrute & "/10" & "','" & Format(txttglK10, "yyyy/MM/dd") & "','" & lblidrute & "','" & fixrute_TU.txtperiode & "','" & fixrute_TU.lblkdteknisi & "','" & lblkdcustomer & "')")
Else
con.Execute ("delete from route_plan_S where idrute_S='" & lblidrute & "/10" & "'")
End If


fixrute_TU.flood.Visible = True
fixrute_TU.Timerflood.Interval = 10

fixrute_TU.TimerALL.Interval = 10


MsgBox "Data Berhasil di Update", vbInformation, "Info !"
Unload Me

End Sub

Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

TimerALL.Interval = 10
End Sub






Private Sub TimerALL_Timer()
On Error Resume Next

Call ALL

TimerALL.Interval = 0
End Sub

Private Sub txttglK2_Change()
Call nul(txttglK2)
End Sub

Private Sub txttglK2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK2_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK2_LostFocus()
On Error GoTo hell

txttglK2 = FormatDateTime(txttglK2, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK2.SetFocus

End Sub


Private Sub txttglK3_Change()
Call nul(txttglK3)
End Sub

Private Sub txttglK3_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK3_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK3_LostFocus()
On Error GoTo hell

txttglK3 = FormatDateTime(txttglK3, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK3.SetFocus

End Sub


Private Sub txttglK4_Change()
Call nul(txttglK4)
End Sub

Private Sub txttglK4_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK4_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK4_LostFocus()
On Error GoTo hell

txttglK4 = FormatDateTime(txttglK4, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK4.SetFocus

End Sub


Private Sub txttglK5_Change()
Call nul(txttglK5)
End Sub

Private Sub txttglK5_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK5_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK5_LostFocus()
On Error GoTo hell

txttglK5 = FormatDateTime(txttglK5, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK5.SetFocus

End Sub


Private Sub txttglK6_Change()
Call nul(txttglK6)
End Sub

Private Sub txttglK6_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK6_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK6_LostFocus()
On Error GoTo hell

txttglK6 = FormatDateTime(txttglK6, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK6.SetFocus

End Sub


Private Sub txttglK7_Change()
Call nul(txttglK7)
End Sub

Private Sub txttglK7_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK7_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK7_LostFocus()
On Error GoTo hell

txttglK7 = FormatDateTime(txttglK7, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK7.SetFocus

End Sub



Private Sub txttglK8_Change()
Call nul(txttglK8)
End Sub

Private Sub txttglK8_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK8_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK8_LostFocus()
On Error GoTo hell

txttglK8 = FormatDateTime(txttglK8, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK8.SetFocus

End Sub


Private Sub txttglK9_Change()
Call nul(txttglK9)
End Sub

Private Sub txttglK9_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK9_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK9_LostFocus()
On Error GoTo hell

txttglK9 = FormatDateTime(txttglK9, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK9.SetFocus

End Sub


Private Sub txttglK10_Change()
Call nul(txttglK10)
End Sub

Private Sub txttglK10_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglK10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If
End Sub

Private Sub txttglK10_KeyPress(KeyAscii As Integer)
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

Private Sub txttglK10_LostFocus()
On Error GoTo hell

txttglK10 = FormatDateTime(txttglK10, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglK10.SetFocus

End Sub





