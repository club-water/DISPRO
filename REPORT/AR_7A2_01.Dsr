VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_7A2_01 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "AR_7A2_01.dsx":0000
End
Attribute VB_Name = "AR_7A2_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub


Private Sub Detail_BeforePrint()
On Error Resume Next

Static i As Long

i = i + 1

fldno = i & "."

fldover = CDate(lbltgl1) - CDate(fldtglSPK2)

If fldkdkategori = "04" Or fldkdkategori = "10" Then
fldstatus = "SEWA"
Else
fldstatus = "MILIK SENDIRI"
End If

If fldnoSPK = "" Then
fldtglSPK1.Visible = False
fldtglSPK2.Visible = False
fldover = 0
Else
fldtglSPK1.Visible = True
fldtglSPK2.Visible = True
fldover = CDate(lbltgl1) - CDate(fldtglSPK2)
End If

If CCur(fldover) > 0 And fldnoSPK <> "" Then
fldtglSPK1.ForeColor = vbRed
fldtglSPK2.ForeColor = vbRed
fldnoSPK.ForeColor = vbRed
fldno.ForeColor = vbRed
fldkdbarang.ForeColor = vbRed
fldnmbarang.ForeColor = vbRed
fldkdkategori.ForeColor = vbRed
fldnmkategori.ForeColor = vbRed
fldtglbpb.ForeColor = vbRed
fldkd1.ForeColor = vbRed
fldkdcustomer.ForeColor = vbRed
fldnmcus.ForeColor = vbRed
fldalamat.ForeColor = vbRed
fldstatus.ForeColor = vbRed
fldtglSJ.ForeColor = vbRed
fldpjm.ForeColor = vbRed
fldswa.ForeColor = vbRed

ElseIf CCur(fldover) <= 0 And CCur(fldover) >= 30 And fldnoSPK <> "" Then
fldtglSPK1.ForeColor = vbYellow
fldtglSPK2.ForeColor = vbYellow
fldnoSPK.ForeColor = vbYellow
fldno.ForeColor = vbYellow
fldkdbarang.ForeColor = vbYellow
fldnmbarang.ForeColor = vbYellow
fldkdkategori.ForeColor = vbYellow
fldnmkategori.ForeColor = vbYellow
fldtglbpb.ForeColor = vbYellow
fldkd1.ForeColor = vbYellow
fldkdcustomer.ForeColor = vbYellow
fldnmcus.ForeColor = vbYellow
fldalamat.ForeColor = vbYellow
fldstatus.ForeColor = vbYellow
fldtglSJ.ForeColor = vbYellow
fldpjm.ForeColor = vbYellow
fldswa.ForeColor = vbYellow

ElseIf fldnoSPK = "" Then
fldtglSPK1.ForeColor = vbBlack
fldtglSPK2.ForeColor = vbBlack
fldnoSPK.ForeColor = vbBlack
fldno.ForeColor = vbBlack
fldkdbarang.ForeColor = vbBlack
fldnmbarang.ForeColor = vbBlack
fldkdkategori.ForeColor = vbBlack
fldnmkategori.ForeColor = vbBlack
fldtglbpb.ForeColor = vbBlack
fldkd1.ForeColor = vbBlack
fldkdcustomer.ForeColor = vbBlack
fldnmcus.ForeColor = vbBlack
fldalamat.ForeColor = vbBlack
fldstatus.ForeColor = vbBlack
fldtglSJ.ForeColor = vbBlack
fldpjm.ForeColor = vbBlack
fldswa.ForeColor = vbBlack

ElseIf CCur(fldover) < 0 And fldnoSPK <> "" Then
fldtglSPK1.ForeColor = vbBlack
fldtglSPK2.ForeColor = vbBlack
fldnoSPK.ForeColor = vbBlack
fldno.ForeColor = vbBlack
fldkdbarang.ForeColor = vbBlack
fldnmbarang.ForeColor = vbBlack
fldkdkategori.ForeColor = vbBlack
fldnmkategori.ForeColor = vbBlack
fldtglbpb.ForeColor = vbBlack
fldkd1.ForeColor = vbBlack
fldkdcustomer.ForeColor = vbBlack
fldnmcus.ForeColor = vbBlack
fldalamat.ForeColor = vbBlack
fldstatus.ForeColor = vbBlack
fldtglSJ.ForeColor = vbBlack
fldpjm.ForeColor = vbBlack
fldswa.ForeColor = vbBlack

End If

End Sub





