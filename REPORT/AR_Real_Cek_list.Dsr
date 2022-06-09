VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_Real_Cek_list 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "AR_Real_Cek_list.dsx":0000
End
Attribute VB_Name = "AR_Real_Cek_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub

Private Sub Detail_Format()
On Error Resume Next
Static i As Long
i = i + 1

fldno = i & "."

fldselisih = DateDiff("d", fldtglplan, fldtglcek)

If CCur(fldselisih) > 0 Then
fldno.ForeColor = vbRed
fldtglcek.ForeColor = vbRed
fldtglplan.ForeColor = vbRed
fldkdcustomer.ForeColor = vbRed
fldalamat.ForeColor = vbRed
fldnmcustomer.ForeColor = vbRed
fldkdbarang.ForeColor = vbRed
fldkd1.ForeColor = vbRed
fldnmbarang.ForeColor = vbRed
fldunit.ForeColor = vbRed
fldketerangan.ForeColor = vbRed
fldselisih.ForeColor = vbRed
fldnmareaC.ForeColor = vbRed
ElseIf CCur(fldselisih) < 0 Then
fldno.ForeColor = &HC000&
fldtglcek.ForeColor = &HC000&
fldtglplan.ForeColor = &HC000&
fldkdcustomer.ForeColor = &HC000&
fldalamat.ForeColor = &HC000&
fldnmcustomer.ForeColor = &HC000&
fldkdbarang.ForeColor = &HC000&
fldkd1.ForeColor = &HC000&
fldnmbarang.ForeColor = &HC000&
fldunit.ForeColor = &HC000&
fldketerangan.ForeColor = &HC000&
fldselisih.ForeColor = &HC000&
fldnmareaC.ForeColor = &HC000&
ElseIf CCur(fldselisih) <= 0 Then
fldno.ForeColor = vbBlack
fldtglcek.ForeColor = vbBlack
fldtglplan.ForeColor = vbBlack
fldkdcustomer.ForeColor = vbBlack
fldalamat.ForeColor = vbBlack
fldnmcustomer.ForeColor = vbBlack
fldkdbarang.ForeColor = vbBlack
fldkd1.ForeColor = vbBlack
fldnmbarang.ForeColor = vbBlack
fldunit.ForeColor = vbBlack
fldketerangan.ForeColor = vbBlack
fldselisih.ForeColor = vbBlack
fldnmareaC.ForeColor = vbBlack
End If

If fldketerangan.Text <> "" Then
fldketerangan.BackStyle = ddBKNormal
fldketerangan.BackColor = vbYellow

Else
fldketerangan.BackStyle = vbTransparent
fldketerangan.BackColor = None
End If

End Sub

