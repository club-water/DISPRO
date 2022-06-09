VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_fixrute_list 
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
   SectionData     =   "AR_fixrute_list.dsx":0000
End
Attribute VB_Name = "AR_fixrute_list"
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
Static i As Long
i = i + 1

fldNO = i & "."

If fldstatus = "SDH DIKUNJUNGI" And fldtglCek < fldtglplan Then
fldtglCek.BackStyle = ddBKNormal
fldtglCek.BackColor = &HC0C0FF
ElseIf fldstatus = "SDH DIKUNJUNGI" And fldtglCek > fldtglplan Then
fldtglCek.BackStyle = ddBKNormal
fldtglCek.BackColor = &HFF8080
Else
fldtglCek.BackStyle = ddBKTransparent
End If


If fldstatus = "BLM DIKUNJUNGI" Then
Frame1.BackColor = vbYellow
Else
Frame1.BackColor = vbWhite
End If

If fldketerangan <> "" Then
fldketerangan.ForeColor = vbRed
Else
fldketerangan.ForeColor = vbBlack
End If

End Sub
