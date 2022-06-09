VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_1A1 
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
   SectionData     =   "AR_1A1.dsx":0000
End
Attribute VB_Name = "AR_1A1"
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
Static i As Long

i = i + 1

fldno = i & "."

fldM_Total = CCur(fldU_beli) + CCur(fldU_Rpjm) + CCur(fldU_Rsewa) + CCur(fldM_Unit)
fldK_Total = CCur(fldU_free) + CCur(fldU_pjm) + CCur(fldU_sewa) + CCur(fldK_unit) + CCur(fldrepair)

If fldsakir = 0 Then
Frame1.BackColor = vbYellow
Else
Frame1.BackColor = vbWhite
End If
End Sub

