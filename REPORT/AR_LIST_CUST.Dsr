VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_LIST_CUST 
   BorderStyle     =   0  'None
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "AR_LIST_CUST.dsx":0000
End
Attribute VB_Name = "AR_LIST_CUST"
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

fldNo = i & "."

fldumur = DateDiff("d", fldtgldibuat, Now)


If fldnmcustomer_IAP = "" And fldumur < 15 Then
fldnmcustomer_IAP.BackStyle = ddBKNormal
fldalamat_IAP.BackStyle = ddBKNormal
fldnmcustomer_IAP.BackColor = vbYellow
fldalamat_IAP.BackColor = vbYellow

ElseIf fldnmcustomer_IAP = "" And fldumur > 14 Then
fldnmcustomer_IAP.BackStyle = ddBKNormal
fldalamat_IAP.BackStyle = ddBKNormal
fldnmcustomer_IAP.BackColor = vbRed
fldalamat_IAP.BackColor = vbRed

Else
fldnmcustomer_IAP.BackStyle = ddBKTransparent
fldalamat_IAP.BackStyle = ddBKTransparent
End If


End Sub
