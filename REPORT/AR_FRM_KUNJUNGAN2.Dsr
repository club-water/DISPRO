VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_FRM_KUNJUNGAN2 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   27146
   _ExtentY        =   13335
   SectionData     =   "AR_FRM_KUNJUNGAN2.dsx":0000
End
Attribute VB_Name = "AR_FRM_KUNJUNGAN2"
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


If fldnmcustomer = "" Then
flddisp = ""
fldSH = ""
fldRG = ""
fldgln1 = ""
fldgln2 = ""
fldgln3 = ""
fldsps1 = ""
fldsps2 = ""
fldsps3 = ""
End If

If fldket_sewa <> "" Then
fldnmcustomer.BackStyle = ddBKNormal
fldnmcustomer.Font.Bold = True
fldalamat.Font.Bold = True
fldtelp.Font.Bold = True
flddisp.Font.Bold = True
fldSH.Font.Bold = True
fldRG.Font.Bold = True
fldgln1.Font.Bold = True
fldgln2.Font.Bold = True
fldgln3.Font.Bold = True
fldsps1.Font.Bold = True
fldsps2.Font.Bold = True
fldsps3.Font.Bold = True
Else
fldnmcustomer.BackStyle = ddBKTransparent
fldnmcustomer.Font.Bold = False
fldalamat.Font.Bold = False
fldtelp.Font.Bold = False
flddisp.Font.Bold = False
fldSH.Font.Bold = False
fldRG.Font.Bold = False
fldgln1.Font.Bold = False
fldgln2.Font.Bold = False
fldgln3.Font.Bold = False
fldsps1.Font.Bold = False
fldsps2.Font.Bold = False
fldsps3.Font.Bold = False

End If


End Sub







