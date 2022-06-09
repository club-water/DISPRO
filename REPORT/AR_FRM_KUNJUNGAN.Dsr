VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_FRM_KUNJUNGAN 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ControlBox      =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "AR_FRM_KUNJUNGAN.dsx":0000
End
Attribute VB_Name = "AR_FRM_KUNJUNGAN"
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
fldqty_GLN = ""
fldQTY_SPS = ""
End If

If fldket_sewa <> "" Then
fldnmcustomer.BackStyle = ddBKNormal
fldnmcustomer.Font.Bold = True
fldalamat.Font.Bold = True
fldtelp.Font.Bold = True
flddisp.Font.Bold = True
fldSH.Font.Bold = True
fldRG.Font.Bold = True
fldqty_GLN.Font.Bold = True
fldQTY_SPS.Font.Bold = True
Else
fldnmcustomer.BackStyle = ddBKTransparent
fldnmcustomer.Font.Bold = False
fldalamat.Font.Bold = False
fldtelp.Font.Bold = False
flddisp.Font.Bold = False
fldSH.Font.Bold = False
fldRG.Font.Bold = False
fldqty_GLN.Font.Bold = False
fldQTY_SPS.Font.Bold = False

End If


End Sub






