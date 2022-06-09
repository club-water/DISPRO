VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_LAP_PENGIRIMAN 
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
   SectionData     =   "AR_LAP_PENGIRIMAN.dsx":0000
End
Attribute VB_Name = "AR_LAP_PENGIRIMAN"
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
If fldurut = 1 Then
fldkdPK.Visible = True
fldcustomer.Visible = True
fldketerangan.Visible = True
Else
fldkdPK.Visible = False
fldcustomer.Visible = False
fldketerangan.Visible = False
End If

If Right(fldkdbarang1, 1) = "|" Then
fldkdbarang1 = Left(fldkdbarang1, Len(fldkdbarang1) - 1)
End If

If Right(fldkdbarang2, 1) = "|" Then
fldkdbarang2 = Left(fldkdbarang2, Len(fldkdbarang2) - 1)
End If

End Sub

