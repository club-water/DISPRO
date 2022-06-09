VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_ANALISASH 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   22172
   _ExtentY        =   16113
   SectionData     =   "AR_ANALISASH.dsx":0000
End
Attribute VB_Name = "AR_ANALISASH"
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

If fldomsetSPS = "" Then
fldomsetSPS = 0
End If


fldrasio = FormatNumber((CLng(fldomsetSPS) / CLng(fldtotal)), 0)


If Cetak_7A5.Opt1.Value = True Then

    If flddl <> "" Then
    flddl = "X"
    Else
    flddl = ""
    End If

End If

If fldnmcust_iap = "" Then
fldkdcust_iap.BackStyle = ddBKNormal
fldnmcust_iap.BackStyle = ddBKNormal
fldalamat_iap.BackStyle = ddBKNormal

fldkdcust_iap.BackColor = vbRed
fldnmcust_iap.BackColor = vbRed
fldalamat_iap.BackColor = vbRed
Else
fldkdcust_iap.BackStyle = ddBKTransparent
fldnmcust_iap.BackStyle = ddBKTransparent
fldalamat_iap.BackStyle = ddBKTransparent

fldkdcust_iap.BackColor = vbWhite
fldnmcust_iap.BackColor = vbWhite
fldalamat_iap.BackColor = vbWhite
End If

If fldrasio < 20 Then
fldrasio.ForeColor = vbRed
Else
fldrasio.ForeColor = vbBlack
End If


End Sub
