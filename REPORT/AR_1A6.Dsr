VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_1A6 
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
   SectionData     =   "AR_1A6.dsx":0000
End
Attribute VB_Name = "AR_1A6"
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

If fldkdgudang1 = "" Then
fldX = ""
fldtglscan.Visible = False
Else
fldX = "X"
fldtglscan.Visible = True
End If

If fldkdgudang1 <> "" Then
   If fldkdgudang1 <> fldkdgudang And Len(fldkdgudang1) > 3 Then
   fldnmgudang1 = fldket
   End If
End If
   
If fldkdgudang1 <> "" Then
   If fldkdgudang1 <> fldkdgudang Then
   fldnmgudang.ForeColor = vbRed
   fldnmgudang1.ForeColor = vbRed
   fldX.ForeColor = vbRed
   fldkdbarang.ForeColor = vbRed
   fldnmbarang.ForeColor = vbRed
   fldkd1.ForeColor = vbRed
   fldunit.ForeColor = vbRed
   fldtglscan.ForeColor = vbRed
   fldno.ForeColor = vbRed
   Else
   fldnmgudang.ForeColor = vbBlack
   fldnmgudang1.ForeColor = vbBlack
   fldX.ForeColor = vbBlack
   fldkdbarang.ForeColor = vbBlack
   fldnmbarang.ForeColor = vbBlack
   fldkd1.ForeColor = vbBlack
   fldunit.ForeColor = vbBlack
   fldtglscan.ForeColor = vbBlack
   fldno.ForeColor = vbBlack
   End If
Else
    fldnmgudang.ForeColor = vbBlack
   fldnmgudang1.ForeColor = vbBlack
   fldX.ForeColor = vbBlack
   fldkdbarang.ForeColor = vbBlack
   fldnmbarang.ForeColor = vbBlack
   fldkd1.ForeColor = vbBlack
   fldunit.ForeColor = vbBlack
   fldtglscan.ForeColor = vbBlack
   fldno.ForeColor = vbBlack
End If
   
   

End Sub


