VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_2A6 
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
   SectionData     =   "AR_2A6.dsx":0000
End
Attribute VB_Name = "AR_2A6"
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

If CCur(fldumur) <= 30 Then
fldrupiah1 = fldrupiah
fldrupiah2 = 0
fldrupiah3 = 0
fldrupiah4 = 0
ElseIf CCur(fldumur) >= 31 And CCur(fldumur) <= 60 Then
fldrupiah1 = 0
fldrupiah2 = fldrupiah
fldrupiah3 = 0
fldrupiah4 = 0
ElseIf CCur(fldumur) >= 61 And CCur(fldumur) <= 90 Then
fldrupiah1 = 0
fldrupiah2 = 0
fldrupiah3 = fldrupiah
fldrupiah4 = 0
ElseIf CCur(fldumur) >= 91 Then
fldrupiah1 = 0
fldrupiah2 = 0
fldrupiah3 = 0
fldrupiah4 = fldrupiah
End If


If fldtglbyr1 = "01/01/1900" Then
fldtglbyr1 = ""
End If

If fldtglbyr2 = "01/01/1900" Then
fldtglbyr2 = ""
End If

If fldtglbyr3 = "01/01/1900" Then
fldtglbyr3 = ""
End If



End Sub




