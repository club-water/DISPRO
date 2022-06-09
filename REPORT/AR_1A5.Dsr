VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_1A5 
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
   SectionData     =   "AR_1A5.dsx":0000
End
Attribute VB_Name = "AR_1A5"
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

fldNO = i & "."


If fldUnit1 = 0 Then
fldharga1 = 0
Else
fldharga1 = FormatNumber(CCur(fldrupiah1) / CCur(fldUnit1), 0)
End If

If fldunit2 = 0 Then
fldharga2 = 0
Else
fldharga2 = FormatNumber(CCur(fldrupiah2) / CCur(fldunit2), 0)
End If

If fldunit3 = 0 Then
fldharga3 = 0
Else
fldharga3 = FormatNumber(CCur(fldrupiah3) / CCur(fldunit3), 0)
End If

If fldunit4 = 0 Then
fldharga4 = 0
Else
fldharga4 = FormatNumber(CCur(fldrupiah4) / CCur(fldunit4), 0)
End If




End Sub



