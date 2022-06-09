VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_1A3 
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
   SectionData     =   "AR_1A3.dsx":0000
End
Attribute VB_Name = "AR_1A3"
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
On Error GoTo hell

Static i As Long

i = i + 1

fldno = i & "."

If fldUnit2.DataValue < 0 And fldBFS = 1 Then
Frame1.BackColor = vbRed

ElseIf fldUnit2 >= 0 And fldBFS = 1 Then
    If CCur(fldunit1) < CCur(fldUnit_BFS) * 0.25 Then
    Frame1.BackColor = vbYellow
    Else
    Frame1.BackColor = vbWhite
    End If
Else
Frame1.BackColor = vbWhite
End If



If fldBFS = 1 Then
fldBFS = "X"
Else
fldBFS = ""
End If


Exit Sub
hell:
Frame1.BackColor = vbGreen

End Sub


