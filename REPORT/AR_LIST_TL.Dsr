VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_LIST_TL 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   25109
   _ExtentY        =   15954
   SectionData     =   "AR_LIST_TL.dsx":0000
End
Attribute VB_Name = "AR_LIST_TL"
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
End Sub
