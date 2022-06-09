VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_S_PENARIKAN 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   26141
   _ExtentY        =   16219
   SectionData     =   "AR_S_PENARIKAN.dsx":0000
End
Attribute VB_Name = "AR_S_PENARIKAN"
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

