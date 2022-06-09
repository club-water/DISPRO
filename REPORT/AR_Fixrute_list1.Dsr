VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_Fixrute_list1 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   8955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   24580
   _ExtentY        =   15796
   SectionData     =   "AR_Fixrute_list1.dsx":0000
End
Attribute VB_Name = "AR_Fixrute_list1"
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

Zoom = 150

fldin = Left(fldin, 5)
fldout = Left(fldout, 5)
End Sub
