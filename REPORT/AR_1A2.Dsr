VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_1A2 
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
   SectionData     =   "AR_1A2.dsx":0000
End
Attribute VB_Name = "AR_1A2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub

Private Sub Detail_Format()
On Error GoTo hell
Static i As Currency

i = i + CCur(fldmasuk.DataValue) - CCur(fldkeluar.DataValue)


fldsaldo = Format(i, "#,###0")


If fldsrt = "0" Then
fldmasuk.Visible = False
fldkeluar.Visible = False
Else
fldmasuk.Visible = True
fldkeluar.Visible = True
End If

Exit Sub
hell:
Unload Me

End Sub

