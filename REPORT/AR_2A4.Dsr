VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_2A4 
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
   SectionData     =   "AR_2A4.dsx":0000
End
Attribute VB_Name = "AR_2A4"
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



If fldrupiah1 <> 0 Then
fldno.ForeColor = vbBlack
fldkdcustomer.ForeColor = vbBlack
fldnmcus.ForeColor = vbBlack
fldalamat.ForeColor = vbBlack
fldTT.ForeColor = vbBlack
fldkdpiutang.ForeColor = vbBlack
fldtglposting.ForeColor = vbBlack
fldbln.ForeColor = vbBlack
fldtahun.ForeColor = vbBlack
fldumur.ForeColor = vbBlack
fldrupiah.ForeColor = vbBlack
fldrupiah1.ForeColor = vbBlack
fldrupiah2.ForeColor = vbBlack
fldrupiah3.ForeColor = vbBlack
fldrupiah4.ForeColor = vbBlack

ElseIf fldrupiah2 <> 0 Then
    If fldTT = "" Then
    fldno.ForeColor = &H80FF&
    fldkdcustomer.ForeColor = &H80FF&
    fldnmcus.ForeColor = &H80FF&
    fldalamat.ForeColor = &H80FF&
    fldTT.ForeColor = &H80FF&
    fldkdpiutang.ForeColor = &H80FF&
    fldtglposting.ForeColor = &H80FF&
    fldbln.ForeColor = &H80FF&
    fldtahun.ForeColor = &H80FF&
    fldumur.ForeColor = &H80FF&
    fldrupiah.ForeColor = &H80FF&
    fldrupiah1.ForeColor = &H80FF&
    fldrupiah2.ForeColor = &H80FF&
    fldrupiah3.ForeColor = &H80FF&
    fldrupiah4.ForeColor = &H80FF&

    Else
    
    fldno.ForeColor = &H8000&
    fldkdcustomer.ForeColor = &H8000&
    fldnmcus.ForeColor = &H8000&
    fldalamat.ForeColor = &H8000&
    fldTT.ForeColor = &H8000&
    fldkdpiutang.ForeColor = &H8000&
    fldtglposting.ForeColor = &H8000&
    fldbln.ForeColor = &H8000&
    fldtahun.ForeColor = &H8000&
    fldumur.ForeColor = &H8000&
    fldrupiah.ForeColor = &H8000&
    fldrupiah1.ForeColor = &H8000&
    fldrupiah2.ForeColor = &H8000&
    fldrupiah3.ForeColor = &H8000&
    fldrupiah4.ForeColor = &H8000&
    End If
ElseIf fldrupiah3 <> 0 Then
fldno.ForeColor = &H80FF&
fldkdcustomer.ForeColor = &H80FF&
fldnmcus.ForeColor = &H80FF&
fldalamat.ForeColor = &H80FF&
fldTT.ForeColor = &H80FF&
fldkdpiutang.ForeColor = &H80FF&
fldtglposting.ForeColor = &H80FF&
fldbln.ForeColor = &H80FF&
fldtahun.ForeColor = &H80FF&
fldumur.ForeColor = &H80FF&
fldrupiah.ForeColor = &H80FF&
fldrupiah1.ForeColor = &H80FF&
fldrupiah2.ForeColor = &H80FF&
fldrupiah3.ForeColor = &H80FF&
fldrupiah4.ForeColor = &H80FF&
ElseIf fldrupiah4 <> 0 Then
fldno.ForeColor = vbRed
fldkdcustomer.ForeColor = vbRed
fldnmcus.ForeColor = vbRed
fldalamat.ForeColor = vbRed
fldTT.ForeColor = vbRed
fldkdpiutang.ForeColor = vbRed
fldtglposting.ForeColor = vbRed
fldbln.ForeColor = vbRed
fldtahun.ForeColor = vbRed
fldumur.ForeColor = vbRed
fldrupiah.ForeColor = vbRed
fldrupiah1.ForeColor = vbRed
fldrupiah2.ForeColor = vbRed
fldrupiah3.ForeColor = vbRed
fldrupiah4.ForeColor = vbRed

End If
End Sub



