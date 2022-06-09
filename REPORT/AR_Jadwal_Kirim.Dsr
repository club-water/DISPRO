VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_Jadwal_Kirim 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   25479
   _ExtentY        =   15319
   SectionData     =   "AR_Jadwal_Kirim.dsx":0000
End
Attribute VB_Name = "AR_JADWAL_KIRIM"
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
'On Error Resume Next
Static i As Long
i = i + 1

fldno = i

flduraian = DateDiff("d", fldtglPK, Date)

Zoom = 150

If fldnmcustomer = "" Then
fldno = ""
fldalamat = ""
fldnmkategori = ""
fldketerangan = ""
fldQty_DISP = ""
FldQTY_lain = ""
fldQty_SHW = ""
flduraian = ""
fldtglPK = ""
End If


If Planning_kirim.Opt1.Value = True Then
flduraian.Visible = True
    If flduraian > 14 Then
    flduraian.ForeColor = vbRed
    Else
    flduraian.ForeColor = vbBlack
    End If

Else
flduraian.Visible = False
End If


End Sub

