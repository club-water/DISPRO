VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_Kwitansi 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "AR_Kwitansi.dsx":0000
End
Attribute VB_Name = "AR_Kwitansi"
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
On Error Resume Next
Dim filename As String
Dim TACC As String

Image2.Picture = LoadPicture(App.Path & "\gambar\TT.gif")
IMG_STEMPEL.Picture = LoadPicture(App.Path & "\gambar\STP.gif")

filename = App.Path & "\Koneksi.ini"
TACC = ReadINI("Koneksi", "ACC", filename)
lblTT = CStr(TACC)


flduang.DataValue = "# " & Terbilang2(flduang.DataValue) & " RUPIAH  #"

Select Case fldBLN.DataValue
Case 1
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN JANUARI" & " " & FLDTHN.DataValue

Case 2
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN FEBRUARI" & " " & FLDTHN.DataValue

Case 3
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN MARET" & " " & FLDTHN.DataValue

Case 4
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN APRIL" & " " & FLDTHN.DataValue

Case 5
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN MEI" & " " & FLDTHN.DataValue

Case 6
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN JUNI" & " " & FLDTHN.DataValue

Case 7
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN JULI" & " " & FLDTHN.DataValue

Case 8
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN AGUSTUS" & " " & FLDTHN.DataValue

Case 9
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN SEPTEMBER" & " " & FLDTHN.DataValue

Case 10
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN OKTOBER" & " " & FLDTHN.DataValue

Case 11
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN NOVEMBER" & " " & FLDTHN.DataValue

Case 12
fldket1 = "BIAYA PEMAKAIAN DISPENSER BULAN DESEMBER" & " " & FLDTHN.DataValue

End Select

fldket2.DataValue = "JUMLAH " & FLDUNIT.DataValue & " UNIT,HARGA PEMAKAIAN PER UNIT = Rp " & Format(fldharga, "#,###0") & ",-"




Select Case Format(fldtglposting, "M")
Case 1
fldtglposting = Format(fldtglposting, "dd ") & "Januari " & Format(fldtglposting, "yyyy ")

Case 2
fldtglposting = Format(fldtglposting, "dd ") & "Februari " & Format(fldtglposting, "yyyy ")

Case 3
fldtglposting = Format(fldtglposting, "dd ") & "Maret " & Format(fldtglposting, "yyyy ")

Case 4
fldtglposting = Format(fldtglposting, "dd ") & "April " & Format(fldtglposting, "yyyy ")

Case 5
fldtglposting = Format(fldtglposting, "dd ") & "Mei " & Format(fldtglposting, "yyyy ")

Case 6
fldtglposting = Format(fldtglposting, "dd ") & "Juni " & Format(fldtglposting, "yyyy ")

Case 7
fldtglposting = Format(fldtglposting, "dd ") & "Juli " & Format(fldtglposting, "yyyy ")

Case 8
fldtglposting = Format(fldtglposting, "dd ") & "Agustus " & Format(fldtglposting, "yyyy ")

Case 9
fldtglposting = Format(fldtglposting, "dd ") & "September " & Format(fldtglposting, "yyyy ")

Case 10
fldtglposting = Format(fldtglposting, "dd ") & "Oktober " & Format(fldtglposting, "yyyy ")

Case 11
fldtglposting = Format(fldtglposting, "dd ") & "November " & Format(fldtglposting, "yyyy ")

Case 12
fldtglposting = Format(fldtglposting, "dd ") & "Desember " & Format(fldtglposting, "yyyy ")

End Select

lblket1 = "* Jatuh Tempo Pembayaran Maksimal 14 Hari Setelah Kwitansi Diterima"
lblKET = "* Mohon Pada Saat Transfer, dicantumkan Kode : " & Left(fldnokwitansi, 6)
End Sub



