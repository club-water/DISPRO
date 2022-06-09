VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_KWITANSI_GAB 
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
   SectionData     =   "AR_KWITANSI_GAB.dsx":0000
End
Attribute VB_Name = "AR_KWITANSI_GAB"
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

If Kwitansi_GAB.txttahun1 = Kwitansi_GAB.txttahun2 Then
fldket1 = "Sewa Pemakaian Dispencer Bulan " & Kwitansi_GAB.CMBbln1.Text & " s/d " & Kwitansi_GAB.cmbbln2.Text & " " & Kwitansi_GAB.txttahun1
Else
fldket1 = "Sewa Pemakaian Dispencer Bulan " & Kwitansi_GAB.CMBbln1.Text & " " & Kwitansi_GAB.txttahun1 & " s/d " & Kwitansi_GAB.cmbbln2.Text & " " & Kwitansi_GAB.txttahun2
End If
fldket1 = UCase(fldket1)

lblket1 = "* Jatuh Tempo Pembayaran Maksimal 14 Hari Setelah Kwitansi Diterima"
lblKET = "* Mohon Pada Saat Transfer, dicantumkan Kode : " & Kwitansi_GAB.lblkdcustomer
End Sub


