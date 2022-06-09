Attribute VB_Name = "Module"
Public con As ADODB.Connection
Public koneksi As String
Public con1 As ADODB.Connection
Public koneksi1 As String
Public rsW As ADODB.Recordset
Dim sqlw As String
Public a As Integer
Public out As Integer
Public out1 As Integer
Public out2 As Integer
Public rstgl_OD As ADODB.Recordset
Public rsSave As ADODB.Recordset
Public alamat_save As String


Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, filename As String) As String
Dim sRet As String
sRet = String(255, Chr(0))
ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), filename))
End Function

Public Sub Koneksi_dbase()
Dim filename As String
Dim Ti_catalog, Td_source, Tprovider, Tu_ID, Tpass As String
Dim Tbawah, Tatas, Tbackup, TDbase, Tfolder, Toffice As String

filename = App.Path & "\Koneksi.ini"
Ti_catalog = ReadINI("Koneksi", "Initial Catalog", filename)
Td_source = ReadINI("Koneksi", "Data Source", filename)
Tprovider = ReadINI("Koneksi", "Provider", filename)
Tu_ID = ReadINI("Koneksi", "User ID", filename)
Tpass = ReadINI("Koneksi", "Password", filename)



koneksi = "Provider=" & Tprovider & ";Persist Security  Info=False;User ID=sa;Password=" & Tpass & ";Initial Catalog=" & Ti_catalog & ";Data Source=" & Td_source & ""

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
con.ConnectionString = koneksi
con.Open

End Sub


Private Sub Main()
Login.Show
End Sub


Public Sub con_mysql()
Dim filename As String
Dim Ti_catalog1, Td_source1, Tprovider1 As String

filename = App.Path & "\Koneksi.ini"
Ti_catalog1 = ReadINI("Koneksi", "Initial Catalog1", filename)
Td_source1 = ReadINI("Koneksi", "Data Source1", filename)
Tprovider1 = ReadINI("Koneksi", "Provider1", filename)



koneksi1 = "Provider=" & Tprovider1 & ";Persist Security  Info=False;Data Source=" & Td_source1 & ""


'koneksi1 = "Provider=" & Tprovider1 & ";Persist Security  Info=False;Initial Catalog=" & Ti_catalog1 & ";Data Source=" & Td_source1 & ""



Set con1 = New ADODB.Connection
con1.CursorLocation = adUseClient
con1.ConnectionString = koneksi1
con1.Open

End Sub


Public Sub nul(t As Variant)
If t = "" Then
t.BackColor = &HFFFF80
Else
t.BackColor = vbWhite
End If
End Sub


'Public Sub BC(X As Variant)
'sqlW = "select * from warna where idwarna=1"
'Set rsW = con.Execute(sqlW)
'X.BackColor = rsW!kdwarna
'End Sub
                                                                                                                                                       

'untuk membuka ms exel dan word
Public Sub EX_WORD(frm As Form)
Dim filename As String
Dim T_word As String

filename = App.Path & "\Koneksi.ini"
T_word = ReadINI("Koneksi", "word" & CStr(UTAMA.lblms_office), filename)

On Error GoTo hell
Call save_out
Shell "" & T_word & " " & alamat_save & "\outfile" & CStr(out) & ".rtf", vbMaximizedFocus

out = 0
Exit Sub
hell:
frm.cmdrtf_Click

End Sub

Public Sub EX_EXEL(frm As Form)
Dim filename As String
Dim T_Exel As String
filename = App.Path & "\Koneksi.ini"


T_Exel = ReadINI("Koneksi", "exel" & CStr(UTAMA.lblms_office), filename)

On Error GoTo hell
Call save_out
Shell "" & T_Exel & "" & alamat_save & "\outfile" & CStr(out1) & ".xls", vbMaximizedFocus

out1 = 0
Exit Sub
hell:
frm.cmdxls_Click
End Sub


Public Sub EX_PDF(frm As Form)
Dim filename As String
Dim T_Pdf As String
filename = App.Path & "\Koneksi.ini"
T_Pdf = ReadINI("Koneksi", "pdf" & CStr(UTAMA.lblms_office), filename)

On Error GoTo hell
Call save_out
Shell "" & T_Pdf & " " & alamat_save & "\outfile" & CStr(out2) & ".pdf", vbMaximizedFocus

out2 = 0
Exit Sub
hell:
frm.cmdPDF_Click
End Sub





'program fungsi terbilang2 didalam fungsi terbilang2

' byval tidak dikasih tidak apa2
Public Function Terbilang2(ByVal x As Long) As String
  Dim abil As Variant
  
  abil = Array("", "SATU", "DUA", "TIGA", "EMPAT", "LIMA", "ENAM", "TUJUH", "DELAPAN", "SEMBILAN", "SEPULUH", "SEBELAS")
' tanda \ backslash merupakan pembulatan dari hasil bagi misal 65 \ 10 = 6
  If x < 12 Then
    Terbilang2 = " " & abil(x)
  ElseIf x < 20 Then
    Terbilang2 = Terbilang2(x - 10) & " BELAS"
  ElseIf x < 100 Then
    Terbilang2 = Terbilang2(x \ 10) & " PULUH" & Terbilang2(x Mod 10)
  ElseIf x < 200 Then
    Terbilang2 = " SERATUS" & Terbilang2(x - 100)
  ElseIf x < 1000 Then
    Terbilang2 = Terbilang2(x \ 100) & " RATUS" & Terbilang2(x Mod 100)
  ElseIf x < 2000 Then
    Terbilang2 = " SERIBU" & Terbilang2(x - 1000)
  ElseIf x < 1000000 Then
    Terbilang2 = Terbilang2(x \ 1000) & " RIBU" & Terbilang2(x Mod 1000)
  ElseIf x < 1000000000 Then
    Terbilang2 = Terbilang2(x \ 1000000) & " JUTA" & Terbilang2(x Mod 1000000)
  End If
End Function


Public Sub Cek_tglOD()
sqltgl_OD = "select * from OD where kdOD='A'"
Set rstgl_OD = con.Execute(sqltgl_OD)

End Sub


Public Sub save_out()
On Error GoTo hell

Set rsSave = con.Execute("select * from user_m where kduser='" & UTAMA.lblkduser & "' ")

If rsSave.RecordCount <> 0 Then
    If rsSave!alamat_save <> "" Then
    alamat_save = rsSave!alamat_save
    Else
    alamat_save = App.Path
    End If
End If



Exit Sub
hell:
alamat_save = App.Path

End Sub


