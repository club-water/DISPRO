Attribute VB_Name = "ModuleGradient"
'modul gradien
Dim rsW As ADODB.Recordset
Dim sqlw As String
Dim AlamatSkin As String

Enum GradMode
gmHorizontal = 0
gmVertical = 1
End Enum


Public Function GradientForm(frm As Form, Mode As GradMode)
Startcolor = vbBlue
Endcolor = vbBlue

Dim rs As Integer, Gs As Integer, Bs As Integer
Dim Re As Integer, Ge As Integer, Be As Integer
Dim Rk As Single, Gk As Single, Bk As Single
Dim R As Integer, G As Integer, b As Integer
Dim i As Integer, j As Single

On Error Resume Next
frm.AutoRedraw = True
frm.ScaleMode = vbPixels

rs = Startcolor And (Not &HFFFFFF00)
Gs = (Startcolor And (Not &HFFFF00FF)) \ &H100&
Bs = (Startcolor And (Not &HFF00FFFF)) \ &HFFFF&
Re = Endcolor And (Not &HFFFFFF00)
Ge = (Endcolor And (Not &HFFFF00FF)) \ &H100&
Be = (Endcolor And (Not &HFF00FFFF)) \ &HFFFF&

j = IIf(Mode = gmHorizontal, frm.ScaleWidth, frm.ScaleHeight)
Rk = (rs - Re) / j: Gk = (Gs - Ge) / j: Bk = (Bs - Be) / j

For i = 0 To j
R = rs - i * Rk: G = Gs - i * Gk: b = Bs - i * Bk
If Mode = gmHorizontal Then
frm.Line (i, 0)-(i - 1, frm.ScaleHeight), RGB(R, G, b), B
Else
frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(R, G, b), B
End If
Next
End Function







