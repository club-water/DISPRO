Attribute VB_Name = "ModuleMSG"
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
 
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long


Public Const NV_CLOSEMSGBOX As Long = &H5000&

Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long




Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByValdwTime As Long)
    KillTimer hWnd, idEvent
    Select Case idEvent
    Case NV_CLOSEMSGBOX
        Dim hMessageBox, hMessageBox1, hMessageBox2 As Long
        hMessageBox = FindWindow("#32770", "Info !")
        hMessageBox1 = FindWindow("#32770", "Error !")
        hMessageBox2 = FindWindow("#32770", "Peringatan !")
        If hMessageBox Or hMessageBox1 Or hMessageBox2 Then
            Call SetForegroundWindow(hMessageBox)
            SendKeys "{enter}"
        End If
    End Select
End Sub







