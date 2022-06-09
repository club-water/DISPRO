Attribute VB_Name = "ModuleMS"
Public Const WARNA_BIRU_TUA    As Long = &H800000
Public Const WARNA_BIRU        As Long = &HED9564
Public Const WARNA_ABU_ABU     As Long = &HDEC4B0
Public Const WARNA_KUNING       As Long = vbYellow

Public Function setBarMenu(ByVal explorerBar As Object, ByVal menuName As String, _
                            ByVal menuCaption As String, ByVal iconIndex As Long) As Object
                           
    Dim cBar As Object
    
    On Error Resume Next
    Set cBar = explorerBar.Bars.Add(, menuName, menuCaption)
    cBar.IsSpecial = True
    cBar.iconIndex = iconIndex
    cBar.TitleForeColor = WARNA_BIRU_TUA
    cBar.TitleForeColorOver = WARNA_BIRU_TUA
    cBar.TitleBackColorLight = WARNA_BIRU
    cBar.TitleBackColorDark = RGB(234, 241, 253)
    cBar.BackColor = WARNA_ABU_ABU
    
        
    Set setBarMenu = cBar
End Function

Public Sub setItemMenu(ByVal cBar As Object, ByVal menuName As String, ByVal menuCaption As String, ByVal iconIndex As Long)
    Dim cItem   As Object
    
    On Error Resume Next
    Set cItem = cBar.Items.Add(, menuName, menuCaption)
    With cItem
        .iconIndex = iconIndex
        .TextColor = WARNA_BIRU_TUA
        .TextColorOver = WARNA_KUNING
    End With
End Sub

Public Sub addMenu(ByVal explorerBar As Object) ', ByVal barIcons As Object, ByVal itmIcons As Object)
    Dim cBar        As Object
        
    On Error Resume Next
    With explorerBar
        .UseExplorerStyle = False
        
        .Redraw = False

        .BackColorStart = WARNA_BIRU
        .BackColorEnd = WARNA_BIRU

        .ImageList = UTAMA.ImageList1
        .BarTitleImageList = UTAMA.vbalImageList1
        
    
        Set cBar = setBarMenu(explorerBar, "MMaster", "Master", 0)
        Call setItemMenu(cBar, "MM1", "Barang", 15)
        Call setItemMenu(cBar, "MM2", "Customer", 15)
        Call setItemMenu(cBar, "MM3", "Gudang", 15)
        
        Set cBar = setBarMenu(explorerBar, "MTA", "Keluar Barang", 0)
        Call setItemMenu(cBar, "MTA1", "Pinjaman", 15)
        Call setItemMenu(cBar, "MTA2", "Sewa", 15)
        Call setItemMenu(cBar, "MTA3", "Free", 15)
        
        Set cBar = setBarMenu(explorerBar, "MTB", "Terima Barang", 0)
        Call setItemMenu(cBar, "MTB1", "Pembelian", 15)
        Call setItemMenu(cBar, "MTB2", "Retur Pinjaman", 15)
        Call setItemMenu(cBar, "MTB3", "Retur Sewa", 15)
        
        Set cBar = setBarMenu(explorerBar, "MTC", "Piutang Sewa", 0)
        Call setItemMenu(cBar, "MTC1", "Cetak Kwitansi", 15)
        Call setItemMenu(cBar, "MTC2", "Pembayaran Piutang Sewa", 15)
        Call setItemMenu(cBar, "MTC3", "Laporan Harian Penagihan", 15)
        
        Set cBar = setBarMenu(explorerBar, "MS", "System", 0)
        Call setItemMenu(cBar, "MS1", "Daily Closing", 15)
        
        
        .Redraw = True
    End With
End Sub

