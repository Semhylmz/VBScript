Option Explicit

Dim WScriptShell : Set WScriptShell = CreateObject("WScript.Shell")

' ==========================================
' AYARLAR: KONTROL EDİLECEK SİSTEMLER
' ==========================================
Dim systems
' Buraya kendi SAP Sistem ID'lerinizi (SID) yazın
systems = Array("SID","SID","SID")

' ==========================================
' YARDIMCI FONKSİYONLAR
' ==========================================
Function WaitForConnection(app)
    Dim t
    t = 0
    Do While app.Children.Count = 0
        WScript.Sleep 500
        t = t + 1
        If t > 40 Then
            MsgBox "Hata: SAP GUI bağlantısı açılmadı veya zaman aşımına uğradı.", vbCritical
            WScript.Quit
        End If
    Loop
    Set WaitForConnection = app.Children(app.Children.Count - 1)
End Function

Function SafeFindById(sess, id)
    On Error Resume Next
    Dim obj : Set obj = sess.FindById(id, False) ' False parametresi hata fırlatmayı engeller
    On Error GoTo 0
    Set SafeFindById = obj
End Function

Sub SendTCode(sess, tcode)
    Dim ok : Set ok = SafeFindById(sess, "wnd[0]/tbar[0]/okcd")
    If Not ok Is Nothing Then
        ok.text = tcode
        SafeFindById(sess, "wnd[0]").sendVKey 0 ' ENTER
        WScript.Sleep 1500
    End If
End Sub

' ==========================================
' ANA İŞLEM DÖNGÜSÜ
' ==========================================
Dim startTime, endTime, totalSeconds
startTime = Timer

Dim i
For i = LBound(systems) To UBound(systems)
    Dim sysName : sysName = systems(i)

    ' 1. SAP Logon Penceresini Bul ve Odaklan
    WScriptShell.AppActivate "SAP Logon"
    WScript.Sleep 600
    WScriptShell.SendKeys "{ESC}"  ' Açık popup varsa kapat
    WScript.Sleep 150
    WScriptShell.SendKeys "{HOME}" ' Listede en başa dön
    WScript.Sleep 150
    
    ' Sistem adını yazarak bul (Listede sysName'i arar)
    WScriptShell.SendKeys sysName 
    WScript.Sleep 250
    WScriptShell.SendKeys "{ENTER}" ' Giriş Yap
    WScript.Sleep 3000 ' Bağlantı açılış süresi (Yavaş ağlar için artırılabilir)

    ' 2. SAP GUI Scripting Engine Bağlantısı
    Dim SapGuiAuto, application, connection, session
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        MsgBox "SAP GUI açık değil. Lütfen önce SAP Logon'u başlatın.", vbCritical
        WScript.Quit
    End If
    On Error GoTo 0

    Set application = SapGuiAuto.GetScriptingEngine
    Set connection = WaitForConnection(application)

    ' Session nesnesinin hazır olmasını bekle
    Do While connection.Children.Count = 0
        WScript.Sleep 400
    Loop
    Set session = connection.Children(0)

    ' Pencereyi Tam Ekran Yap
    SafeFindById(session, "wnd[0]").maximize
    WScript.Sleep 500

    ' ==========================================
    ' RUTİN KONTROLLER
    ' ==========================================

    ' --- 1. ST22 (ABAP Dumps) ---
    SendTCode session, "ST22"
    WScript.Sleep 1000
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- 2. ST04 (DB Performance) ---
    SendTCode session, "ST04"
    WScript.Sleep 2000
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- 3. SM19 (Security Audit) ---
    SendTCode session, "SM19"
    Dim tabAdmin : Set tabAdmin = SafeFindById(session, "wnd[0]/usr/tabsTABSTRIP2/tabpADMIN")
    If Not tabAdmin Is Nothing Then 
        tabAdmin.select
        WScript.Sleep 1700
    End If
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- 4. SCC4 (Client Settings) ---
    SendTCode session, "SCC4"
    WScript.Sleep 1500
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- 5. DB13 (DB Planning Calendar) ---
    SendTCode session, "DB13"
    WScript.Sleep 2000
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- 6. DB12 (Backup Logs - Scroll Örneği) ---
    SendTCode session, "DB12"
    WScript.Sleep 2000
    Dim alv : Set alv = SafeFindById(session, "wnd[0]/usr/cntlBACKUPCAT_ALV_CONTAINER/shellcont/shell")
    If Not alv Is Nothing Then
        ' Listeyi aşağı kaydırarak logları kontrol etme simülasyonu
        alv.firstVisibleRow = 2
        WScript.Sleep 200
        alv.firstVisibleRow = 10
        WScript.Sleep 200
        alv.firstVisibleRow = 20
    End If
    WScript.Sleep 1000
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- 7. SOST (Mail Queue - Hata Filtresi) ---
    SendTCode session, "SOST"
    ' Hata sekmesine geç
    Dim tabErr : Set tabErr = SafeFindById(session, "wnd[0]/usr/subSUB:SAPLSBCS_OUT:1100/subTOPSUB:SAPLSBCS_OUT:1110/tabsTAB1/tabpTAB1_FC2")
    If Not tabErr Is Nothing Then tabErr.select
    
    ' Hata checkbox'ını işaretle ve Yenile
    Dim chkErr : Set chkErr = SafeFindById(session, "wnd[0]/usr/subSUB:SAPLSBCS_OUT:1100/subTOPSUB:SAPLSBCS_OUT:1110/tabsTAB1/tabpTAB1_FC2/ssubTAB1_SCA:SAPLSBCS_OUT:0001/chkG_ERROR")
    If Not chkErr Is Nothing Then chkErr.selected = True
    
    Dim btnRefresh : Set btnRefresh = SafeFindById(session, "wnd[0]/usr/subSUB:SAPLSBCS_OUT:1100/subTOPSUB:SAPLSBCS_OUT:1110/tabsTAB1/tabpTAB1_FC2/ssubTAB1_SCA:SAPLSBCS_OUT:0001/btnREFRICO1")
    If Not btnRefresh Is Nothing Then btnRefresh.press
    
    WScript.Sleep 2000
    SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press ' Geri

    ' --- 8. SM37 (Job Overview - Hata Filtresi) ---
    SendTCode session, "SM37"
    ' Sadece 'Canceled' olanları seç
    Dim sSched, sReady, sFin, sUser
    Set sSched = SafeFindById(session, "wnd[0]/usr/chkBTCH2170-SCHEDUL") : If Not sSched Is Nothing Then sSched.selected = False
    Set sReady = SafeFindById(session, "wnd[0]/usr/chkBTCH2170-READY")  : If Not sReady Is Nothing Then sReady.selected = False
    Set sFin   = SafeFindById(session, "wnd[0]/usr/chkBTCH2170-FINISHED"): If Not sFin Is Nothing Then sFin.selected = False
    Set sUser  = SafeFindById(session, "wnd[0]/usr/txtBTCH2170-USERNAME"): If Not sUser Is Nothing Then sUser.text = "*"
    
    SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press ' Execute (F8)
    WScript.Sleep 2000
    SafeFindById(session, "wnd[0]").sendVKey 3 ' Geri

    ' --- ÇIKIŞ ---
    SendTCode session, "/nex" ' Direkt çıkış komutu
    WScript.Sleep 1000
    connection.CloseConnection
Next

endTime = Timer
totalSeconds = Round(endTime - startTime, 2)

' ==========================================
' SONUÇ BİLDİRİMİ (POWERSHELL TETİKLEME)
' ==========================================
' PowerShell dosyanızın yolunu aşağıdan kontrol edin: C:\Scripts\TeamsNotify.ps1
WScriptShell.Run "powershell.exe -ExecutionPolicy Bypass -File ""C:\Users\semih.yilmaz\Desktop\TeamsNotify.ps1"" -Duration """ & totalSeconds & """ -SystemCount """ & (UBound(systems)+1) & """", 0, True

MsgBox "Tüm sistem kontrolleri tamamlandı.", vbInformation, "İşlem Bitti"
