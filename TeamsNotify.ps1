param(
    [string]$Duration,
    [string]$SystemCount
)

# ==========================================
# AYARLAR: TEAMS WEBHOOK
# ==========================================
# Teams Kanalı -> Connectors -> Incoming Webhook yolundan aldığınız URL'i buraya yapıştırın.
$webhookUrl = "https://outlook.office.com/webhook/SİZİN_WEBHOOK_URL_ADRESİNİZ..."

# ==========================================
# MESAJ İÇERİĞİ OLUŞTURMA
# ==========================================
$channelMessageText = @"
🚀 **SAP Rutin Kontrol Otomasyonu Tamamlandı**

✅ **Durum:** Başarılı
📊 **Kontrol Edilen Sistem Sayısı:** $SystemCount
⏱️ **Toplam Süre:** $Duration saniye

**Kontrol Edilen Modüller:**
- ST22 (Dump Analizi)
- ST04 (DB Performansı)
- SM19 (Güvenlik Logları)
- SCC4 (Client Ayarları)
- DB13 (Takvim Planları)
- DB12 (Yedekleme Logları)
- SOST (Mail Kuyruğu)
- SM37 (İptal Olan Joblar)

_Bu mesaj SAP GUI Scripting otomasyonu tarafından otomatik gönderilmiştir._
"@

$message = @{
    text = $channelMessageText
}

# ==========================================
# GÖNDERİM VE LOGLAMA
# ==========================================
try {
    # JSON Formatına Çevir
    $json = $message | ConvertTo-Json -Depth 3
    
    # Teams'e POST isteği at
    Invoke-RestMethod -Uri $webhookUrl -Method Post -ContentType 'application/json' -Body $json
}
catch {
    # Hata durumunda sessizce log tut (Kullanıcıyı rahatsız etme)
    $err = $_ | Out-String
    $logPath = "$env:ProgramData\TeamsNotify\sendlog.txt"
    
    # Log klasörü yoksa oluştur
    if (!(Test-Path (Split-Path $logPath))) {
        New-Item -ItemType Directory -Path (Split-Path $logPath) -Force | Out-Null
    }
    
    # Hatayı dosyaya yaz
    Add-Content -Path $logPath -Value ("[{0}] Gönderim hatası: {1}" -f (Get-Date), $err)
}
