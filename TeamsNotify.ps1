param(
    [string]$Duration,
    [string]$SystemCount
)

# ==========================================
# TEAMS WEBHOOK AYARI
# ==========================================

$webhookUrl = "Sizin  webhook url"

if ([string]::IsNullOrWhiteSpace($webhookUrl)) {
    Write-Host "Webhook URL boş. Mesaj gönderilmedi."
    exit
}

# ==========================================
# ADAPTIVE CARD OLUŞTURMA
# ==========================================

$card = @{
    type = "message"
    attachments = @(
        @{
            contentType = "application/vnd.microsoft.card.adaptive"
            content = @{
                type = "AdaptiveCard"
                version = "1.4"
                body = @(
                    @{
                        type = "TextBlock"
                        text = "SAP Ekran Kontrol Otomasyonu Tamamlandı"
                        weight = "Bolder"
                        size = "Large"
                        color = "Good"
                    },
                    @{
                        type = "FactSet"
                        facts = @(
                            @{
                                title = "Durum:"
                                value = "Basarılı"
                            },
                            @{
                                title = "Kontrol Edilen Sistem:"
                                value = $SystemCount
                            },
                            @{
                                title = "Toplam Süre:"
                                value = "$Duration saniye"
                            }
                        )
                    },
                    @{
                        type = "TextBlock"
                        text = "Kontrol Edilen Modüller"
                        weight = "Bolder"
                        size = "Medium"
                        spacing = "Medium"
                    },
                    @{
                        type = "TextBlock"
                        text = "- ST22 (Dump Analizi)`n- ST04 (DB Performansı)`n- SM19 (Güvenlik Logları)`n- SCC4 (Client Ayarları)`n- DB13 (Takvim Planları)`n- DB12 (Yedekleme Logları)`n- SOST (Mail Kuyrugu)`n- SM37 (İptal Olan Joblar)"
                        wrap = $true
                    },
                    @{
                        type = "TextBlock"
                        text = "Bu mesaj SAP GUI Scripting otomasyonu tarafından gönderildi."
                        spacing = "Medium"
                        isSubtle = $true
                        wrap = $true
                    }
                )
            }
        }
    )
}

try {
    $json = $card | ConvertTo-Json -Depth 10 -Compress

    Invoke-RestMethod -Uri $webhookUrl -Method Post -ContentType "application/json" -Body $json
}
catch {
    $err = $_ | Out-String
    $logPath = "$env:ProgramData\TeamsNotify\sendlog.txt"
    
    if (!(Test-Path (Split-Path $logPath))) {
        New-Item -ItemType Directory -Path (Split-Path $logPath) -Force | Out-Null
    }

    Add-Content -Path $logPath -Value ("[{0}] Gönderim hatası: {1}" -f (Get-Date), $err)
}
