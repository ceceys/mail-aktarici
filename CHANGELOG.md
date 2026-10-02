# Sürüm notları

## 1.1.0 · 2026-10-02

- Yeni arayüz: tek ekran, kaydırma yok. Solda kaynaklar, sağda kayıt yeri, biçim ve günlük; altta her zaman görünen başlat çubuğu.
- Sunucu (IMAP / POP3) ayarları ayrı küçük pencerede.
- **Kaldığı yerden devam:** yarım kalan aktarım aynı yere tekrar başlatılınca sağlam yazılmış mailler atlanır.
- Kayıt diski çıkarılırsa ya da dolarsa program kapanmaz, sebebini yazıp durur.
- Beklenmeyen hatada sebep Masaüstüne `MailAktarici_HATA_tarih.txt` olarak yazılır; günlüğün kopyası her zaman `%LOCALAPPDATA%\MailAktarici` altında.
- Dosyalar önce geçici adla yazılıp sonra yeniden adlandırılır: kesintide yarım dosya kalmaz.
- Outlook artık normal şekilde açılıyor; açılışta bir soru sorarsa (güvenli mod, profil seçimi, dosya onarımı) program soruyu öne getirip cevabı bekliyor. Önceki sürümdeki 0x80080005 hatası bundan kaynaklanıyordu.
- Yer yetmeyecekse aktarım başlamadan soruluyor.
- 40 GB üstündeki veride her hesap ayrı PST'ye yazılıyor (PST sınırı 50 GB).
- "Teslim edilemedi" bildirimlerindeki `message/delivery-status` eki artık MIME kurallarına uygun yazılıyor.
- Raporda PST kaynaklarının dosya yolu gösteriliyor.
- by cecey imzası, LinkedIn ve GitHub bağlantıları.

## 1.0.0 · 2026-09-29

- İlk sürüm: Outlook hesapları, diskteki PST'ler ve IMAP/POP3 sunucularından EML, MSG ve tek PST çıktısı; Excel listesi ve doğrulama raporu.
