# Zor durumlar içeren bir test PST dosyası üretir, sonra Outlook profilinden çıkarır.
# Kullanım: pwsh tests\test_pst_olustur.ps1 -Dizin <klasör>
# Çıktı: <Dizin>\kaynak\test_kaynak.pst ve <Dizin>\ekler\ (eklerin asılları, karşılaştırma için)
#
# DİKKAT: Folders.Items.Add(olMailItem) ile oluşturulan gönderilmemiş mail, Save() anında
# varsayılan hesabın TASLAKLAR klasörüne gider (IMAP ise sunucuya da eşitlenir). Bu yüzden
# öğeler PostItem (6) olarak oluşturulup MessageClass = IPM.Note yapılır: PostItem bulunduğu
# klasöre kaydedilir. Betik başında ve sonunda varsayılan Taslaklar sayılır; artarsa durur.
param([Parameter(Mandatory)] [string] $Dizin)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$kaynak = Join-Path $Dizin 'kaynak'
$ekler = Join-Path $Dizin 'ekler'
New-Item -ItemType Directory -Force $kaynak, $ekler | Out-Null
$pst = Join-Path $kaynak 'test_kaynak.pst'
if (Test-Path $pst) { Remove-Item $pst }

# Ek dosyaları (bilinen içerik)
$uzunAd = 'Fatura Şubat 2026 – müşteri ödeme bildirimi ve çok uzun bir dosya adı örneği (kopya) İĞÜŞÖÇ.pdf'
$pdf = Join-Path $ekler $uzunAd
$rnd = New-Object byte[] 200000; (New-Object Random 42).NextBytes($rnd); [IO.File]::WriteAllBytes($pdf, $rnd)
$txt = Join-Path $ekler 'not ğüşıöç.txt'
[IO.File]::WriteAllText($txt, "Türkçe metin: ğüşıöçĞÜŞİÖÇ`r`nikinci satır`r`n", [Text.UTF8Encoding]::new($false))
$png = Join-Path $ekler 'logo.png'
$bmp = New-Object System.Drawing.Bitmap 40, 20; $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::OrangeRed); $g.Dispose()
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

$wasRunning = [bool](Get-Process OUTLOOK -ErrorAction SilentlyContinue)
$app = New-Object -ComObject Outlook.Application
$ns = $app.GetNamespace('MAPI')
$draftsBefore = $ns.GetDefaultFolder(16).Items.Count
$ns.AddStoreEx($pst, 2)
$st = @($ns.Stores) | Where-Object { $_.FilePath -like '*\test_kaynak.pst' } | Select-Object -First 1
if (-not $st) { throw 'Test PST profilde bulunamadı' }
$root = $st.GetRootFolder()
try {
    $root.Name = 'Test Kaynak'
    $f1 = $root.Folders.Add('Test Klasörü')
    $deep = $f1.Folders.Add('Alt Klasör').Folders.Add('İç İçe: çok derin ve uzun adlı bir klasör örneği numara bir')
    $kisiler = $root.Folders.Add('Kişilerim', 10)

    function Yeni($klasor, [string]$konu) {
        $p = $klasor.Items.Add(6)   # olPostItem: bulunduğu klasöre kaydedilir
        $p.Subject = $konu
        $p
    }
    function MaileCevir($p) {
        $p.MessageClass = 'IPM.Note'
        $p.Save()
    }

    # 1) HTML + gömülü resim + uzun Türkçe adlı PDF + txt
    $m = Yeni $f1 'Şubat faturası – ödeme bildirimi 📎 (ekli)'
    $m.HTMLBody = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1254"></head><body><p>Merhaba, <b>ğüşıöç ĞÜŞİÖÇ</b></p><img src="cid:logo123@test"></body></html>'
    $a = $m.Attachments.Add($png)
    $a.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', 'logo123@test')
    [void]$m.Attachments.Add($pdf)
    [void]$m.Attachments.Add($txt)
    MaileCevir $m

    # 2) Düz metin, dosya adında yasak karakterler, ekli mail (gömülü)
    $ic = Yeni $f1 'İç mail: ekin içindeki mail'
    $ic.Body = "Bu mail başka bir mailin ekidir.`r`nSatır 2"
    [void]$ic.Attachments.Add($txt)
    MaileCevir $ic
    $m2 = Yeni $f1 'a/b\c:d*e?f"g<h>i|j yasak karakterler'
    $m2.BodyFormat = 1
    $m2.Body = "Düz metin gövde`r`n.nokta ile başlayan satır`r`n"
    [void]$m2.Attachments.Add($ic, 5)
    MaileCevir $m2

    # 3) RTF gövde
    $m3 = Yeni $deep 'RTF biçimli mail'
    $m3.BodyFormat = 3
    $m3.Body = 'RTF gövde metni ğüş'
    MaileCevir $m3

    # 4) Konusu boş, işaretli ve okunmamış
    $m4 = Yeni $deep ''
    $m4.Body = 'Konusu boş, işaretli, okunmamış'
    $m4.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x10900003', 2)
    MaileCevir $m4
    $m4.UnRead = $true
    $m4.Save()

    # 5) Kişi (EML'e girmez, tek PST'ye girer)
    $c = $kisiler.Items.Add(2); $c.FullName = 'Deneme Kişi'; $c.Email1Address = 'deneme@ornek.com'; $c.Save()

    "Oluşturuldu: " + $f1.Items.Count + " + " + $deep.Items.Count + " mail, " + $kisiler.Items.Count + " kişi"
}
finally {
    $ns.RemoveStore($root)
    $draftsAfter = $ns.GetDefaultFolder(16).Items.Count
    "Varsayılan Taslaklar: önce $draftsBefore, sonra $draftsAfter"
    if ($draftsAfter -ne $draftsBefore) { Write-Warning 'TASLAKLAR DEĞİŞTİ: test öğesi gerçek hesaba sızmış olabilir, elle kontrol edin!' }
    if (-not $wasRunning) { $app.Quit() }
}
"PST=$pst"
