"""Aktarım klasöründeki .eml dosyalarını Python'un bağımsız MIME ayrıştırıcısıyla doğrular.

Kullanım: py eml_dogrula.py <Mailler_... klasörü> [--ayrinti]
Her dosya için: ayrıştırma kusurları, Subject/From/Date varlığı, gövde çözülebiliyor mu,
ek adları ve boyutları; _liste.tsv'deki konu ile EML'deki konu aynı mı.
"""
import email
import sys
from email import policy
from pathlib import Path


def parcalar(msg):
    ekler, govde = [], []
    for p in msg.walk():
        if p.is_multipart():
            continue
        disp = p.get_content_disposition()
        ctype = p.get_content_type()
        if disp in ("attachment", "inline") and ctype not in ("text/plain", "text/html") or p.get_filename():
            veri = p.get_payload(decode=True) or b""
            if ctype == "message/rfc822":
                veri = p.get_payload()[0].as_bytes() if isinstance(p.get_payload(), list) else veri
            ekler.append((p.get_filename(), ctype, len(veri), p.get("Content-ID")))
        elif ctype.startswith("text/"):
            metin = p.get_content()
            govde.append((ctype, len(metin)))
    return govde, ekler


def main():
    kok = Path(sys.argv[1])
    ayrinti = "--ayrinti" in sys.argv
    toplam = hatali = 0
    for liste in sorted(kok.rglob("_liste.tsv")):
        hesap = liste.parent
        beklenen = {}
        for satir in liste.read_text(encoding="utf-8").splitlines()[1:]:
            a = satir.split("\t")
            if a[0]:
                beklenen[a[0]] = a
        for rel, a in beklenen.items():
            if not rel.lower().endswith(".eml"):
                continue
            yol = hesap / rel
            toplam += 1
            sorunlar = []
            if not yol.exists():
                print(f"YOK  {yol}")
                hatali += 1
                continue
            ham = yol.read_bytes()
            msg = email.message_from_bytes(ham, policy=policy.default)
            kusur = list(msg.defects)
            for p in msg.walk():
                kusur += list(p.defects)
            if kusur:
                sorunlar.append("kusur: " + ", ".join(type(k).__name__ for k in kusur))
            for h in ("From", "Date", "Subject", "Message-ID"):
                if msg[h] is None:
                    sorunlar.append(h + " yok")
            konu = str(msg["Subject"] or "")
            if konu.strip() != a[6].strip() and a[6]:
                sorunlar.append(f"konu farklı: eml={konu!r} liste={a[6]!r}")
            try:
                govde, ekler = parcalar(msg)
            except Exception as e:  # noqa: BLE001
                sorunlar.append(f"gövde çözülemedi: {e}")
                govde, ekler = [], []
            if b"\r\n" not in ham[:2000] or b"\n" in ham.replace(b"\r\n", b""):
                sorunlar.append("satır sonu CRLF değil")
            if sorunlar:
                hatali += 1
                print(f"SORUN {rel}: " + " | ".join(sorunlar))
            elif ayrinti:
                print(f"TAMAM {rel}\n      konu={konu!r} kimden={msg['From']!s} tarih={msg['Date']!s}\n"
                      f"      gövde={govde} ekler={ekler}")
    print(f"\n{toplam} EML denetlendi, {hatali} sorunlu.")
    return 1 if hatali else 0


if __name__ == "__main__":
    sys.exit(main())
