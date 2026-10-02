using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace MailAktarici
{
    // Günlük, ilerleme ve durdurma isteği. Arayüz ve komut satırı aynı nesneyi kullanır.
    // Günlük iki yere yazılır: aktarım klasörüne (gunluk.txt) ve bilgisayarın kendi diskine
    // (%LOCALAPPDATA%\MailAktarici). Hedef disk koparsa ikinci kopya kalır. Yazma hatası asla
    // dışarı sızmaz: günlüğe yazılamıyor diye program kapanmaz.
    sealed class Reporter
    {
        readonly object lk = new object();
        StreamWriter file, local;
        readonly List<string> early = new List<string>();
        readonly Queue<string> recent = new Queue<string>();
        DateTime lastHardFlush = DateTime.MinValue;
        bool fileBroken;
        long done, total;
        int errors, warnings;
        volatile string status = "";
        public volatile bool CancelRequested;
        public Action<string> Sink;
        public string LocalLogPath { get; private set; }

        public long Done { get { return Interlocked.Read(ref done); } }
        public long Total { get { return Interlocked.Read(ref total); } }
        public int Errors { get { return errors; } }
        public int Warnings { get { return warnings; } }
        public string Status { get { return status; } set { status = value ?? ""; } }

        public void AddTotal(long n) { Interlocked.Add(ref total, n); }
        public void Step() { Interlocked.Increment(ref done); }

        public void ResetProgress()
        {
            Interlocked.Exchange(ref done, 0);
            Interlocked.Exchange(ref total, 0);
            Interlocked.Exchange(ref errors, 0);
            Interlocked.Exchange(ref warnings, 0);
        }

        public static string LocalDir()
        {
            string d = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailAktarici");
            try { Directory.CreateDirectory(d); }
            catch { d = Path.GetTempPath(); }
            return d;
        }

        // Yerel kopya: her iş başında yeni dosya
        public void OpenLocal()
        {
            lock (lk)
            {
                CloseLocal();
                try
                {
                    LocalLogPath = Path.Combine(LocalDir(), "gunluk_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".txt");
                    local = new StreamWriter(LocalLogPath, true, new UTF8Encoding(true));
                }
                catch { local = null; }
            }
        }

        void CloseLocal()
        {
            if (local != null)
            {
                try { local.Dispose(); }
                catch { }
                local = null;
            }
        }

        public void OpenFile(string path)
        {
            lock (lk)
            {
                try
                {
                    file = new StreamWriter(path, true, new UTF8Encoding(true));
                    fileBroken = false;
                    foreach (string l in early) file.WriteLine(l);
                    file.Flush();
                }
                catch { file = null; }
                early.Clear();
            }
        }

        public void CloseFile()
        {
            lock (lk)
            {
                if (file != null)
                {
                    try { file.Dispose(); }
                    catch { }
                    file = null;
                }
                CloseLocal();
            }
        }

        public void Info(string m) { Write("", m); }
        public void Warn(string m) { Interlocked.Increment(ref warnings); Write("UYARI: ", m); }
        public void Error(string m) { Interlocked.Increment(ref errors); Write("HATA: ", m); }

        void Write(string prefix, string m)
        {
            string line = DateTime.Now.ToString("HH:mm:ss") + "  " + prefix + m;
            lock (lk)
            {
                recent.Enqueue(line);
                while (recent.Count > 400) recent.Dequeue();
                if (file == null && !fileBroken)
                {
                    early.Add(line);
                    if (early.Count > 5000) early.RemoveAt(0);
                }
                bool hard = (DateTime.UtcNow - lastHardFlush).TotalSeconds >= 2 || prefix.Length > 0;
                if (file != null)
                {
                    try
                    {
                        file.WriteLine(line);
                        file.Flush();
                        if (hard) ((FileStream)file.BaseStream).Flush(true);   // diske kadar
                    }
                    catch
                    {
                        // Hedef disk koptu/doldu: günlük yerel kopyada devam eder.
                        fileBroken = true;
                        try { file.Dispose(); }
                        catch { }
                        file = null;
                    }
                }
                if (local != null)
                {
                    try
                    {
                        local.WriteLine(line);
                        local.Flush();
                        if (hard) ((FileStream)local.BaseStream).Flush(true);
                    }
                    catch { CloseLocal(); }
                }
                if (hard) lastHardFlush = DateTime.UtcNow;
            }
            var s = Sink;
            if (s != null)
            {
                try { s(line); }
                catch { }
            }
        }

        public string[] RecentLines()
        {
            lock (lk) return recent.ToArray();
        }

        public void Check()
        {
            if (CancelRequested) throw new OperationCanceledException();
        }
    }

    // Kayıt yeri yazılamaz hale geldi (disk çıkarıldı, bağlantı koptu, disk doldu).
    sealed class TargetLostException : Exception
    {
        public TargetLostException(string m, Exception inner) : base(m, inner) { }
    }

    // Beklenmeyen çökme: sebebi bilgisayarın kendi Masaüstüne ve yerel klasöre yazılır.
    static class CrashReport
    {
        public static Reporter Current;

        public static string Write(Exception ex, string where)
        {
            var sb = new StringBuilder();
            sb.AppendLine("MAIL AKTARICI HATA RAPORU");
            sb.AppendLine("Zaman     : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("Program   : " + Program.Version + "  (" + (Environment.Is64BitProcess ? "64" : "32") + " bit)");
            sb.AppendLine("Bilgisayar: " + Environment.MachineName + "  kullanıcı: " + Environment.UserName);
            sb.AppendLine("Windows   : " + Environment.OSVersion + "  64 bit: " + Environment.Is64BitOperatingSystem);
            sb.AppendLine("Yer       : " + where);
            sb.AppendLine();
            sb.AppendLine("HATA:");
            sb.AppendLine(ex == null ? "(ayrıntı yok)" : ex.ToString());
            var r = Current;
            if (r != null)
            {
                sb.AppendLine();
                sb.AppendLine("SON GÜNLÜK SATIRLARI:");
                foreach (string l in r.RecentLines()) sb.AppendLine(l);
            }
            string name = "MailAktarici_HATA_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".txt";
            string written = null;
            foreach (string dir in new[] { Util.DesktopPath(), Reporter.LocalDir() })
            {
                try
                {
                    string p = Path.Combine(dir, name);
                    File.WriteAllText(p, sb.ToString(), new UTF8Encoding(true));
                    if (written == null) written = p;
                }
                catch { }
            }
            return written;
        }
    }
}
