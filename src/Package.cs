using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace MailAktarici
{
    sealed class ManifestEntry
    {
        public string File, Folder, Role, Flags, MessageId, Subject, From, Status;
        public DateTime? DateUtc;
        public long Size;
        public int Attachments = -1;
        public int FilesWritten;   // diskte duran .eml/.msg sayısı (doğrulama için)
    }

    sealed class FolderStat
    {
        public string Logical, Role;
        public long Source, Exported, Failed;
    }

    // Aktarım klasörü: hesap alt klasörleri, Excel listesi, rapor, günlük.
    // Yarım kalan bir klasörle yeniden açılırsa (Resume) diskte sağlam duran mailler yeniden yazılmaz.
    sealed class Package
    {
        public readonly string Root;
        public readonly Reporter R;
        public readonly bool Resume;
        readonly CsvWriter csv;
        readonly List<AccountSink> accounts = new List<AccountSink>();
        readonly HashSet<string> accountNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly DateTime started = DateTime.Now;
        public readonly List<string> Notes = new List<string>();
        public long Reused;   // devamda yeniden yazılmadan kullanılan dosya sayısı

        Package(string root, Reporter r, bool resume)
        {
            Root = root;
            R = r;
            Resume = resume;
            r.OpenFile(Path.Combine(root, "gunluk.txt"));
            csv = new CsvWriter(Path.Combine(root, "mail_listesi.csv"),
                "Hesap", "Klasör", "Tarih", "Gönderen", "Konu", "Ek", "Boyut (KB)", "Okundu", "Durum", "Dosya");
        }

        public static Package Create(string parentDir, Reporter r)
        {
            string baseName = "Mailler_" + Util.SafeName(Environment.MachineName, 30) + "_" + DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            string root = Path.Combine(parentDir, baseName);
            for (int n = 2; Directory.Exists(root); n++) root = Path.Combine(parentDir, baseName + "_" + n);
            Directory.CreateDirectory(root);
            var p = new Package(root, r, false);
            r.Info("Aktarım klasörü: " + root);
            return p;
        }

        public static Package Continue(string existingRoot, Reporter r)
        {
            try { File.Delete(Path.Combine(existingRoot, "ozet_rapor.txt")); }
            catch { }
            var p = new Package(existingRoot, r, true);
            r.Info("YARIM KALAN AKTARIMA DEVAM: " + existingRoot + " (diskte sağlam duran mailler yeniden yazılmayacak)");
            return p;
        }

        // Bu bilgisayar için bitmemiş aktarım klasörleri (en yenisi başta)
        public static List<string> FindUnfinished(string parentDir)
        {
            var list = new List<string>();
            try
            {
                string prefix = "Mailler_" + Util.SafeName(Environment.MachineName, 30) + "_";
                foreach (var d in new DirectoryInfo(parentDir).GetDirectories(prefix + "*").OrderByDescending(x => x.CreationTime))
                {
                    string rep = Path.Combine(d.FullName, "ozet_rapor.txt");
                    if (!File.Exists(rep)) { list.Add(d.FullName); continue; }
                    string text = File.ReadAllText(rep);
                    if (text.Contains("DURDURDU") || text.Contains("KOPTU") || text.Contains("DOLDU")) list.Add(d.FullName);
                }
            }
            catch { }
            return list;
        }

        public AccountSink OpenAccount(string name, string source)
        {
            string safe = Util.SafeName(name, 50);
            string unique = safe;
            for (int n = 2; !accountNames.Add(unique); n++) unique = safe + " (" + n + ")";
            var a = new AccountSink(this, name, source, Path.Combine(Root, unique));
            accounts.Add(a);
            return a;
        }

        int csvRows;

        internal void CsvRow(AccountSink a, ManifestEntry e)
        {
            csv.Row(a.Name, e.Folder,
                e.DateUtc.HasValue ? e.DateUtc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm") : "",
                e.From, e.Subject,
                e.Attachments >= 0 ? e.Attachments.ToString() : "",
                e.Size > 0 ? Math.Ceiling(e.Size / 1024.0).ToString("0") : "",
                (e.Flags ?? "").Contains("\\Seen") ? "Evet" : "Hayır",
                e.Status,
                Path.Combine(Path.GetFileName(a.Dir), e.File ?? ""));
            if (++csvRows % 25 == 0) csv.FlushHard();
        }

        // Yazma hatası kayıt yerinin gitmesinden mi (disk çıkarıldı/doldu), yoksa tek bir mailden mi?
        // Kayıt yeri gittiyse aktarım anlamlı bir mesajla durur; tek mailse o mail hatalı sayılıp devam edilir.
        public void CheckTarget(Exception ex)
        {
            if (!(ex is IOException) && !(ex is UnauthorizedAccessException) && !(ex is System.Runtime.InteropServices.COMException)) return;
            int hr = ex.HResult;
            if (hr == unchecked((int)0x80070070) || hr == unchecked((int)0x80070027))
                throw new TargetLostException("Kayıt diski DOLDU. Aktarım durdu; yer açıp aynı yere tekrar başlatırsanız kaldığı yerden devam eder.", ex);
            bool gone;
            try { gone = !Directory.Exists(Root); }
            catch { gone = true; }
            if (gone)
                throw new TargetLostException("Kayıt yerine ulaşılamıyor: disk çıkarılmış ya da bağlantısı kopmuş (" + Root + "). " +
                    "Diski tekrar takıp aynı yere başlatırsanız kaldığı yerden devam eder.", ex);
        }

        // Yarım dosya kalmasın: önce geçici ada yazılır, sonra yeniden adlandırılır.
        public static void WriteAtomic(string path, byte[] data)
        {
            string tmp = Path.Combine(Path.GetDirectoryName(path), "~" + Guid.NewGuid().ToString("N").Substring(0, 12) + ".tmp");
            File.WriteAllBytes(tmp, data);
            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }

        // Devam ederken diskte duran dosya sağlam mı? (yarım yazılmış dosya yeniden yazılır)
        public static bool LooksComplete(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                if (!fi.Exists || fi.Length == 0) return false;
                if (path.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                    return fi.Length >= 1024 && fi.Length % 512 == 0;   // Outlook .msg bileşik dosyası 512 baytlık bloklardan oluşur
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var head = new byte[(int)Math.Min(fs.Length, 64 * 1024)];
                    fs.Read(head, 0, head.Length);
                    string h = Encoding.ASCII.GetString(head);
                    int b = h.IndexOf("boundary=\"", StringComparison.OrdinalIgnoreCase);
                    int hdrEnd = h.IndexOf("\r\n\r\n", StringComparison.Ordinal);
                    if (b < 0 || (hdrEnd >= 0 && b > hdrEnd)) return true;   // tek parçalı: sonu doğrulanamaz, kabul
                    int q = h.IndexOf('"', b + 10);
                    if (q < 0) return false;
                    string closing = "--" + h.Substring(b + 10, q - b - 10) + "--";
                    int tail = (int)Math.Min(fs.Length, 512);
                    fs.Seek(-tail, SeekOrigin.End);
                    var end = new byte[tail];
                    fs.Read(end, 0, tail);
                    return Encoding.ASCII.GetString(end).TrimEnd('\r', '\n', ' ').EndsWith(closing, StringComparison.Ordinal);
                }
            }
            catch { return false; }
        }

        public void FlushAll()
        {
            foreach (var a in accounts) a.FlushHard();
            try { csv.FlushHard(); }
            catch { }
        }

        // Rapor yazılır, dosyalar kapatılır. Dönüş: raporun yolu (yazılamadıysa null).
        public string Finish(string state)
        {
            foreach (var a in accounts) a.Dispose();
            try { csv.Dispose(); }
            catch { }

            var sb = new StringBuilder();
            sb.AppendLine("MAİL AKTARIM RAPORU");
            sb.AppendLine("===================");
            sb.AppendLine("Bilgisayar : " + Environment.MachineName + "  (Windows kullanıcısı: " + Environment.UserName + ")");
            sb.AppendLine("Başlangıç  : " + started.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("Bitiş      : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("Klasör     : " + Root);
            sb.AppendLine("Durum      : " + state);
            sb.AppendLine("Program    : Mail Aktarıcı " + Program.Version + " · by " + Program.Author + " · " + Program.LinkedInUrl);
            if (Resume) sb.AppendLine("Devam      : önceki yarım aktarımdan " + Reused.ToString("N0") + " dosya yeniden yazılmadan kullanıldı");
            sb.AppendLine();

            long srcAll = 0, expAll = 0, failAll = 0;
            bool allMatch = true;
            int idx = 0;
            foreach (var a in accounts)
            {
                idx++;
                sb.AppendLine("[" + idx + "] " + a.Name + "   (" + a.Source + ")");
                sb.AppendLine("    " + Pad("Klasör", 44) + LPad("Kaynakta", 10) + LPad("Aktarılan", 11) + LPad("Hata", 7));
                long s = 0, e = 0, f = 0;
                foreach (var fs in a.Folders)
                {
                    sb.AppendLine("    " + Pad(fs.Logical, 44) + LPad(fs.Source.ToString("N0"), 10) + LPad(fs.Exported.ToString("N0"), 11) + LPad(fs.Failed.ToString("N0"), 7)
                        + (fs.Exported + fs.Failed != fs.Source ? "   <- sayı tutmuyor" : ""));
                    s += fs.Source; e += fs.Exported; f += fs.Failed;
                }
                sb.AppendLine("    " + Pad("TOPLAM", 44) + LPad(s.ToString("N0"), 10) + LPad(e.ToString("N0"), 11) + LPad(f.ToString("N0"), 7));
                if (a.Oldest.HasValue)
                    sb.AppendLine("    Tarih aralığı: " + a.Oldest.Value.ToLocalTime().ToString("yyyy-MM-dd") + " ile " + a.Newest.Value.ToLocalTime().ToString("yyyy-MM-dd") + " arası");
                if (a.HeaderOnly > 0)
                    sb.AppendLine("    DİKKAT: " + a.HeaderOnly + " mailin yalnız başlığı bilgisayarda, gövdesi sunucuda kalmış (listede 'SADECE_BASLIK').");
                long onDisk = a.CountFilesOnDisk(".eml") + a.CountFilesOnDisk(".msg");
                long listed = a.ListedFiles;
                bool match = onDisk == listed && e + f == s;
                if (!match) allMatch = false;
                sb.AppendLine("    Diskte doğrulama: " + onDisk.ToString("N0") + " dosya, listede " + listed.ToString("N0") + " dosya  ->  " + (match ? "TUTUYOR" : "TUTMUYOR, günlüğe bakın"));
                foreach (string w in a.Warnings) sb.AppendLine("    DİKKAT: " + w);
                sb.AppendLine();
                srcAll += s; expAll += e; failAll += f;
            }
            if (accounts.Count == 0) sb.AppendLine("Hiç hesap aktarılmadı.").AppendLine();

            sb.AppendLine("GENEL TOPLAM: kaynakta " + srcAll.ToString("N0") + " öğe, aktarılan " + expAll.ToString("N0") + ", hata " + failAll.ToString("N0")
                + (allMatch && failAll == 0 && accounts.Count > 0 ? "  ->  EKSİKSİZ" : ""));
            foreach (string n in Notes) sb.AppendLine(n);
            sb.AppendLine();
            sb.AppendLine("SONRAKİ ADIMLAR");
            sb.AppendLine("---------------");
            sb.AppendLine("1. Bu klasörün tamamını flash belleğe kopyalayın (içindeki her şey gerekli).");
            sb.AppendLine("2. Outlook ile başka bir adrese almak için: yeni hesabı Outlook'a ekleyin, sonra");
            sb.AppendLine("   Dosya > Aç ve Dışarı Aktar > İçeri/Dışarı Aktar > Başka bir program veya dosyadan içeri aktar >");
            sb.AppendLine("   Outlook Veri Dosyası (.pst) > TumMailler.pst > \"Öğeleri şu klasöre aktar\" kısmında yeni hesabı seçin.");
            sb.AppendLine("3. EML dosyaları her mail programında açılır (çift tıklayın). Thunderbird'e klasör klasör sürüklenebilir;");
            sb.AppendLine("   IMAP ile başka bir sunucuya yüklemek için her hesap klasöründeki _liste.tsv tarih, okundu bilgisi");
            sb.AppendLine("   ve klasör rolünü (gelen, gonderilen...) tutar.");
            sb.AppendLine("4. mail_listesi.csv Excel'de açılır: hangi mail hangi dosyada, tek bakışta.");

            string path = Path.Combine(Root, "ozet_rapor.txt");
            string written = null;
            try
            {
                File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
                written = path;
                R.Info("Rapor yazıldı: " + path);
            }
            catch (Exception ex)
            {
                // Hedef gittiyse rapor yerel klasöre yazılır.
                try
                {
                    written = Path.Combine(Reporter.LocalDir(), "ozet_rapor_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".txt");
                    File.WriteAllText(written, sb.ToString(), new UTF8Encoding(true));
                    R.Warn("Rapor kayıt yerine yazılamadı (" + ex.Message + "), yerel kopya: " + written);
                }
                catch { written = null; }
            }
            if (R.LocalLogPath != null) R.Info("Günlüğün yerel kopyası: " + R.LocalLogPath);
            R.CloseFile();
            return written;
        }

        static string Pad(string s, int n)
        {
            s = s ?? "";
            if (s.Length > n - 1) s = s.Substring(0, n - 2) + "…";
            return s.PadRight(n);
        }

        static string LPad(string s, int n) { return (s ?? "").PadLeft(n); }
    }

    // Bir kaynak hesabın (Outlook hesabı, PST dosyası, sunucu hesabı) çıktısı.
    sealed class AccountSink : IDisposable
    {
        public readonly string Name, Source, Dir;
        readonly Package pkg;
        StreamWriter list, folders;
        readonly Dictionary<string, string> dirMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, HashSet<string>> used = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        public readonly List<FolderStat> Folders = new List<FolderStat>();
        public readonly List<string> Warnings = new List<string>();
        public long HeaderOnly, ListedFiles;
        public DateTime? Oldest, Newest;
        int rows;

        public Package Pkg { get { return pkg; } }

        internal AccountSink(Package p, string name, string source, string dir)
        {
            pkg = p;
            Name = name;
            Source = source;
            Dir = dir;
            Directory.CreateDirectory(dir);
            var enc = new UTF8Encoding(false);
            list = new StreamWriter(Path.Combine(dir, "_liste.tsv"), false, enc);
            list.WriteLine("#dosya\tklasor\trol\ttarih_utc\tbayraklar\tmessage_id\tkonu\tgonderen\tboyut\tdurum");
            folders = new StreamWriter(Path.Combine(dir, "_klasorler.tsv"), false, enc);
            folders.WriteLine("#klasor\trol\tkaynakta\taktarilan\thata");
            File.WriteAllText(Path.Combine(dir, "_hesap.txt"),
                "Hesap      : " + name + "\r\nKaynak     : " + source + "\r\nBilgisayar : " + Environment.MachineName +
                "\r\nTarih      : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "\r\n", new UTF8Encoding(true));
        }

        public string FolderDir(string logical)
        {
            string d;
            if (dirMap.TryGetValue(logical, out d)) return d;
            string[] segs = logical.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (segs.Length == 0) segs = new[] { "_" };
            string rel = string.Join("\\", segs.Select(s => Util.SafeName(s, 40)).ToArray());
            d = Path.Combine(Dir, rel);
            if (d.Length > 190)
                d = Path.Combine(Dir, "_uzun_yol", Util.SafeName(segs[segs.Length - 1], 30) + "_" + Util.ShortHash(logical));
            Directory.CreateDirectory(d);
            if (pkg.Resume)
            {
                // Önceki çalışmadan kalmış yarım geçici dosyalar
                try { foreach (string t in Directory.GetFiles(d, "~*.tmp")) File.Delete(t); }
                catch { }
            }
            dirMap[logical] = d;
            return d;
        }

        // Uzantısız, benzersiz tam yol döner (aynı ad hem .eml hem .msg için kullanılır).
        // Devam modunda diskteki dosyaya bakılmaz: aynı sıra aynı adları üretir, böylece
        // önceki çalışmada yazılmış mail kendi dosyasıyla eşleşir.
        public string ReserveBase(string dir, DateTime? utc, string subject)
        {
            string datePart = utc.HasValue ? utc.Value.ToLocalTime().ToString("yyyy-MM-dd_HHmm") : "tarihsiz";
            if (string.IsNullOrWhiteSpace(subject)) subject = "(konu yok)";
            int room = 250 - dir.Length - 1 - datePart.Length - 1 - 4 - 5;
            string subj = Util.SafeName(subject, Math.Max(4, Math.Min(60, room)));
            string b = datePart + "_" + subj;
            HashSet<string> set;
            if (!used.TryGetValue(dir, out set)) { set = new HashSet<string>(StringComparer.OrdinalIgnoreCase); used[dir] = set; }
            string cand = b;
            for (int n = 2; set.Contains(cand) || (!pkg.Resume && (File.Exists(Path.Combine(dir, cand + ".eml")) || File.Exists(Path.Combine(dir, cand + ".msg")))); n++)
                cand = b + "_" + n;
            set.Add(cand);
            return Path.Combine(dir, cand);
        }

        public string Relative(string fullPath)
        {
            return fullPath.StartsWith(Dir + "\\", StringComparison.OrdinalIgnoreCase) ? fullPath.Substring(Dir.Length + 1) : fullPath;
        }

        public void AddMessage(ManifestEntry e)
        {
            ListedFiles += e.FilesWritten;
            if (e.DateUtc.HasValue && e.DateUtc.Value.Year > 1971)
            {
                if (!Oldest.HasValue || e.DateUtc < Oldest) Oldest = e.DateUtc;
                if (!Newest.HasValue || e.DateUtc > Newest) Newest = e.DateUtc;
            }
            try
            {
                list.WriteLine(string.Join("\t", new[] {
                    Util.TsvClean(e.File), Util.TsvClean(e.Folder), e.Role ?? "",
                    e.DateUtc.HasValue ? Util.IsoUtc(e.DateUtc.Value) : "",
                    Util.TsvClean(e.Flags), Util.TsvClean(e.MessageId), Util.TsvClean(e.Subject), Util.TsvClean(e.From),
                    e.Size.ToString(CultureInfo.InvariantCulture), e.Status ?? "" }));
                pkg.CsvRow(this, e);
                if (++rows % 25 == 0) FlushHard();
            }
            catch (Exception ex)
            {
                pkg.CheckTarget(ex);
                throw;
            }
        }

        public void AddFolder(string logical, string role, long source, long exported, long failed)
        {
            Folders.Add(new FolderStat { Logical = logical, Role = role, Source = source, Exported = exported, Failed = failed });
            try
            {
                folders.WriteLine(Util.TsvClean(logical) + "\t" + (role ?? "") + "\t" + source + "\t" + exported + "\t" + failed);
                FlushHard();
            }
            catch (Exception ex) { pkg.CheckTarget(ex); }
        }

        public void FlushHard()
        {
            foreach (var w in new[] { list, folders })
            {
                if (w == null) continue;
                try
                {
                    w.Flush();
                    ((FileStream)w.BaseStream).Flush(true);
                }
                catch { }
            }
        }

        public long CountFilesOnDisk(string ext)
        {
            try { return Directory.EnumerateFiles(Dir, "*" + ext, SearchOption.AllDirectories).LongCount(f => f.EndsWith(ext, StringComparison.OrdinalIgnoreCase)); }
            catch { return -1; }
        }

        public void Dispose()
        {
            foreach (var w in new[] { list, folders })
            {
                if (w == null) continue;
                try { w.Dispose(); }
                catch { }
            }
            list = folders = null;
        }
    }
}
