using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace MailAktarici
{
    static class Util
    {
        static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();
        static readonly string[] Reserved = { "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };

        // Dosya/klasör adı olarak kullanılabilir hale getirir, en fazla max karakter.
        public static string SafeName(string s, int max)
        {
            if (s == null) s = "";
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
                sb.Append(c < 32 || Array.IndexOf(InvalidChars, c) >= 0 ? '_' : c);
            string r = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
            if (r.Length > max)
            {
                int cut = max;
                if (cut > 0 && char.IsHighSurrogate(r[cut - 1])) cut--;
                r = r.Substring(0, cut);
            }
            while (r.Length > 0 && (r[r.Length - 1] == '.' || r[r.Length - 1] == ' '))
                r = r.Substring(0, r.Length - 1);
            r = r.TrimStart(' ');
            if (r.Length == 0) r = "_";
            string stem = r.Split('.')[0].ToUpperInvariant();
            if (Array.IndexOf(Reserved, stem) >= 0) r = "_" + r;
            return r;
        }

        public static string ShortHash(string s)
        {
            using (var sha = SHA1.Create())
            {
                byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(s ?? ""));
                return BitConverter.ToString(h, 0, 4).Replace("-", "").ToLowerInvariant();
            }
        }

        public static string HexHash(byte[] data)
        {
            using (var sha = SHA1.Create())
                return BitConverter.ToString(sha.ComputeHash(data ?? new byte[0])).Replace("-", "").ToLowerInvariant();
        }

        public static string TsvClean(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char c in s) sb.Append(c == '\t' || c == '\r' || c == '\n' ? ' ' : c);
            return sb.ToString().Trim();
        }

        public static string FormatSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024L * 1024) return (bytes / 1024.0).ToString("0.#") + " KB";
            if (bytes < 1024L * 1024 * 1024) return (bytes / 1048576.0).ToString("0.#") + " MB";
            return (bytes / 1073741824.0).ToString("0.##") + " GB";
        }

        // ---- Tarihler ----

        public static string Rfc2822Date(DateTime utc)
        {
            DateTime local = DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToLocalTime();
            return local.ToString("ddd, dd MMM yyyy HH:mm:ss ", CultureInfo.InvariantCulture)
                + Offset(TimeZoneInfo.Local.GetUtcOffset(local));
        }

        public static string IsoUtc(DateTime utc)
        {
            return DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }

        static string Offset(TimeSpan off)
        {
            string sign = off < TimeSpan.Zero ? "-" : "+";
            off = off.Duration();
            return sign + off.Hours.ToString("00") + off.Minutes.ToString("00");
        }

        static readonly string[] Months = { "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec" };

        static int MonthIndex(string m)
        {
            return Array.IndexOf(Months, m.Substring(0, Math.Min(3, m.Length)).ToLowerInvariant()) + 1;
        }

        // IMAP INTERNALDATE: "17-Jul-1996 02:44:25 -0700"
        public static DateTime? ParseImapDate(string s)
        {
            var m = Regex.Match(s ?? "", @"(\d{1,2})-([A-Za-z]{3})-(\d{4}) (\d{1,2}):(\d{2}):(\d{2}) ([+-])(\d{2})(\d{2})");
            if (!m.Success) return null;
            int mon = MonthIndex(m.Groups[2].Value);
            if (mon == 0) return null;
            try
            {
                var dt = new DateTime(int.Parse(m.Groups[3].Value), mon, int.Parse(m.Groups[1].Value),
                    int.Parse(m.Groups[4].Value), int.Parse(m.Groups[5].Value), int.Parse(m.Groups[6].Value));
                var off = new TimeSpan(int.Parse(m.Groups[8].Value), int.Parse(m.Groups[9].Value), 0);
                if (m.Groups[7].Value == "-") off = off.Negate();
                return new DateTimeOffset(dt, off).UtcDateTime;
            }
            catch { return null; }
        }

        // Mail "Date:" başlığı, hoşgörülü okuma.
        public static DateTime? ParseRfc2822Date(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            s = Regex.Replace(s, @"\([^)]*\)", " ");
            var m = Regex.Match(s, @"(\d{1,2})\s+([A-Za-z]{3})[A-Za-z]*\.?\s+(\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*([+-]\d{4}|[A-Za-z]{1,5})?");
            if (!m.Success) return null;
            int mon = MonthIndex(m.Groups[2].Value);
            if (mon == 0) return null;
            try
            {
                int year = int.Parse(m.Groups[3].Value);
                if (year < 100) year += year < 50 ? 2000 : 1900;
                int sec = m.Groups[6].Success ? int.Parse(m.Groups[6].Value) : 0;
                if (sec > 59) sec = 59;
                var dt = new DateTime(year, mon, int.Parse(m.Groups[1].Value),
                    int.Parse(m.Groups[4].Value), int.Parse(m.Groups[5].Value), sec);
                TimeSpan off = TimeSpan.Zero;
                string z = m.Groups[7].Success ? m.Groups[7].Value : "";
                if (z.Length == 5 && (z[0] == '+' || z[0] == '-'))
                {
                    off = new TimeSpan(int.Parse(z.Substring(1, 2)), int.Parse(z.Substring(3, 2)), 0);
                    if (z[0] == '-') off = off.Negate();
                }
                else
                {
                    switch (z.ToUpperInvariant())
                    {
                        case "EST": off = TimeSpan.FromHours(-5); break;
                        case "EDT": off = TimeSpan.FromHours(-4); break;
                        case "CST": off = TimeSpan.FromHours(-6); break;
                        case "CDT": off = TimeSpan.FromHours(-5); break;
                        case "MST": off = TimeSpan.FromHours(-7); break;
                        case "MDT": off = TimeSpan.FromHours(-6); break;
                        case "PST": off = TimeSpan.FromHours(-8); break;
                        case "PDT": off = TimeSpan.FromHours(-7); break;
                    }
                }
                if (off.Duration() > TimeSpan.FromHours(14)) off = TimeSpan.Zero;
                return new DateTimeOffset(dt, off).UtcDateTime;
            }
            catch { return null; }
        }

        // ---- IMAP klasör adları (değiştirilmiş UTF-7, RFC 3501 5.1.3) ----

        public static string ImapUtf7Decode(string s)
        {
            if (string.IsNullOrEmpty(s) || s.IndexOf('&') < 0) return s ?? "";
            var sb = new StringBuilder();
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (c != '&') { sb.Append(c); i++; continue; }
                int j = s.IndexOf('-', i + 1);
                if (j < 0) { sb.Append(s.Substring(i)); break; }
                if (j == i + 1) { sb.Append('&'); i = j + 1; continue; }
                string b64 = s.Substring(i + 1, j - i - 1).Replace(',', '/');
                while (b64.Length % 4 != 0) b64 += "=";
                try { sb.Append(Encoding.BigEndianUnicode.GetString(Convert.FromBase64String(b64))); }
                catch { sb.Append(s.Substring(i, j - i + 1)); }
                i = j + 1;
            }
            return sb.ToString();
        }

        public static string ImapUtf7Encode(string s)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (c >= 0x20 && c <= 0x7e)
                {
                    sb.Append(c == '&' ? "&-" : c.ToString());
                    i++;
                    continue;
                }
                int j = i;
                while (j < s.Length && (s[j] < 0x20 || s[j] > 0x7e)) j++;
                string b64 = Convert.ToBase64String(Encoding.BigEndianUnicode.GetBytes(s.Substring(i, j - i)));
                sb.Append('&').Append(b64.TrimEnd('=').Replace('/', ',')).Append('-');
                i = j;
            }
            return sb.ToString();
        }

        // ---- Satır sonları ----

        public static string NormalizeNewlines(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            return s.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        public static string DesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        }

        public static bool IsCloudSynced(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string p = path.ToLowerInvariant();
            return p.Contains("\\onedrive") || p.Contains("\\google drive") || p.Contains("\\dropbox") || p.Contains("\\icloud");
        }
    }

    // Klasör rolleri; ileride başka bir hesaba yüklerken hangi klasörün "Gönderilmiş"
    // olduğunu dil bağımsız bilmek için listeye yazılır.
    static class Roles
    {
        public const string Inbox = "gelen", Sent = "gonderilen", Drafts = "taslak", Trash = "silinen",
            Junk = "istenmeyen", Archive = "arsiv", Outbox = "giden";

        static readonly Dictionary<string, string> ByName = Build();

        static Dictionary<string, string> Build()
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Action<string, string[]> add = (role, names) => { foreach (var n in names) d[Norm(n)] = role; };
            add(Inbox, new[] { "inbox", "gelen kutusu", "posteingang" });
            add(Sent, new[] { "sent", "sent items", "sent messages", "sent mail", "gönderilmiş öğeler",
                "gönderilenler", "gönderilen", "gönderilmiş", "gönderilmiş postalar", "gesendete elemente" });
            add(Drafts, new[] { "drafts", "draft", "taslaklar", "taslak", "entwürfe" });
            add(Trash, new[] { "trash", "deleted items", "deleted messages", "bin", "silinmiş öğeler",
                "silinenler", "çöp kutusu", "çöp", "gelöschte elemente" });
            add(Junk, new[] { "junk", "junk e-mail", "junk email", "spam", "bulk mail", "önemsiz e-posta",
                "önemsiz posta", "istenmeyen", "istenmeyen e-posta", "gereksiz", "gereksiz e-posta" });
            add(Archive, new[] { "archive", "archives", "arşiv", "arsiv" });
            add(Outbox, new[] { "outbox", "giden kutusu" });
            return d;
        }

        static string Norm(string s)
        {
            return (s ?? "").Trim().ToLowerInvariant().Replace("̇", "").Replace('ı', 'i');
        }

        public static string FromName(string name)
        {
            string r;
            return ByName.TryGetValue(Norm(name), out r) ? r : "";
        }

        public static string FromImapFlags(string flags)
        {
            string f = (flags ?? "").ToLowerInvariant();
            if (f.Contains("\\sent")) return Sent;
            if (f.Contains("\\drafts")) return Drafts;
            if (f.Contains("\\trash")) return Trash;
            if (f.Contains("\\junk") || f.Contains("\\spam")) return Junk;
            if (f.Contains("\\archive")) return Archive;
            return "";
        }
    }

    // Excel'in doğrudan açabildiği CSV (UTF-8 BOM, bölge ayarındaki liste ayırıcı).
    sealed class CsvWriter : IDisposable
    {
        readonly StreamWriter w;
        readonly string sep;

        public CsvWriter(string path, params string[] header)
        {
            w = new StreamWriter(path, false, new UTF8Encoding(true));
            sep = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            if (string.IsNullOrEmpty(sep)) sep = ";";
            Row(header);
        }

        public void Row(params string[] cells)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < cells.Length; i++)
            {
                if (i > 0) sb.Append(sep);
                string c = cells[i] ?? "";
                if (c.IndexOf(sep, StringComparison.Ordinal) >= 0 || c.IndexOf('"') >= 0 || c.IndexOf('\n') >= 0 || c.IndexOf('\r') >= 0)
                    c = "\"" + c.Replace("\"", "\"\"") + "\"";
                sb.Append(c);
            }
            lock (w) { w.WriteLine(sb.ToString()); }
        }

        public void FlushHard()
        {
            lock (w)
            {
                w.Flush();
                ((FileStream)w.BaseStream).Flush(true);
            }
        }

        public void Dispose() { lock (w) { w.Dispose(); } }
    }
}
