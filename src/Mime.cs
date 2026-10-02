using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MailAktarici
{
    // RFC 2047 başlık kodlama / çözme
    static class Rfc2047
    {
        public static bool IsAscii(string s)
        {
            if (s == null) return true;
            foreach (char c in s)
                if (c > 126 || (c < 32 && c != '\t')) return false;
            return true;
        }

        // Yapısız başlık değeri (Subject vb.). Gerekirse birden fazla encoded-word, katlanmış.
        public static string Encode(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r", " ").Replace("\n", " ");
            if (IsAscii(s) && s.IndexOf("=?", StringComparison.Ordinal) < 0 && s.Length < 900) return s;
            var parts = new List<string>();
            int i = 0;
            while (i < s.Length)
            {
                int start = i, bytes = 0;
                while (i < s.Length)
                {
                    int len = (char.IsHighSurrogate(s[i]) && i + 1 < s.Length) ? 2 : 1;
                    int b = Encoding.UTF8.GetByteCount(s.Substring(i, len));
                    if (bytes > 0 && bytes + b > 45) break;
                    bytes += b;
                    i += len;
                }
                parts.Add("=?utf-8?B?" + Convert.ToBase64String(Encoding.UTF8.GetBytes(s.Substring(start, i - start))) + "?=");
            }
            return string.Join("\r\n ", parts);
        }

        // Adres görünen adı (phrase)
        public static string Phrase(string name)
        {
            if (!IsAscii(name)) return Encode(name);
            if (Regex.IsMatch(name, @"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~ -]+$")) return name;
            return "\"" + name.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        public static string Address(string name, string addr)
        {
            name = (name ?? "").Trim().Replace("\r", " ").Replace("\n", " ");
            addr = (addr ?? "").Trim();
            if (addr.Length == 0 && name.Length == 0) return null;
            if (addr.IndexOf('@') < 0)
            {
                if (name.Length == 0) name = addr;
                addr = "bilinmiyor@adres.invalid";
            }
            if (name.Length == 0 || string.Equals(name, addr, StringComparison.OrdinalIgnoreCase)) return addr;
            return Phrase(name) + " <" + addr + ">";
        }

        static readonly Regex EncodedWord = new Regex(@"=\?([^?\s]+)\?([BbQq])\?([^?\s]*)\?=");

        public static string Decode(string s)
        {
            if (string.IsNullOrEmpty(s) || s.IndexOf("=?", StringComparison.Ordinal) < 0) return s ?? "";
            var sb = new StringBuilder();
            int pos = 0;
            string pendCharset = null;
            var pend = new List<byte>();
            Match m = EncodedWord.Match(s);
            while (m.Success)
            {
                string between = s.Substring(pos, m.Index - pos);
                string cs = m.Groups[1].Value;
                int star = cs.IndexOf('*');
                if (star >= 0) cs = cs.Substring(0, star);
                bool adjacent = pendCharset != null && between.Trim().Length == 0;
                if (!adjacent || !string.Equals(pendCharset, cs, StringComparison.OrdinalIgnoreCase))
                {
                    Flush(sb, ref pendCharset, pend);
                    if (!adjacent) sb.Append(between);
                }
                pendCharset = cs;
                string data = m.Groups[3].Value;
                pend.AddRange(m.Groups[2].Value.ToUpperInvariant() == "B" ? B64(data) : QDecode(data));
                pos = m.Index + m.Length;
                m = m.NextMatch();
            }
            Flush(sb, ref pendCharset, pend);
            sb.Append(s.Substring(pos));
            return sb.ToString();
        }

        static void Flush(StringBuilder sb, ref string charset, List<byte> bytes)
        {
            if (charset == null) return;
            sb.Append(GetEncoding(charset).GetString(bytes.ToArray()));
            bytes.Clear();
            charset = null;
        }

        public static Encoding GetEncoding(string name)
        {
            try { return Encoding.GetEncoding(name.Trim().Trim('"')); }
            catch
            {
                string n = name.ToLowerInvariant();
                if (n.Contains("utf8") || n.Contains("utf-8")) return Encoding.UTF8;
                return Encoding.GetEncoding(28591);
            }
        }

        static byte[] B64(string s)
        {
            s = Regex.Replace(s, @"[^A-Za-z0-9+/]", "");
            while (s.Length % 4 != 0) s += "=";
            try { return Convert.FromBase64String(s); }
            catch { return new byte[0]; }
        }

        static byte[] QDecode(string s)
        {
            var o = new List<byte>();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '_') o.Add(32);
                else if (c == '=' && i + 2 < s.Length && Uri.IsHexDigit(s[i + 1]) && Uri.IsHexDigit(s[i + 2]))
                {
                    o.Add(Convert.ToByte(s.Substring(i + 1, 2), 16));
                    i += 2;
                }
                else o.Add((byte)c);
            }
            return o.ToArray();
        }

        // RFC 2231 parametre (filename*=utf-8''...), uzun adlar parçalanır.
        public static string Rfc2231Param(string param, string value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            var enc = new StringBuilder();
            foreach (byte b in bytes)
            {
                char c = (char)b;
                if (b < 128 && (char.IsLetterOrDigit(c) || "!#$&+-.^_`|~".IndexOf(c) >= 0)) enc.Append(c);
                else enc.Append('%').Append(b.ToString("X2"));
            }
            string e = enc.ToString();
            var chunks = new List<string>();
            int i = 0;
            while (i < e.Length)
            {
                int len = Math.Min(60, e.Length - i);
                if (i + len < e.Length)
                {
                    if (e[i + len - 1] == '%') len -= 1;
                    else if (len >= 2 && e[i + len - 2] == '%') len -= 2;
                }
                chunks.Add(e.Substring(i, len));
                i += len;
            }
            var sb = new StringBuilder();
            if (chunks.Count <= 1)
                sb.Append(";\r\n\t" + param + "*=utf-8''" + (chunks.Count == 1 ? chunks[0] : ""));
            else
                for (int k = 0; k < chunks.Count; k++)
                    sb.Append(";\r\n\t" + param + "*" + k + "*=" + (k == 0 ? "utf-8''" : "") + chunks[k]);
            return sb.ToString();
        }

        // Content-Type / Content-Disposition için ad parametreleri
        public static string NameParams(string param, string starParam, string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "";
            if (IsAscii(fileName) && fileName.Length <= 70 && fileName.IndexOf('"') < 0 && fileName.IndexOf('\\') < 0)
                return ";\r\n\t" + param + "=\"" + fileName + "\"";
            string ew = "=?utf-8?B?" + Convert.ToBase64String(Encoding.UTF8.GetBytes(fileName)) + "?=";
            string s = ";\r\n\t" + param + "=\"" + ew + "\"";
            if (starParam != null) s += Rfc2231Param(starParam, fileName);
            return s;
        }
    }

    // Ham başlık alanı: ad + katlanmış tam metin (CRLF'li)
    sealed class HeaderField
    {
        public string Name;
        public string Raw;
        public string Value
        {
            get
            {
                int c = Raw.IndexOf(':');
                string v = c < 0 ? "" : Raw.Substring(c + 1);
                return Regex.Replace(v, @"\r\n[ \t]+", " ").Trim();
            }
        }
    }

    static class Headers
    {
        static readonly Regex FieldStart = new Regex(@"^([!-9;-~]+)[ \t]*:");

        public static List<HeaderField> Parse(string block)
        {
            var list = new List<HeaderField>();
            if (string.IsNullOrEmpty(block)) return list;
            string[] lines = block.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            HeaderField cur = null;
            foreach (string line in lines)
            {
                if (line.Length == 0) { if (cur != null) break; else continue; }
                if ((line[0] == ' ' || line[0] == '\t'))
                {
                    if (cur != null) cur.Raw += "\r\n" + line;
                    continue;
                }
                var m = FieldStart.Match(line);
                if (m.Success)
                {
                    cur = new HeaderField { Name = m.Groups[1].Value, Raw = line };
                    list.Add(cur);
                }
                else cur = null;   // "Microsoft Mail Internet Headers Version 2.0" gibi çöp satırlar
            }
            return list;
        }

        public static string Get(List<HeaderField> h, string name)
        {
            foreach (var f in h)
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase)) return f.Value;
            return null;
        }

        public static bool Has(List<HeaderField> h, string name)
        {
            return Get(h, name) != null;
        }

        // Ham mail baytlarından başlık bloğunu metin olarak alır.
        public static string HeaderText(byte[] msg)
        {
            int limit = Math.Min(msg.Length, 256 * 1024);
            int end = limit;
            for (int i = 0; i + 1 < limit; i++)
            {
                if (msg[i] == '\n' && (msg[i + 1] == '\n' || (msg[i + 1] == '\r' && i + 2 < limit && msg[i + 2] == '\n')))
                {
                    end = i + 1;
                    break;
                }
            }
            string s = new UTF8Encoding(false, false).GetString(msg, 0, end);
            if (s.IndexOf('�') >= 0) s = Encoding.GetEncoding(1254).GetString(msg, 0, end);
            return s;
        }
    }

    // Basit MIME ağacı ve yazıcı
    sealed class Entity
    {
        public string ContentType;
        public readonly List<string> ExtraHeaders = new List<string>();
        public byte[] Data;
        public string TransferEncoding = "base64";
        public List<Entity> Parts;

        public static Entity Multipart(string subtype, params Entity[] parts)
        {
            var e = new Entity { ContentType = "multipart/" + subtype, Parts = new List<Entity>() };
            foreach (var p in parts) if (p != null) e.Parts.Add(p);
            return e;
        }

        public static Entity Text(string subtype, string text)
        {
            return new Entity
            {
                ContentType = "text/" + subtype + "; charset=utf-8",
                Data = Encoding.UTF8.GetBytes(Util.NormalizeNewlines(text ?? ""))
            };
        }
    }

    static class MimeWriter
    {
        static readonly byte[] Crlf = { 13, 10 };

        public static byte[] Build(List<string> topHeaders, Entity body)
        {
            using (var ms = new MemoryStream())
            {
                foreach (string h in topHeaders) Line(ms, h);
                Line(ms, "MIME-Version: 1.0");
                WriteEntity(ms, body);
                return ms.ToArray();
            }
        }

        static void Line(Stream s, string text)
        {
            byte[] b = Encoding.UTF8.GetBytes(text);
            s.Write(b, 0, b.Length);
            s.Write(Crlf, 0, 2);
        }

        static void WriteEntity(Stream s, Entity e)
        {
            if (e.Parts != null)
            {
                string b = "=_MailAktarici_" + Guid.NewGuid().ToString("N");
                Line(s, "Content-Type: " + e.ContentType + ";\r\n\tboundary=\"" + b + "\"");
                foreach (string h in e.ExtraHeaders) Line(s, h);
                Line(s, "");
                Line(s, "This is a multi-part message in MIME format.");
                foreach (Entity p in e.Parts)
                {
                    Line(s, "");
                    Line(s, "--" + b);
                    WriteEntity(s, p);
                }
                Line(s, "");
                Line(s, "--" + b + "--");
                return;
            }
            Line(s, "Content-Type: " + e.ContentType);
            Line(s, "Content-Transfer-Encoding: " + e.TransferEncoding);
            foreach (string h in e.ExtraHeaders) Line(s, h);
            Line(s, "");
            byte[] data = e.Data ?? new byte[0];
            if (e.TransferEncoding == "base64")
            {
                if (data.Length > 0) Line(s, Convert.ToBase64String(data, Base64FormattingOptions.InsertLineBreaks));
            }
            else
            {
                s.Write(data, 0, data.Length);
                if (data.Length < 2 || data[data.Length - 2] != 13 || data[data.Length - 1] != 10) s.Write(Crlf, 0, 2);
            }
        }

        static readonly Dictionary<string, string> Types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ".pdf", "application/pdf" }, { ".doc", "application/msword" },
            { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            { ".xls", "application/vnd.ms-excel" },
            { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            { ".ppt", "application/vnd.ms-powerpoint" },
            { ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
            { ".jpg", "image/jpeg" }, { ".jpeg", "image/jpeg" }, { ".png", "image/png" }, { ".gif", "image/gif" },
            { ".bmp", "image/bmp" }, { ".tif", "image/tiff" }, { ".tiff", "image/tiff" }, { ".svg", "image/svg+xml" },
            { ".webp", "image/webp" }, { ".txt", "text/plain" }, { ".htm", "text/html" }, { ".html", "text/html" },
            { ".csv", "text/csv" }, { ".xml", "application/xml" }, { ".json", "application/json" },
            { ".zip", "application/zip" }, { ".rar", "application/vnd.rar" }, { ".7z", "application/x-7z-compressed" },
            { ".eml", "message/rfc822" }, { ".msg", "application/vnd.ms-outlook" }, { ".ics", "text/calendar" },
            { ".vcf", "text/vcard" }, { ".mp3", "audio/mpeg" }, { ".wav", "audio/wav" }, { ".mp4", "video/mp4" },
            { ".udf", "application/octet-stream" }
        };

        // Veri kodlanmadan (7bit/8bit) MIME içine konabilir mi? NUL yoksa ve satırlar 998 baytı
        // aşmıyorsa satır sonları CRLF yapılmış kopyayı, değilse null döner.
        public static byte[] AsSafeText(byte[] data)
        {
            if (data == null) return null;
            var ms = new MemoryStream(data.Length + 64);
            int lineLen = 0;
            for (int i = 0; i < data.Length; i++)
            {
                byte b = data[i];
                if (b == 0) return null;
                if (b == 13 || b == 10)
                {
                    if (b == 13 && i + 1 < data.Length && data[i + 1] == 10) i++;
                    ms.WriteByte(13);
                    ms.WriteByte(10);
                    lineLen = 0;
                    continue;
                }
                if (++lineLen > 998) return null;
                ms.WriteByte(b);
            }
            return ms.ToArray();
        }

        public static string GuessType(string fileName)
        {
            string t;
            string f = fileName ?? "";
            int dot = f.LastIndexOf('.');
            string ext = dot < 0 ? "" : f.Substring(dot);
            return Types.TryGetValue(ext, out t) ? t : "application/octet-stream";
        }
    }
}
