using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;

namespace MailAktarici
{
    enum Security { SslTls, StartTls, None }

    sealed class ServerConfig
    {
        public string Protocol = "IMAP";   // IMAP | POP3
        public string Host;
        public int Port;
        public Security Sec = Security.SslTls;
        public string User, Pass;
        public bool IgnoreCert;

        public bool IsPop { get { return string.Equals(Protocol, "POP3", StringComparison.OrdinalIgnoreCase); } }

        public static int DefaultPort(bool pop, Security sec)
        {
            if (pop) return sec == Security.SslTls ? 995 : 110;
            return sec == Security.SslTls ? 993 : 143;
        }
    }

    class ProtocolException : Exception
    {
        public ProtocolException(string m) : base(m) { }
    }

    // Satır tabanlı TCP/TLS bağlantısı (IMAP ve POP3 ortak).
    abstract class LineConnection : IDisposable
    {
        TcpClient tcp;
        protected Stream Stream;
        readonly byte[] buf = new byte[1 << 16];
        int pos, len;
        protected string Host;
        string certError;

        protected void Connect(ServerConfig c)
        {
            Host = c.Host;
            tcp = new TcpClient();
            var t = tcp.ConnectAsync(c.Host, c.Port);
            try
            {
                if (!t.Wait(20000)) throw new IOException("Sunucuya 20 saniyede bağlanılamadı: " + c.Host + ":" + c.Port);
            }
            catch (AggregateException ex)
            {
                throw new IOException("Sunucuya bağlanılamadı (" + c.Host + ":" + c.Port + "): " + ex.InnerException.Message);
            }
            tcp.ReceiveTimeout = 180000;
            tcp.SendTimeout = 180000;
            Stream = tcp.GetStream();
            if (c.Sec == Security.SslTls) StartTls(c.IgnoreCert);
        }

        protected void StartTls(bool ignoreCert)
        {
            if (pos < len) throw new ProtocolException("TLS öncesi beklenmeyen veri");
            var ssl = new SslStream(Stream, false, (s, cert, chain, errors) =>
            {
                if (errors == SslPolicyErrors.None || ignoreCert) return true;
                certError = errors.ToString();
                return false;
            });
            try
            {
                ssl.AuthenticateAsClient(Host, null, (SslProtocols)(3072 | 768 | 192), false);   // TLS 1.2 / 1.1 / 1.0
            }
            catch (AuthenticationException ex)
            {
                if (certError != null)
                    throw new ProtocolException("Sunucu sertifikası doğrulanamadı (" + certError + "). Sunucu adını kontrol edin ya da " +
                        "'Sertifika hatalarını yoksay' kutusunu işaretleyin.");
                throw new ProtocolException("Güvenli bağlantı kurulamadı: " + ex.Message);
            }
            Stream = ssl;
        }

        void Fill()
        {
            len = Stream.Read(buf, 0, buf.Length);
            pos = 0;
            if (len <= 0) { len = 0; throw new IOException("Sunucu bağlantıyı kapattı."); }
        }

        // CRLF (ya da yalnız LF) hariç satır baytları
        protected byte[] ReadLineBytes()
        {
            var ms = new MemoryStream();
            while (true)
            {
                if (pos >= len) Fill();
                int idx = Array.IndexOf(buf, (byte)'\n', pos, len - pos);
                if (idx >= 0)
                {
                    ms.Write(buf, pos, idx - pos);
                    pos = idx + 1;
                    break;
                }
                ms.Write(buf, pos, len - pos);
                pos = len;
                if (ms.Length > 64 * 1024 * 1024) throw new ProtocolException("Sunucudan aşırı uzun satır geldi.");
            }
            byte[] line = ms.ToArray();
            if (line.Length > 0 && line[line.Length - 1] == '\r') Array.Resize(ref line, line.Length - 1);
            return line;
        }

        protected string ReadLine()
        {
            return Encoding.UTF8.GetString(ReadLineBytes());
        }

        protected byte[] ReadExact(int n)
        {
            var data = new byte[n];
            int off = 0;
            while (off < n)
            {
                if (pos >= len) Fill();
                int k = Math.Min(n - off, len - pos);
                Buffer.BlockCopy(buf, pos, data, off, k);
                pos += k;
                off += k;
            }
            return data;
        }

        protected void Send(string s)
        {
            byte[] b = Encoding.UTF8.GetBytes(s);
            Stream.Write(b, 0, b.Length);
            Stream.Flush();
        }

        public virtual void Dispose()
        {
            try { if (Stream != null) Stream.Dispose(); }
            catch { }
            try { if (tcp != null) tcp.Close(); }
            catch { }
        }
    }

    sealed class ImapResponse
    {
        public string Tag;
        public readonly List<string> Lines = new List<string>();
        public readonly List<byte[]> Literals = new List<byte[]>();
        public string First { get { return Lines.Count > 0 ? Lines[0] : ""; } }
        public string Joined { get { return string.Join(" ", Lines.ToArray()); } }

        public string Status
        {
            get
            {
                string[] p = First.Split(new[] { ' ' }, 3);
                return p.Length > 1 ? p[1].ToUpperInvariant() : "";
            }
        }
    }

    sealed class MailboxInfo
    {
        public string Raw, Name, Flags;
        public char Delim;
    }

    sealed class ImapMessageMeta
    {
        public uint Uid;
        public string Flags = "";
        public DateTime? InternalDate;
        public long Size;
    }

    sealed class ImapClient : LineConnection
    {
        int tagNo;
        public readonly HashSet<string> Caps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public string Greeting;

        static readonly Regex LiteralEnd = new Regex(@"\{(\d+)\+?\}$");

        public static ImapClient Open(ServerConfig c, Reporter r)
        {
            var im = new ImapClient();
            try
            {
                im.Connect(c);
                var g = im.ReadResponse();
                im.Greeting = g.First;
                if (!g.First.StartsWith("* OK", StringComparison.OrdinalIgnoreCase) && !g.First.StartsWith("* PREAUTH", StringComparison.OrdinalIgnoreCase))
                    throw new ProtocolException("IMAP sunucusu bağlantıyı kabul etmedi: " + g.First);
                im.Capability();
                if (c.Sec == Security.StartTls)
                {
                    if (!im.Caps.Contains("STARTTLS")) throw new ProtocolException("Sunucu STARTTLS desteklemiyor; 'SSL/TLS' ya da 'Şifresiz' seçin.");
                    im.Expect(im.Run("STARTTLS"), "STARTTLS");
                    im.StartTls(c.IgnoreCert);
                    im.Capability();
                }
                if (c.User != null && !g.First.StartsWith("* PREAUTH", StringComparison.OrdinalIgnoreCase)) im.Login(c.User, c.Pass ?? "");
                return im;
            }
            catch
            {
                im.Dispose();
                throw;
            }
        }

        string NextTag() { return "M" + (++tagNo).ToString("0000"); }

        ImapResponse ReadResponse()
        {
            var resp = new ImapResponse();
            while (true)
            {
                string line = ReadLine();
                resp.Lines.Add(line);
                var m = LiteralEnd.Match(line);
                if (!m.Success) break;
                long n = long.Parse(m.Groups[1].Value);
                if (n > int.MaxValue) throw new ProtocolException("Çok büyük veri bloğu");
                resp.Literals.Add(ReadExact((int)n));
            }
            int sp = resp.First.IndexOf(' ');
            resp.Tag = sp < 0 ? resp.First : resp.First.Substring(0, sp);
            return resp;
        }

        // Komutu gönderir, etiketli cevaba kadar okur. Etiketsiz cevaplar untagged listesine.
        ImapResponse Run(string cmd, List<ImapResponse> untagged = null)
        {
            string tag = NextTag();
            Send(tag + " " + cmd + "\r\n");
            return ReadUntilTagged(tag, untagged);
        }

        ImapResponse ReadUntilTagged(string tag, List<ImapResponse> untagged)
        {
            while (true)
            {
                var resp = ReadResponse();
                if (resp.Tag == tag) return resp;
                if (resp.Tag == "*" && resp.First.StartsWith("* BYE", StringComparison.OrdinalIgnoreCase))
                    throw new IOException("Sunucu bağlantıyı kapattı: " + resp.First);
                if (resp.Tag == "+") { Send("\r\n"); continue; }
                if (untagged != null) untagged.Add(resp);
            }
        }

        void Expect(ImapResponse tagged, string what)
        {
            if (tagged.Status != "OK")
                throw new ProtocolException(what + " başarısız: " + tagged.First);
        }

        void Capability()
        {
            var un = new List<ImapResponse>();
            var t = Run("CAPABILITY", un);
            Caps.Clear();
            foreach (var u in un)
            {
                if (!u.First.StartsWith("* CAPABILITY", StringComparison.OrdinalIgnoreCase)) continue;
                foreach (string c in u.First.Substring(12).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) Caps.Add(c);
            }
            if (Caps.Count == 0)
            {
                var m = Regex.Match(t.First, @"\[CAPABILITY ([^\]]*)\]", RegexOptions.IgnoreCase);
                if (m.Success) foreach (string c in m.Groups[1].Value.Split(' ')) Caps.Add(c);
            }
        }

        public static string Quote(string s)
        {
            return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        void Login(string user, string pass)
        {
            bool asciiSafe = Rfc2047.IsAscii(user) && Rfc2047.IsAscii(pass);
            ImapResponse t;
            if ((!asciiSafe || Caps.Contains("LOGINDISABLED")) && Caps.Contains("AUTH=PLAIN"))
            {
                string tag = NextTag();
                Send(tag + " AUTHENTICATE PLAIN\r\n");
                var cont = ReadResponse();
                if (cont.Tag != "+") throw new ProtocolException("Giriş başarısız: " + cont.First);
                Send(Convert.ToBase64String(Encoding.UTF8.GetBytes("\0" + user + "\0" + pass)) + "\r\n");
                t = ReadUntilTagged(tag, null);
            }
            else
            {
                t = Run("LOGIN " + Quote(user) + " " + Quote(pass));
            }
            if (t.Status != "OK")
                throw new ProtocolException("Giriş başarısız (kullanıcı adı/parola ya da sunucu izni): " + t.First.Substring(Math.Min(t.First.Length, t.First.IndexOf(' ') + 1)));
            Capability();
        }

        public List<MailboxInfo> List()
        {
            var un = new List<ImapResponse>();
            Expect(Run("LIST \"\" \"*\"", un), "Klasör listesi");
            var rx = new Regex(@"^\* (?:X?LIST) \(([^)]*)\) (NIL|""(?:[^""\\]|\\.)*"")\s?(.*)$", RegexOptions.IgnoreCase);
            var list = new List<MailboxInfo>();
            foreach (var u in un)
            {
                var m = rx.Match(u.First);
                if (!m.Success) continue;
                string delimTok = m.Groups[2].Value;
                char delim = delimTok.Equals("NIL", StringComparison.OrdinalIgnoreCase) ? '\0' : Unquote(delimTok)[0];
                string raw;
                if (u.Literals.Count > 0) raw = Encoding.UTF8.GetString(u.Literals[0]);
                else
                {
                    raw = m.Groups[3].Value.Trim();
                    if (raw.StartsWith("\"")) raw = Unquote(raw);
                }
                list.Add(new MailboxInfo { Raw = raw, Name = Util.ImapUtf7Decode(raw), Delim = delim, Flags = m.Groups[1].Value });
            }
            return list;
        }

        static string Unquote(string s)
        {
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"') s = s.Substring(1, s.Length - 2);
            return s.Replace("\\\"", "\"").Replace("\\\\", "\\");
        }

        public int Examine(string raw)
        {
            var un = new List<ImapResponse>();
            Expect(Run("EXAMINE " + Quote(raw), un), "Klasör açma (" + Util.ImapUtf7Decode(raw) + ")");
            int exists = 0;
            foreach (var u in un)
            {
                var m = Regex.Match(u.First, @"^\* (\d+) EXISTS", RegexOptions.IgnoreCase);
                if (m.Success) exists = int.Parse(m.Groups[1].Value);
            }
            return exists;
        }

        public int StatusMessages(string raw)
        {
            var un = new List<ImapResponse>();
            var t = Run("STATUS " + Quote(raw) + " (MESSAGES)", un);
            if (t.Status != "OK") return -1;
            foreach (var u in un)
            {
                var m = Regex.Match(u.Joined, @"MESSAGES (\d+)", RegexOptions.IgnoreCase);
                if (m.Success) return int.Parse(m.Groups[1].Value);
            }
            return -1;
        }

        public List<ImapMessageMeta> FetchMeta()
        {
            var un = new List<ImapResponse>();
            Expect(Run("UID FETCH 1:* (UID FLAGS INTERNALDATE RFC822.SIZE)", un), "Mail listesi");
            var byUid = new Dictionary<uint, ImapMessageMeta>();
            foreach (var u in un)
            {
                string s = u.Joined;
                if (!Regex.IsMatch(s, @"^\* \d+ FETCH", RegexOptions.IgnoreCase)) continue;
                var mu = Regex.Match(s, @"UID (\d+)", RegexOptions.IgnoreCase);
                if (!mu.Success) continue;
                uint uid = uint.Parse(mu.Groups[1].Value);
                ImapMessageMeta meta;
                if (!byUid.TryGetValue(uid, out meta)) { meta = new ImapMessageMeta { Uid = uid }; byUid[uid] = meta; }
                var mf = Regex.Match(s, @"FLAGS \(([^)]*)\)", RegexOptions.IgnoreCase);
                if (mf.Success) meta.Flags = mf.Groups[1].Value;
                var md = Regex.Match(s, @"INTERNALDATE ""([^""]+)""", RegexOptions.IgnoreCase);
                if (md.Success) meta.InternalDate = Util.ParseImapDate(md.Groups[1].Value);
                var ms = Regex.Match(s, @"RFC822\.SIZE (\d+)", RegexOptions.IgnoreCase);
                if (ms.Success) meta.Size = long.Parse(ms.Groups[1].Value);
            }
            var list = new List<ImapMessageMeta>(byUid.Values);
            list.Sort((a, b) => a.Uid.CompareTo(b.Uid));
            return list;
        }

        // Mailin ham hali (BODY.PEEK: okundu işareti değişmez). Silinmişse null.
        public byte[] FetchBody(uint uid)
        {
            var un = new List<ImapResponse>();
            Expect(Run("UID FETCH " + uid + " (UID BODY.PEEK[])", un), "Mail indirme");
            foreach (var u in un)
            {
                if (u.Literals.Count == 0) continue;
                for (int i = 0; i < u.Lines.Count - 1 && i < u.Literals.Count; i++)
                    if (u.Lines[i].IndexOf("BODY[]", StringComparison.OrdinalIgnoreCase) >= 0) return u.Literals[i];
            }
            return null;
        }

        public void Logout()
        {
            try { Run("LOGOUT"); }
            catch { }
        }
    }

    sealed class Pop3Client : LineConnection
    {
        public static Pop3Client Open(ServerConfig c)
        {
            var p = new Pop3Client();
            try
            {
                p.Connect(c);
                p.Expect(p.ReadLine(), "Karşılama");
                if (c.Sec == Security.StartTls)
                {
                    p.Cmd("STLS");
                    p.StartTls(c.IgnoreCert);
                }
                if (c.User != null)
                {
                    p.Cmd("USER " + c.User);
                    try { p.Cmd("PASS " + (c.Pass ?? "")); }
                    catch (ProtocolException ex) { throw new ProtocolException("Giriş başarısız (kullanıcı adı/parola): " + ex.Message); }
                }
                return p;
            }
            catch
            {
                p.Dispose();
                throw;
            }
        }

        string Expect(string line, string what)
        {
            if (!line.StartsWith("+OK", StringComparison.OrdinalIgnoreCase))
                throw new ProtocolException(what + ": " + line);
            return line;
        }

        string Cmd(string c)
        {
            Send(c + "\r\n");
            string verb = c.Split(' ')[0];
            return Expect(ReadLine(), verb);
        }

        byte[] ReadMultiline()
        {
            var ms = new MemoryStream();
            while (true)
            {
                byte[] line = ReadLineBytes();
                if (line.Length == 1 && line[0] == '.') break;
                int off = line.Length > 0 && line[0] == '.' ? 1 : 0;
                ms.Write(line, off, line.Length - off);
                ms.WriteByte(13);
                ms.WriteByte(10);
            }
            return ms.ToArray();
        }

        public List<KeyValuePair<int, string>> Uidl()
        {
            var list = new List<KeyValuePair<int, string>>();
            try { Cmd("UIDL"); }
            catch (ProtocolException) { return null; }
            foreach (string l in Encoding.UTF8.GetString(ReadMultiline()).Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] p = l.Split(' ');
                int n;
                if (p.Length >= 2 && int.TryParse(p[0], out n)) list.Add(new KeyValuePair<int, string>(n, p[1]));
            }
            return list;
        }

        public int Stat()
        {
            string s = Cmd("STAT");
            string[] p = s.Split(' ');
            int n;
            return p.Length > 1 && int.TryParse(p[1], out n) ? n : 0;
        }

        public byte[] Retr(int n)
        {
            Cmd("RETR " + n);
            return ReadMultiline();
        }

        // QUIT: DELE hiç gönderilmediği için sunucudan hiçbir şey silinmez.
        public void Quit()
        {
            try { Cmd("QUIT"); }
            catch { }
        }
    }
}
