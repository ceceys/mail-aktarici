using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailAktarici
{
    // Sunucudan doğrudan indirme. Mailler olduğu gibi (ham) .eml olarak yazılır;
    // sunucuda hiçbir şey değişmez: IMAP'te EXAMINE + BODY.PEEK, POP3'te DELE yok.
    static class ServerExporter
    {
        public static void Run(ServerConfig c, Package pkg, Reporter r, int limitPerFolder)
        {
            if (c.IsPop) RunPop3(c, pkg, r, limitPerFolder);
            else RunImap(c, pkg, r, limitPerFolder);
        }

        static void RunImap(ServerConfig c, Package pkg, Reporter r, int limit)
        {
            r.Info("== IMAP sunucusundan indiriliyor: " + c.User + " @ " + c.Host + ":" + c.Port);
            ImapClient im = ImapClient.Open(c, r);
            try
            {
                List<MailboxInfo> boxes = im.List().Where(b => b.Flags.IndexOf("\\Noselect", StringComparison.OrdinalIgnoreCase) < 0
                    && b.Flags.IndexOf("\\NonExistent", StringComparison.OrdinalIgnoreCase) < 0).ToList();
                var counts = new Dictionary<string, int>();
                foreach (var b in boxes)
                {
                    int n = im.StatusMessages(b.Raw);
                    counts[b.Raw] = n;
                    if (n > 0) r.AddTotal(limit > 0 ? Math.Min(n, limit) : n);
                }
                r.Info("   " + boxes.Count + " klasör, " + counts.Values.Where(v => v > 0).Sum().ToString("N0") + " mail");
                AccountSink sink = pkg.OpenAccount(c.User + " (IMAP)", "IMAP sunucusu " + c.Host);
                foreach (var b in boxes)
                {
                    r.Check();
                    string logical = b.Delim != '\0' ? b.Name.Replace('\\', '_').Replace(b.Delim, '\\') : b.Name.Replace('\\', '_');
                    string top = logical.Split('\\')[0];
                    string role = string.Equals(b.Name, "INBOX", StringComparison.OrdinalIgnoreCase) ? Roles.Inbox : Roles.FromImapFlags(b.Flags);
                    if (role == "" && logical.IndexOf('\\') < 0) role = Roles.FromName(top);
                    if (role == "" && top.Equals("INBOX", StringComparison.OrdinalIgnoreCase) && logical.Count(ch => ch == '\\') == 1)
                        role = Roles.FromName(logical.Substring(6));   // INBOX.Sent gibi ad alanları
                    for (int attempt = 1; attempt <= 2; attempt++)
                    {
                        try
                        {
                            im = ExportMailbox(im, c, r, sink, b, logical, role, limit);
                            break;
                        }
                        catch (ProtocolException ex)
                        {
                            r.Error(logical + ": " + ex.Message);
                            break;
                        }
                        catch (IOException ex)
                        {
                            if (attempt == 2) { r.Error(logical + ": " + ex.Message); break; }
                            r.Warn("Bağlantı koptu (" + ex.Message + "), yeniden bağlanılıyor...");
                            im.Dispose();
                            im = ImapClient.Open(c, r);
                        }
                    }
                }
            }
            finally
            {
                im.Logout();
                im.Dispose();
            }
        }

        static ImapClient ExportMailbox(ImapClient im, ServerConfig c, Reporter r, AccountSink sink, MailboxInfo b, string logical, string role, int limit)
        {
            r.Status = "IMAP › " + logical;
            string dir = sink.FolderDir(logical);
            int exists = im.Examine(b.Raw);
            if (exists == 0)
            {
                sink.AddFolder(logical, role, 0, 0, 0);
                return im;
            }
            List<ImapMessageMeta> metas = im.FetchMeta();
            if (limit > 0 && metas.Count > limit) metas = metas.Take(limit).ToList();
            long exported = 0, failed = 0;
            foreach (var meta in metas)
            {
                r.Check();
                byte[] data = null;
                for (int attempt = 1; attempt <= 4 && data == null; attempt++)
                {
                    try
                    {
                        data = im.FetchBody(meta.Uid);
                        if (data == null) break;
                    }
                    catch (IOException ex)
                    {
                        if (attempt == 4) { r.Error(logical + " UID " + meta.Uid + ": " + ex.Message); break; }
                        r.Warn("Bağlantı koptu (" + ex.Message + "), yeniden bağlanılıyor...");
                        im.Dispose();
                        System.Threading.Thread.Sleep(2000 * attempt);
                        im = ImapClient.Open(c, r);
                        im.Examine(b.Raw);
                    }
                }
                r.Step();
                if (data == null)
                {
                    failed++;
                    sink.AddMessage(new ManifestEntry { Folder = logical, Role = role, Status = "HATA" });
                    continue;
                }
                SaveRaw(sink, dir, logical, role, data, meta.InternalDate, SystemFlags(meta.Flags));
                exported++;
            }
            sink.AddFolder(logical, role, metas.Count, exported, failed);
            r.Info("   " + logical + ": " + exported + "/" + metas.Count + (failed > 0 ? ", " + failed + " hata" : ""));
            return im;
        }

        static string SystemFlags(string flags)
        {
            var keep = new[] { "\\Seen", "\\Answered", "\\Flagged", "\\Draft", "$Forwarded" };
            return string.Join(" ", (flags ?? "").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(f => keep.Contains(f, StringComparer.OrdinalIgnoreCase)).ToArray());
        }

        static void SaveRaw(AccountSink sink, string dir, string logical, string role, byte[] data, DateTime? internalDate, string flags)
        {
            var h = Headers.Parse(Headers.HeaderText(data));
            string subject = Rfc2047.Decode(Headers.Get(h, "Subject") ?? "");
            DateTime? date = internalDate ?? Util.ParseRfc2822Date(Headers.Get(h, "Date"));
            string path = sink.ReserveBase(dir, date, subject) + ".eml";
            if (sink.Pkg.Resume && Package.LooksComplete(path)) sink.Pkg.Reused++;
            else
            {
                try { Package.WriteAtomic(path, data); }
                catch (Exception ex)
                {
                    sink.Pkg.CheckTarget(ex);
                    throw;
                }
            }
            if (date.HasValue)
            {
                try { File.SetLastWriteTimeUtc(path, date.Value); File.SetCreationTimeUtc(path, date.Value); }
                catch { }
            }
            sink.AddMessage(new ManifestEntry
            {
                File = sink.Relative(path), Folder = logical, Role = role, DateUtc = date, Flags = flags,
                MessageId = Headers.Get(h, "Message-ID"), Subject = subject, From = Rfc2047.Decode(Headers.Get(h, "From") ?? ""),
                Size = data.Length, Status = "TAM", FilesWritten = 1
            });
        }

        static void RunPop3(ServerConfig c, Package pkg, Reporter r, int limit)
        {
            r.Info("== POP3 sunucusundan indiriliyor: " + c.User + " @ " + c.Host + ":" + c.Port);
            Pop3Client p = Pop3Client.Open(c);
            try
            {
                int count = p.Stat();
                int take = limit > 0 ? Math.Min(limit, count) : count;
                r.AddTotal(take);
                r.Info("   Sunucuda " + count.ToString("N0") + " mail var (POP3 yalnız gelen kutusunu görür).");
                AccountSink sink = pkg.OpenAccount(c.User + " (POP3)", "POP3 sunucusu " + c.Host);
                string logical = "Gelen Kutusu";
                string dir = sink.FolderDir(logical);
                long exported = 0, failed = 0;
                for (int n = 1; n <= take; n++)
                {
                    r.Check();
                    r.Status = "POP3 › " + n + " / " + take;
                    try
                    {
                        byte[] data = p.Retr(n);
                        // POP3'te okundu bilgisi yok; eski mailler olduğu için okundu sayılır.
                        SaveRaw(sink, dir, logical, Roles.Inbox, data, null, "\\Seen");
                        exported++;
                    }
                    catch (ProtocolException ex)
                    {
                        failed++;
                        r.Error("POP3 mail #" + n + ": " + ex.Message);
                        sink.AddMessage(new ManifestEntry { Folder = logical, Role = Roles.Inbox, Status = "HATA" });
                    }
                    r.Step();
                }
                sink.AddFolder(logical, Roles.Inbox, take, exported, failed);
                r.Info("   " + logical + ": " + exported + "/" + take);
            }
            finally
            {
                p.Quit();
                p.Dispose();
            }
        }

        // Arayüzdeki "Bağlantıyı Test Et" düğmesi
        public static string TestConnection(ServerConfig c, Reporter r)
        {
            if (c.IsPop)
            {
                using (var p = Pop3Client.Open(c))
                {
                    int n = c.User != null ? p.Stat() : -1;
                    p.Quit();
                    return c.User != null ? "Bağlantı ve giriş başarılı. Sunucuda " + n.ToString("N0") + " mail var." : "Bağlantı başarılı (giriş denenmedi).";
                }
            }
            using (var im = ImapClient.Open(c, r))
            {
                string res;
                if (c.User != null)
                {
                    var boxes = im.List();
                    res = "Bağlantı ve giriş başarılı. " + boxes.Count + " klasör bulundu.";
                }
                else res = "Bağlantı başarılı (giriş denenmedi). Sunucu: " + im.Greeting;
                im.Logout();
                return res;
            }
        }
    }
}
