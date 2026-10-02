using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MailAktarici
{
    sealed class ExportOptions
    {
        public bool Eml = true;
        public bool Msg = false;
        public bool SinglePst = true;
        public int LimitPerFolder = 0;   // 0 = hepsi (sınır yalnız deneme içindir)
    }

    // Outlook'taki bir kaynak: profildeki hesap/veri dosyası ya da diskten eklenen PST.
    sealed class OutlookSource
    {
        public object Store;
        public string Name, Kind, FilePath, StoreId;
        public bool AttachedByUs;
        public string TempCopy;
        public long Items, Bytes;
        public bool DefaultSelected = true;

        public override string ToString()
        {
            string s = Name + "   —   " + Kind;
            if (Items > 0 || Bytes > 0) s += "   (" + Items.ToString("N0") + " mail, " + Util.FormatSize(Bytes) + ")";
            return s;
        }
    }

    sealed class ItemInfo
    {
        public string Headers, MessageId, SenderSmtp, SenderName, SenderEmail, SenderType, Subject, MessageClass, InReplyTo, References;
        public int MsgFlags, FlagStatus, LastVerb;
        public long Size;
        public DateTime? Delivery, Submit, Created;

        public DateTime? BestDate { get { return Delivery ?? Submit ?? Created; } }
    }

    sealed class OutlookExporter
    {
        const int OlMail = 0, OlPost = 6;
        readonly object ns;
        readonly Package pkg;
        readonly Reporter r;
        readonly ExportOptions opt;
        readonly string temp;
        string currentSubject;   // uyarı mesajlarında hangi mail olduğu görünsün

        static readonly object[] ItemTags =
        {
            Com.Tag("007D001F"), // 0  PR_TRANSPORT_MESSAGE_HEADERS
            Com.Tag("1035001F"), // 1  PR_INTERNET_MESSAGE_ID
            Com.Tag("0E070003"), // 2  PR_MESSAGE_FLAGS
            Com.Tag("10900003"), // 3  PR_FLAG_STATUS
            Com.Tag("10810003"), // 4  PR_LAST_VERB_EXECUTED
            Com.Tag("0E060040"), // 5  PR_MESSAGE_DELIVERY_TIME
            Com.Tag("00390040"), // 6  PR_CLIENT_SUBMIT_TIME
            Com.Tag("30070040"), // 7  PR_CREATION_TIME
            Com.Tag("0E080003"), // 8  PR_MESSAGE_SIZE
            Com.Tag("5D01001F"), // 9  PR_SENDER_SMTP_ADDRESS
            Com.Tag("0C1A001F"), // 10 PR_SENDER_NAME
            Com.Tag("0C1F001F"), // 11 PR_SENDER_EMAIL_ADDRESS
            Com.Tag("0C1E001F"), // 12 PR_SENDER_ADDRTYPE
            Com.Tag("0037001F"), // 13 PR_SUBJECT
            Com.Tag("001A001F"), // 14 PR_MESSAGE_CLASS
            Com.Tag("1042001F"), // 15 PR_IN_REPLY_TO_ID
            Com.Tag("1039001F"), // 16 PR_INTERNET_REFERENCES
        };

        static readonly object[] AttTags =
        {
            Com.Tag("370E001F"), // 0 PR_ATTACH_MIME_TAG
            Com.Tag("3712001F"), // 1 PR_ATTACH_CONTENT_ID
            Com.Tag("7FFE000B"), // 2 PR_ATTACHMENT_HIDDEN
        };

        static readonly string TagHidden = Com.Tag("10F4000B");
        static readonly string TagSize = Com.Tag("0E080003");
        static readonly string TagSmtp = Com.Tag("39FE001F");

        static readonly HashSet<string> DropHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Content-Type", "Content-Transfer-Encoding", "MIME-Version", "Content-Disposition", "Content-ID",
            "Content-Description", "Content-Length", "Lines", "Content-MD5"
        };

        public OutlookExporter(OutlookSession s, Package pkg, Reporter r, ExportOptions opt)
        {
            ns = s.Ns;
            this.pkg = pkg;
            this.r = r;
            this.opt = opt ?? new ExportOptions();
            temp = Path.Combine(Path.GetTempPath(), "MailAktarici_" + System.Diagnostics.Process.GetCurrentProcess().Id);
            Directory.CreateDirectory(temp);
        }

        public void Cleanup()
        {
            try { Directory.Delete(temp, true); }
            catch { }
        }

        // ---------------- Kaynaklar ----------------

        public static List<OutlookSource> ListProfileStores(object ns)
        {
            var list = new List<OutlookSource>();
            object stores = Com.Get(ns, "Stores");
            int n = Com.Int(stores, "Count", 0);
            for (int i = 1; i <= n; i++)
            {
                object st = Com.Item(stores, i);
                int exType = Com.Int(st, "ExchangeStoreType", 3);
                string fp = Com.Str(st, "FilePath") ?? "";
                string kind;
                bool sel = true;
                switch (exType)
                {
                    case 0: kind = "Exchange / Microsoft 365 posta kutusu"; break;
                    case 1: kind = "Başkasının Exchange posta kutusu (yetkili erişim)"; sel = false; break;
                    case 2: kind = "Exchange ortak klasörleri"; sel = false; break;
                    case 4: kind = "Ek Exchange posta kutusu"; break;
                    default:
                        if (fp.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)) kind = "PST veri dosyası (POP hesabı / arşiv)";
                        else if (fp.EndsWith(".nst", StringComparison.OrdinalIgnoreCase)) { kind = "Grup/paylaşılan veri"; sel = false; }
                        else kind = "IMAP hesabı";
                        break;
                }
                string name = Com.Str(st, "DisplayName") ?? ("Hesap " + i);
                if (fp.EndsWith("Internet Calendar Subscriptions.pst", StringComparison.OrdinalIgnoreCase) ||
                    fp.EndsWith("SharePoint Lists.pst", StringComparison.OrdinalIgnoreCase))
                    sel = false;
                list.Add(new OutlookSource
                {
                    Store = st, Name = name, Kind = kind, FilePath = fp,
                    StoreId = Com.Str(st, "StoreID"), DefaultSelected = sel
                });
            }
            Com.Release(stores);
            return list;
        }

        object FindStoreByPath(string path)
        {
            string want = OutlookSession.NormalizePath(path);
            object stores = Com.Get(ns, "Stores");
            int n = Com.Int(stores, "Count", 0);
            object found = null;
            for (int i = 1; i <= n && found == null; i++)
            {
                object st = Com.Item(stores, i);
                string fp = Com.Str(st, "FilePath");
                if (!string.IsNullOrEmpty(fp) && OutlookSession.NormalizePath(fp) == want) found = st;
                else Com.Release(st);
            }
            Com.Release(stores);
            return found;
        }

        public OutlookSource AttachPst(string path)
        {
            object st = FindStoreByPath(path);
            bool attached = false;
            string tempCopy = null;
            if (st == null)
            {
                try { Com.Call(ns, "AddStore", path); }
                catch (Exception ex)
                {
                    // Salt okunur ya da başka bir yerde kilitli PST: geçici kopyası açılır.
                    r.Warn("PST doğrudan açılamadı (" + Com.Message(ex) + "), geçici kopyası deneniyor: " + path);
                    tempCopy = Path.Combine(temp, Guid.NewGuid().ToString("N").Substring(0, 8) + "_" + Path.GetFileName(path));
                    File.Copy(path, tempCopy);
                    File.SetAttributes(tempCopy, FileAttributes.Normal);
                    Com.Call(ns, "AddStore", tempCopy);
                }
                attached = true;
                st = FindStoreByPath(tempCopy ?? path);
            }
            if (st == null) throw new ApplicationException("PST dosyası Outlook'ta açılamadı: " + path);
            return new OutlookSource
            {
                Store = st,
                Name = "PST - " + Path.GetFileNameWithoutExtension(path),
                Kind = "PST dosyası: " + path,
                FilePath = path,
                StoreId = Com.Str(st, "StoreID"),
                AttachedByUs = attached,
                TempCopy = tempCopy
            };
        }

        public void Detach(OutlookSource src)
        {
            if (!src.AttachedByUs) return;
            try
            {
                object root = Com.Call(src.Store, "GetRootFolder");
                Com.Call(ns, "RemoveStore", root);
                Com.Release(root);
                r.Info("Outlook'tan çıkarıldı: " + src.FilePath);
            }
            catch (Exception ex) { r.Warn("PST Outlook'tan çıkarılamadı (" + src.FilePath + "): " + Com.Message(ex)); }
            if (src.TempCopy != null)
            {
                try { File.Delete(src.TempCopy); }
                catch { }
            }
        }

        // ---------------- Klasör gezme ----------------

        sealed class StoreMap
        {
            public readonly Dictionary<string, string> Roles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public readonly HashSet<string> Skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        StoreMap MapStore(object store)
        {
            var m = new StoreMap();
            int[] roleIds = { 6, 5, 16, 3, 23, 4 };
            string[] roleNames = { MailAktarici.Roles.Inbox, MailAktarici.Roles.Sent, MailAktarici.Roles.Drafts,
                MailAktarici.Roles.Trash, MailAktarici.Roles.Junk, MailAktarici.Roles.Outbox };
            for (int i = 0; i < roleIds.Length; i++)
            {
                string id = DefaultFolderId(store, roleIds[i]);
                if (id != null && !m.Roles.ContainsKey(id)) m.Roles[id] = roleNames[i];
            }
            // Eşitleme Sorunları ve alt klasörleri: Outlook'un kendi kayıtları, mail değil.
            foreach (int k in new[] { 20, 19, 21, 22 })
            {
                string id = DefaultFolderId(store, k);
                if (id != null) m.Skip.Add(id);
            }
            return m;
        }

        static string DefaultFolderId(object store, int kind)
        {
            try
            {
                object f = Com.Call(store, "GetDefaultFolder", kind);
                string id = Com.Str(f, "EntryID");
                Com.Release(f);
                return id;
            }
            catch { return null; }
        }

        bool IsSkipped(object folder, StoreMap map)
        {
            string id = Com.Str(folder, "EntryID");
            if (id != null && map.Skip.Contains(id)) return true;
            object hidden = Com.Prop(folder, TagHidden);
            return hidden is bool && (bool)hidden;
        }

        string RoleOf(object folder, StoreMap map, bool topLevel, string name)
        {
            string id = Com.Str(folder, "EntryID");
            string role;
            if (id != null && map.Roles.TryGetValue(id, out role)) return role;
            if (!topLevel) return "";
            // "Gönderilmiş Öğeler (Yalnızca bu bilgisayar)" gibi yerel kopyalar da tanınsın
            string clean = Regex.Replace(name ?? "", @"\s*\([^)]*\)\s*$", "");
            return Roles.FromName(clean);
        }

        delegate void FolderVisitor(object folder, string logical, string role, int itemType);

        void Walk(object parent, string parentPath, StoreMap map, FolderVisitor visit)
        {
            object folders = Com.Get(parent, "Folders");
            int n = Com.Int(folders, "Count", 0);
            for (int i = 1; i <= n; i++)
            {
                r.Check();
                object f = null;
                try { f = Com.Item(folders, i); }
                catch (Exception ex) { r.Warn("Klasör okunamadı (" + parentPath + " #" + i + "): " + Com.Message(ex)); continue; }
                try
                {
                    if (IsSkipped(f, map)) continue;
                    string name = Com.Str(f, "Name") ?? "_";
                    string logical = parentPath == null ? name.Replace('\\', '_') : parentPath + "\\" + name.Replace('\\', '_');
                    string role = RoleOf(f, map, parentPath == null, name);
                    visit(f, logical, role, Com.Int(f, "DefaultItemType", OlMail));
                    Walk(f, logical, map, visit);
                }
                finally { Com.Release(f); }
            }
            Com.Release(folders);
        }

        static bool IsMailFolder(int type) { return type == OlMail || type == OlPost; }

        public void Count(OutlookSource src)
        {
            long items = 0, bytes = 0;
            object root = Com.Call(src.Store, "GetRootFolder");
            var map = MapStore(src.Store);
            Walk(root, null, map, (f, logical, role, type) =>
            {
                if (!IsMailFolder(type)) return;
                object it = Com.Get(f, "Items");
                items += Com.Int(it, "Count", 0);
                Com.Release(it);
                bytes += FolderBytes(f);
            });
            Com.Release(root);
            src.Items = items;
            src.Bytes = bytes;
        }

        // Klasördeki maillerin toplam boyutu. IMAP klasörlerinde klasör düzeyindeki boyut özelliği
        // 0 döndüğü için öğe tablosu tek çağrıda 5000 satır okunarak toplanır.
        static long FolderBytes(object folder)
        {
            object table = null, cols = null;
            long sum = 0;
            try
            {
                table = Com.Call(folder, "GetTable", Type.Missing, 0);
                cols = Com.Get(table, "Columns");
                Com.Call(cols, "RemoveAll");
                Com.Call(cols, "Add", TagSize);
                while (!Com.Bool(table, "EndOfTable"))
                {
                    var a = Com.Call(table, "GetArray", 5000) as Array;
                    if (a == null || a.Rank != 2 || a.GetLength(0) == 0) break;
                    int lo = a.GetLowerBound(0), c = a.GetLowerBound(1);
                    for (int i = lo; i <= a.GetUpperBound(0); i++)
                    {
                        object v = a.GetValue(i, c);
                        if (v is int || v is long) sum += Convert.ToInt64(v);
                    }
                }
            }
            catch { }
            finally
            {
                Com.Release(cols);
                Com.Release(table);
            }
            return sum;
        }

        // ---------------- Dışa aktarma ----------------

        public void ExportSource(OutlookSource src)
        {
            r.Info("== " + src.Name + " (" + src.Kind + ") aktarılıyor, " + src.Items.ToString("N0") + " mail");
            bool isPst = (src.FilePath ?? "").EndsWith(".pst", StringComparison.OrdinalIgnoreCase);
            AccountSink sink = pkg.OpenAccount(src.Name, "Outlook · " + src.Kind + (isPst && !src.Kind.Contains(src.FilePath) ? " · " + src.FilePath : ""));
            object root = Com.Call(src.Store, "GetRootFolder");
            var map = MapStore(src.Store);
            var skippedTypes = new List<string>();
            Walk(root, null, map, (f, logical, role, type) =>
            {
                if (IsMailFolder(type)) ExportFolder(f, logical, role, sink);
                else skippedTypes.Add(logical);
            });
            Com.Release(root);
            if (skippedTypes.Count > 0)
                r.Info("   Mail dışı klasörler (takvim, kişiler, görevler) EML'e alınmadı" +
                    (opt.SinglePst ? ", tek PST dosyasında bulunur: " : ": ") + string.Join(", ", skippedTypes.Take(8).ToArray()) +
                    (skippedTypes.Count > 8 ? " ..." : ""));
            if (src.Kind.StartsWith("Exchange") && sink.Oldest.HasValue && (DateTime.UtcNow - sink.Oldest.Value).TotalDays < 400)
                sink.Warnings.Add("En eski mail " + sink.Oldest.Value.ToLocalTime().ToString("yyyy-MM-dd") + " tarihli. Outlook 'çevrimdışı tutulacak posta' " +
                    "ayarı 1 yıl olabilir; daha eski mailler sunucuda kalmış olabilir. Hesap Ayarları > Değiştir > kaydırıcıyı 'Tümü' yapıp " +
                    "eşitleme bitince tekrar çalıştırın.");
        }

        void ExportFolder(object folder, string logical, string role, AccountSink sink)
        {
            object items = Com.Get(folder, "Items");
            long count = Com.Int(items, "Count", 0);
            long exported = 0, failed = 0, seen = 0;
            r.Status = sink.Name + " › " + logical;
            string dir;
            try { dir = sink.FolderDir(logical); }
            catch (Exception ex)
            {
                pkg.CheckTarget(ex);
                throw;
            }
            object it = null;
            try { if (count > 0) it = Com.Call(items, "GetFirst"); }
            catch (Exception ex) { r.Error(logical + ": klasör açılamadı: " + Com.Message(ex)); }
            while (it != null)
            {
                r.Check();
                seen++;
                try { ExportItem(it, sink, dir, logical, role, ref exported, ref failed); }
                catch (OperationCanceledException) { throw; }
                catch (TargetLostException) { throw; }
                catch (Exception ex)
                {
                    failed++;
                    r.Error(logical + ": bir öğe aktarılamadı: " + Com.Message(ex));
                    sink.AddMessage(new ManifestEntry { Folder = logical, Role = role, Status = "HATA", Subject = Com.Str(it, "Subject") });
                }
                finally { Com.Release(it); }
                r.Step();
                if (opt.LimitPerFolder > 0 && seen >= opt.LimitPerFolder) { count = seen; break; }
                long t = Prof.Start();
                try { it = Com.Call(items, "GetNext"); }
                catch (Exception ex) { r.Error(logical + ": sonraki öğeye geçilemedi: " + Com.Message(ex)); it = null; }
                Prof.Stop("getnext", t);
            }
            Com.Release(items);
            if (seen < count)
                r.Warn(logical + ": klasörde " + count + " öğe görünüyordu, " + seen + " tanesi okunabildi.");
            sink.AddFolder(logical, role, Math.Max(count, seen), exported, failed);
            if (count > 0) r.Info("   " + logical + ": " + exported + "/" + Math.Max(count, seen) + (failed > 0 ? ", " + failed + " hata" : ""));
        }

        ItemInfo ReadInfo(object item)
        {
            object[] v = Com.Props(item, ItemTags);
            var i = new ItemInfo
            {
                Headers = Com.S(v, 0), MessageId = Com.S(v, 1), MsgFlags = Com.I(v, 2), FlagStatus = Com.I(v, 3),
                LastVerb = Com.I(v, 4), Delivery = Com.D(v, 5), Submit = Com.D(v, 6), Created = Com.D(v, 7),
                Size = Com.L(v, 8), SenderSmtp = Com.S(v, 9), SenderName = Com.S(v, 10), SenderEmail = Com.S(v, 11),
                SenderType = Com.S(v, 12), Subject = Com.S(v, 13), MessageClass = Com.S(v, 14),
                InReplyTo = Com.S(v, 15), References = Com.S(v, 16)
            };
            if (i.Subject == null) i.Subject = Com.Str(item, "Subject");
            if (i.MessageClass == null) i.MessageClass = Com.Str(item, "MessageClass") ?? "IPM.Note";
            if (!i.Delivery.HasValue) i.Delivery = Com.LocalDate(item, "ReceivedTime");
            if (!i.Submit.HasValue) i.Submit = Com.LocalDate(item, "SentOn");
            if (!i.Created.HasValue) i.Created = Com.LocalDate(item, "CreationTime");
            if (i.Size == 0) i.Size = Com.Int(item, "Size", 0);
            return i;
        }

        static string FlagsOf(ItemInfo i)
        {
            var fl = new List<string>();
            if ((i.MsgFlags & 0x1) != 0) fl.Add("\\Seen");
            if (i.FlagStatus == 2) fl.Add("\\Flagged");
            if (i.LastVerb == 102 || i.LastVerb == 103) fl.Add("\\Answered");
            if (i.LastVerb == 104) fl.Add("$Forwarded");
            if ((i.MsgFlags & 0x8) != 0) fl.Add("\\Draft");
            return string.Join(" ", fl.ToArray());
        }

        string SenderAddress(object item, ItemInfo i)
        {
            if (!string.IsNullOrEmpty(i.SenderSmtp)) return i.SenderSmtp;
            if (!string.Equals(i.SenderType, "EX", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(i.SenderEmail) && i.SenderEmail.Contains("@"))
                return i.SenderEmail;
            object sender = Com.Get(item, "Sender");
            if (sender != null)
            {
                try
                {
                    object xu = Com.Call(sender, "GetExchangeUser");
                    string smtp = Com.Str(xu, "PrimarySmtpAddress");
                    Com.Release(xu);
                    if (!string.IsNullOrEmpty(smtp)) return smtp;
                }
                catch { }
                finally { Com.Release(sender); }
            }
            object acc = Com.Get(item, "SendUsingAccount");
            string a = Com.Str(acc, "SmtpAddress");
            Com.Release(acc);
            return !string.IsNullOrEmpty(a) ? a : i.SenderEmail;
        }

        void ExportItem(object item, AccountSink sink, string dir, string logical, string role, ref long exported, ref long failed)
        {
            long t = Prof.Start();
            ItemInfo info = ReadInfo(item);
            if (string.IsNullOrEmpty(info.MessageId))
                info.MessageId = Headers.Get(Headers.Parse(info.Headers), "Message-ID") ?? MessageId(item, info);
            var e = new ManifestEntry
            {
                Folder = logical, Role = role, Subject = info.Subject, DateUtc = info.BestDate,
                MessageId = info.MessageId, Size = info.Size, Flags = FlagsOf(info)
            };
            Prof.Stop("ozellikler", t);
            t = Prof.Start();
            string senderAddr = null;
            try { senderAddr = SenderAddress(item, info); }
            catch { }
            e.From = string.IsNullOrEmpty(info.SenderName) ? senderAddr : (senderAddr != null && senderAddr != info.SenderName ? info.SenderName + " <" + senderAddr + ">" : info.SenderName);
            Prof.Stop("gonderen", t);
            t = Prof.Start();
            object atts = Com.Get(item, "Attachments");
            e.Attachments = Com.Int(atts, "Count", 0);
            Com.Release(atts);
            bool headerOnly = Com.Int(item, "DownloadState", 1) == 0;
            Prof.Stop("ek-sayisi", t);

            string basePath = sink.ReserveBase(dir, e.DateUtc, info.Subject);
            string emlPath = basePath + ".eml", msgPath = basePath + ".msg";
            string err = null;
            currentSubject = info.Subject;
            if (opt.Eml)
            {
                if (pkg.Resume && Package.LooksComplete(emlPath)) pkg.Reused++;
                else
                {
                    try
                    {
                        byte[] eml = BuildEml(item, info, senderAddr, 0);
                        t = Prof.Start();
                        Package.WriteAtomic(emlPath, eml);
                        Stamp(emlPath, e.DateUtc);
                        Prof.Stop("dosya-yazma", t);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (TargetLostException) { throw; }
                    catch (Exception ex)
                    {
                        pkg.CheckTarget(ex);
                        err = "EML: " + Com.Message(ex);
                    }
                }
            }
            if (opt.Msg)
            {
                if (pkg.Resume && Package.LooksComplete(msgPath)) pkg.Reused++;
                else
                {
                    t = Prof.Start();
                    string tmp = Path.Combine(dir, "~" + Guid.NewGuid().ToString("N").Substring(0, 12) + ".msg");
                    try
                    {
                        Com.Call(item, "SaveAs", tmp, 9);   // olMSGUnicode
                        if (File.Exists(msgPath)) File.Delete(msgPath);
                        File.Move(tmp, msgPath);
                        Stamp(msgPath, e.DateUtc);
                    }
                    catch (Exception ex)
                    {
                        try { if (File.Exists(tmp)) File.Delete(tmp); }
                        catch { }
                        pkg.CheckTarget(ex);
                        err = (err == null ? "" : err + " | ") + "MSG: " + Com.Message(ex);
                    }
                    Prof.Stop("msg", t);
                }
            }
            // Doğrulama için: bu maile ait diskte gerçekten duran dosyalar
            bool emlOk = File.Exists(emlPath), msgOk = File.Exists(msgPath);
            e.FilesWritten = (emlOk ? 1 : 0) + (msgOk ? 1 : 0);
            e.File = emlOk ? sink.Relative(emlPath) : msgOk ? sink.Relative(msgPath) : null;
            if (e.FilesWritten == 0)
            {
                failed++;
                e.Status = "HATA";
                r.Error(logical + " › " + (info.Subject ?? "(konu yok)") + ": " + err);
            }
            else
            {
                exported++;
                e.Status = headerOnly ? "SADECE_BASLIK" : (err != null ? "KISMEN" : "TAM");
                if (err != null) r.Warn(logical + " › " + (info.Subject ?? "(konu yok)") + ": " + err);
            }
            if (headerOnly) sink.HeaderOnly++;
            sink.AddMessage(e);
        }

        static void Stamp(string path, DateTime? utc)
        {
            if (!utc.HasValue) return;
            try
            {
                File.SetLastWriteTimeUtc(path, utc.Value);
                File.SetCreationTimeUtc(path, utc.Value);
            }
            catch { }
        }

        // ---------------- EML üretimi ----------------

        byte[] BuildEml(object item, ItemInfo info, string senderAddr, int depth)
        {
            var top = new List<string>();
            List<HeaderField> orig = Headers.Parse(info.Headers);
            bool useOrig = orig.Count >= 3 && (Headers.Has(orig, "From") || Headers.Has(orig, "Received"));
            if (useOrig)
            {
                foreach (var f in orig)
                    if (!DropHeaders.Contains(f.Name)) top.Add(f.Raw);
                if (!Headers.Has(orig, "Date") && info.BestDate.HasValue) top.Add("Date: " + Util.Rfc2822Date(info.BestDate.Value));
                if (!Headers.Has(orig, "From")) AddFrom(top, info, senderAddr);
                if (!Headers.Has(orig, "Subject")) top.Add("Subject: " + Rfc2047.Encode(info.Subject));
                if (!Headers.Has(orig, "Message-ID")) top.Add("Message-ID: " + MessageId(item, info));
            }
            else
            {
                DateTime? date = info.Submit ?? info.Delivery ?? info.Created;
                if (date.HasValue) top.Add("Date: " + Util.Rfc2822Date(date.Value));
                AddFrom(top, info, senderAddr);
                AddRecipients(top, item);
                top.Add("Subject: " + Rfc2047.Encode(info.Subject));
                top.Add("Message-ID: " + MessageId(item, info));
                if (!string.IsNullOrEmpty(info.InReplyTo)) top.Add("In-Reply-To: " + info.InReplyTo);
                if (!string.IsNullOrEmpty(info.References)) top.Add("References: " + info.References);
            }

            long t = Prof.Start();
            int fmt = Com.Int(item, "BodyFormat", 1);
            Prof.Stop("bodyformat", t);
            t = Prof.Start();
            string text = Com.Str(item, "Body");
            Prof.Stop("body", t);
            t = Prof.Start();
            string html = (fmt == 2 || fmt == 3) ? Com.Str(item, "HTMLBody") : null;
            Prof.Stop("htmlbody", t);
            t = Prof.Start();
            Entity body;
            if (!string.IsNullOrEmpty(html))
            {
                html = Regex.Replace(html, @"(<meta[^>]*charset\s*=\s*[""']?)([\w-]+)", "$1utf-8", RegexOptions.IgnoreCase);
                body = Entity.Multipart("alternative", Entity.Text("plain", text), Entity.Text("html", html));
            }
            else body = Entity.Text("plain", text);

            var inline = new List<Entity>();
            var regular = new List<Entity>();
            object atts = Com.Get(item, "Attachments");
            int n = Com.Int(atts, "Count", 0);
            for (int i = 1; i <= n; i++)
            {
                object a = null;
                try
                {
                    a = Com.Item(atts, i);
                    bool isInline;
                    Entity part = BuildAttachment(a, html, depth, out isInline);
                    if (part != null) (isInline ? inline : regular).Add(part);
                }
                catch (Exception ex) { r.Warn((info.Subject ?? "(konu yok)") + ": ek #" + i + " alınamadı: " + Com.Message(ex)); }
                finally { Com.Release(a); }
            }
            Com.Release(atts);
            Prof.Stop("ekler", t);

            if ((info.MessageClass ?? "").StartsWith("IPM.Schedule.Meeting", StringComparison.OrdinalIgnoreCase))
            {
                string ics = Path.Combine(temp, Guid.NewGuid().ToString("N") + ".ics");
                try
                {
                    Com.Call(item, "SaveAs", ics, 8);   // olICal
                    var part = new Entity { ContentType = "text/calendar; charset=utf-8; name=\"davet.ics\"", Data = File.ReadAllBytes(ics) };
                    part.ExtraHeaders.Add("Content-Disposition: attachment; filename=\"davet.ics\"");
                    regular.Add(part);
                }
                catch { }
                finally { TryDelete(ics); }
            }

            if (inline.Count > 0)
            {
                var rel = Entity.Multipart("related", body);
                rel.Parts.AddRange(inline);
                body = rel;
            }
            if (regular.Count > 0)
            {
                var mixed = Entity.Multipart("mixed", body);
                mixed.Parts.AddRange(regular);
                body = mixed;
            }
            return MimeWriter.Build(top, body);
        }

        void AddFrom(List<string> top, ItemInfo info, string senderAddr)
        {
            string from = Rfc2047.Address(info.SenderName, senderAddr);
            if (from != null) top.Add("From: " + from);
        }

        string MessageId(object item, ItemInfo info)
        {
            if (!string.IsNullOrEmpty(info.MessageId)) return info.MessageId;
            // Kalıcı kimlik: aynı mail tekrar aktarılırsa aynı Message-ID üretilsin.
            string entry = Com.Str(item, "EntryID") ?? Guid.NewGuid().ToString();
            return "<" + Util.HexHash(Encoding.ASCII.GetBytes(entry)).Substring(0, 24) + "@mailaktarici.local>";
        }

        void AddRecipients(List<string> top, object item)
        {
            var to = new List<string>();
            var cc = new List<string>();
            var bcc = new List<string>();
            object recips = Com.Get(item, "Recipients");
            int n = Com.Int(recips, "Count", 0);
            for (int i = 1; i <= n; i++)
            {
                object rc = null;
                try
                {
                    rc = Com.Item(recips, i);
                    string name = Com.Str(rc, "Name");
                    string addr = Com.Str(rc, "Address");
                    if (string.IsNullOrEmpty(addr) || !addr.Contains("@"))
                    {
                        string smtp = Com.Prop(rc, TagSmtp) as string;
                        if (string.IsNullOrEmpty(smtp))
                        {
                            object ae = Com.Get(rc, "AddressEntry");
                            try
                            {
                                object xu = Com.Call(ae, "GetExchangeUser");
                                smtp = Com.Str(xu, "PrimarySmtpAddress");
                                Com.Release(xu);
                            }
                            catch { }
                            Com.Release(ae);
                        }
                        if (!string.IsNullOrEmpty(smtp)) addr = smtp;
                    }
                    string formatted = Rfc2047.Address(name, addr);
                    if (formatted == null) continue;
                    int type = Com.Int(rc, "Type", 1);
                    (type == 2 ? cc : type == 3 ? bcc : to).Add(formatted);
                }
                catch { }
                finally { Com.Release(rc); }
            }
            Com.Release(recips);
            if (to.Count > 0) top.Add("To: " + string.Join(",\r\n ", to.ToArray()));
            if (cc.Count > 0) top.Add("Cc: " + string.Join(",\r\n ", cc.ToArray()));
            if (bcc.Count > 0) top.Add("Bcc: " + string.Join(",\r\n ", bcc.ToArray()));
        }

        Entity BuildAttachment(object a, string html, int depth, out bool isInline)
        {
            isInline = false;
            int type = Com.Int(a, "Type", 1);
            if (type == 4) return null;   // olByReference: yalnız bağlantı, dosyanın kendisi mailde yok
            string fileName = Com.Str(a, "FileName");
            if (string.IsNullOrEmpty(fileName)) fileName = Com.Str(a, "DisplayName");
            if (string.IsNullOrEmpty(fileName)) fileName = type == 5 ? "ekli_mail.msg" : "ek";
            object[] ap = Com.Props(a, AttTags);
            // Geçici dosya adı yalnız harf/rakam: OpenSharedItem "..." ya da "%" içeren yolu URL sanıp reddediyor.
            string tmp = Path.Combine(temp, Guid.NewGuid().ToString("N") + (type == 5 ? ".msg" : ".bin"));
            Com.Call(a, "SaveAsFile", tmp);
            try
            {
                if (type == 5 && depth < 3)
                {
                    // Ekli mail: kendi başına EML'e çevrilip message/rfc822 olarak gömülür.
                    object shared = null;
                    try
                    {
                        shared = Com.Call(ns, "OpenSharedItem", tmp);
                        ItemInfo sub = ReadInfo(shared);
                        string subSender = null;
                        try { subSender = SenderAddress(shared, sub); }
                        catch { }
                        byte[] eml = BuildEml(shared, sub, subSender, depth + 1);
                        bool eight = eml.Any(b => b > 127);
                        var part = new Entity { ContentType = "message/rfc822", Data = eml, TransferEncoding = eight ? "8bit" : "7bit" };
                        string emlName = Util.SafeName(string.IsNullOrEmpty(sub.Subject) ? "ekli mail" : sub.Subject, 60) + ".eml";
                        part.ExtraHeaders.Add("Content-Disposition: attachment" + Rfc2047.NameParams("filename", "filename", emlName));
                        return part;
                    }
                    catch (Exception ex)
                    {
                        r.Warn((currentSubject ?? "(konu yok)") + ": ekli mail EML'e çevrilemedi, .msg olarak eklendi: " + Com.Message(ex));
                    }
                    finally { Com.Release(shared); }
                }
                byte[] data = File.ReadAllBytes(tmp);
                string mimeTag = Com.S(ap, 0);
                string ctype = !string.IsNullOrEmpty(mimeTag) && mimeTag.Contains("/") && !mimeTag.Contains(" ") ? mimeTag.Trim() : MimeWriter.GuessType(fileName);
                string cid = (Com.S(ap, 1) ?? "").Trim().Trim('<', '>');
                isInline = html != null && cid.Length > 0 && html.IndexOf("cid:" + cid, StringComparison.OrdinalIgnoreCase) >= 0;
                string nameForMail = type == 5 && !fileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) ? fileName + ".msg" : fileName;
                // message/* ve multipart/* türleri base64 ile kodlanamaz (RFC 2046). Metin olarak
                // güvenliyse olduğu gibi (7bit/8bit) konur, değilse genel ikili tür olarak eklenir.
                string transfer = "base64";
                string lowerType = ctype.ToLowerInvariant();
                if (lowerType.StartsWith("multipart/")) ctype = "application/octet-stream";
                else if (lowerType.StartsWith("message/"))
                {
                    byte[] safe = MimeWriter.AsSafeText(data);
                    if (safe != null)
                    {
                        data = safe;
                        transfer = data.Any(b => b > 127) ? "8bit" : "7bit";
                    }
                    else ctype = "application/octet-stream";
                }
                var ent = new Entity { ContentType = ctype + Rfc2047.NameParams("name", null, nameForMail), Data = data, TransferEncoding = transfer };
                ent.ExtraHeaders.Add("Content-Disposition: " + (isInline ? "inline" : "attachment") + Rfc2047.NameParams("filename", "filename", nameForMail));
                if (cid.Length > 0) ent.ExtraHeaders.Add("Content-ID: <" + cid + ">");
                return ent;
            }
            finally { TryDelete(tmp); }
        }

        static void TryDelete(string p)
        {
            try { if (File.Exists(p)) File.Delete(p); }
            catch { }
        }

        // ---------------- Tek PST dosyası ----------------

        const long PstSplitBytes = 40L * 1024 * 1024 * 1024;

        // Unicode PST'nin varsayılan üst sınırı 50 GB. Toplam 40 GB'ı aşıyorsa her hesap ayrı PST'ye yazılır.
        public void BuildPsts(List<OutlookSource> sources, string root)
        {
            var withMail = sources.Where(s => s.Items > 0 || s.Bytes > 0).ToList();
            if (withMail.Count == 0) return;
            long total = withMail.Sum(s => s.Bytes);
            if (total <= PstSplitBytes)
            {
                BuildSinglePst(withMail, Path.Combine(root, "TumMailler.pst"));
                return;
            }
            r.Info("Toplam veri " + Util.FormatSize(total) + ": PST'nin 50 GB sınırı yüzünden her hesap ayrı PST dosyasına yazılacak.");
            foreach (var s in withMail)
            {
                r.Check();
                if (s.Bytes > 45L * 1024 * 1024 * 1024)
                    r.Warn(s.Name + " tek başına " + Util.FormatSize(s.Bytes) + ": PST sınırını aşabilir. Bu hesap zaten bir PST dosyasıysa, Outlook kapalıyken özgün dosyayı kopyalamak daha doğru: " + s.FilePath);
                try { BuildSinglePst(new List<OutlookSource> { s }, Path.Combine(root, "TumMailler_" + Util.SafeName(s.Name, 40) + ".pst")); }
                catch (OperationCanceledException) { throw; }
                catch (TargetLostException) { throw; }
                catch (Exception ex)
                {
                    pkg.CheckTarget(ex);
                    r.Error("PST oluşturulamadı (" + s.Name + "): " + Com.Message(ex));
                }
            }
        }

        public void BuildSinglePst(List<OutlookSource> sources, string pstPath)
        {
            r.Info("== PST dosyası hazırlanıyor: " + pstPath);
            r.Status = "PST dosyası hazırlanıyor: " + Path.GetFileName(pstPath);
            if (File.Exists(pstPath))
            {
                // Yarım kalmış önceki çalışmadan: profilde bağlı kaldıysa çıkar, sonra sil (yoksa kopyalar iki kez girer)
                object old = FindStoreByPath(pstPath);
                if (old != null)
                {
                    try
                    {
                        object oldRoot = Com.Call(old, "GetRootFolder");
                        Com.Call(ns, "RemoveStore", oldRoot);
                        Com.Release(oldRoot);
                    }
                    catch { }
                    Com.Release(old);
                }
                File.Delete(pstPath);
                r.Info("   Önceki yarım PST silindi, baştan oluşturuluyor.");
            }
            Com.Call(ns, "AddStoreEx", pstPath, 2);   // olStoreUnicode
            object dst = FindStoreByPath(pstPath);
            if (dst == null) throw new ApplicationException("Yeni PST dosyası Outlook'ta açılamadı.");
            object root = Com.Call(dst, "GetRootFolder");
            try { Com.Set(root, "Name", "Aktarılan Mailler " + Environment.MachineName + " " + DateTime.Now.ToString("yyyy-MM-dd")); }
            catch { }
            long srcTotal = 0, dstTotal = 0;
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (var src in sources)
                {
                    r.Check();
                    string fname = Util.SafeName(src.Name, 60);
                    for (int k = 2; !used.Add(fname); k++) fname = Util.SafeName(src.Name, 55) + " " + k;
                    object rootFolders = Com.Get(root, "Folders");
                    object accFolder = Com.Call(rootFolders, "Add", fname);
                    Com.Release(rootFolders);
                    object sroot = Com.Call(src.Store, "GetRootFolder");
                    var map = MapStore(src.Store);
                    object tops = Com.Get(sroot, "Folders");
                    int n = Com.Int(tops, "Count", 0);
                    for (int i = 1; i <= n; i++)
                    {
                        r.Check();
                        object f = Com.Item(tops, i);
                        try
                        {
                            if (IsSkipped(f, map)) continue;
                            string name = Com.Str(f, "Name");
                            long sc = CountAll(f);
                            r.Status = "PST'ye kopyalanıyor: " + src.Name + " › " + name + " (" + sc.ToString("N0") + " öğe)";
                            object copy = null;
                            try { copy = Com.Call(f, "CopyTo", accFolder); }
                            catch (Exception ex)
                            {
                                pkg.CheckTarget(ex);
                                r.Warn(src.Name + " › " + name + ": klasör toptan kopyalanamadı (" + Com.Message(ex) + "), öğe öğe kopyalanıyor.");
                                copy = CopyItemwise(f, accFolder);
                            }
                            long dc = copy != null ? CountAll(copy) : 0;
                            Com.Release(copy);
                            srcTotal += sc;
                            dstTotal += dc;
                            if (dc != sc) r.Warn("PST: " + src.Name + " › " + name + " kaynakta " + sc + ", PST'de " + dc + " öğe.");
                        }
                        finally { Com.Release(f); }
                    }
                    Com.Release(tops);
                    Com.Release(sroot);
                    Com.Release(accFolder);
                    r.Info("   PST'ye eklendi: " + src.Name);
                }
            }
            finally
            {
                try { Com.Call(ns, "RemoveStore", root); }
                catch (Exception ex) { r.Warn("Yeni PST Outlook'tan çıkarılamadı: " + Com.Message(ex)); }
                Com.Release(root);
                Com.Release(dst);
            }
            long size = 0;
            try { size = new FileInfo(pstPath).Length; }
            catch { }
            string line = "TEK PST: " + Path.GetFileName(pstPath) + " (" + Util.FormatSize(size) + "): kaynakta " + srcTotal.ToString("N0") +
                " öğe, PST'de " + dstTotal.ToString("N0") + " öğe (takvim/kişiler dahil)  ->  " + (srcTotal == dstTotal ? "TUTUYOR" : "TUTMUYOR, günlüğe bakın");
            pkg.Notes.Add(line);
            r.Info(line);
        }

        long CountAll(object folder)
        {
            object items = Com.Get(folder, "Items");
            long n = Com.Int(items, "Count", 0);
            Com.Release(items);
            object subs = Com.Get(folder, "Folders");
            int c = Com.Int(subs, "Count", 0);
            for (int i = 1; i <= c; i++)
            {
                object s = Com.Item(subs, i);
                n += CountAll(s);
                Com.Release(s);
            }
            Com.Release(subs);
            return n;
        }

        static int FolderTypeFor(int itemType)
        {
            switch (itemType)
            {
                case 1: return 9;    // takvim
                case 2: return 10;   // kişiler
                case 3: return 13;   // görevler
                case 4: return 11;   // günlük
                case 5: return 12;   // notlar
                default: return 6;   // mail
            }
        }

        // CopyTo'nun reddettiği klasörler için: her öğe .msg'ye yazılıp PST'ye taşınır.
        // Kaynak klasöre hiçbir şey yazılmaz (Copy() IMAP'te sunucuya kopya bırakırdı).
        object CopyItemwise(object src, object dstParent)
        {
            string name = Com.Str(src, "Name") ?? "Klasör";
            object dstFolders = Com.Get(dstParent, "Folders");
            object dst = null;
            int type = FolderTypeFor(Com.Int(src, "DefaultItemType", 0));
            for (int k = 1; dst == null && k < 50; k++)
            {
                try { dst = Com.Call(dstFolders, "Add", k == 1 ? name : name + " " + k, type); }
                catch { }
            }
            Com.Release(dstFolders);
            if (dst == null) throw new ApplicationException("PST'de klasör oluşturulamadı: " + name);

            object items = Com.Get(src, "Items");
            object it = Com.Int(items, "Count", 0) > 0 ? Com.Call(items, "GetFirst") : null;
            while (it != null)
            {
                r.Check();
                string tmp = Path.Combine(temp, Guid.NewGuid().ToString("N") + ".msg");
                object shared = null, moved = null;
                try
                {
                    Com.Call(it, "SaveAs", tmp, 9);
                    shared = Com.Call(ns, "OpenSharedItem", tmp);
                    moved = Com.Call(shared, "Move", dst);
                }
                catch (Exception ex) { r.Warn("PST: öğe kopyalanamadı (" + name + "): " + Com.Message(ex)); }
                finally
                {
                    Com.Release(moved);
                    Com.Release(shared);
                    Com.Release(it);
                    TryDelete(tmp);
                }
                it = Com.Call(items, "GetNext");
            }
            Com.Release(items);

            object subs = Com.Get(src, "Folders");
            int c = Com.Int(subs, "Count", 0);
            for (int i = 1; i <= c; i++)
            {
                object s = Com.Item(subs, i);
                try { Com.Release(CopyItemwise(s, dst)); }
                catch (Exception ex) { r.Warn("PST: alt klasör kopyalanamadı: " + Com.Message(ex)); }
                finally { Com.Release(s); }
            }
            Com.Release(subs);
            return dst;
        }
    }

    // Disklerde PST araması
    static class PstFinder
    {
        static readonly HashSet<string> SkipAnywhere = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "$Recycle.Bin", "System Volume Information", "$WINDOWS.~BT", "$WinREAgent", "WinSxS", "node_modules", ".git"
        };
        static readonly HashSet<string> SkipAtRoot = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Windows", "Program Files", "Program Files (x86)", "ProgramData", "Recovery", "PerfLogs"
        };

        public static List<string> DefaultRoots()
        {
            var roots = new List<string>();
            foreach (var d in DriveInfo.GetDrives())
            {
                try { if (d.DriveType == DriveType.Fixed && d.IsReady) roots.Add(d.RootDirectory.FullName); }
                catch { }
            }
            return roots;
        }

        // Bu programın daha önce ürettiği TumMailler.pst tekrar kaynak olarak alınmasın.
        public static bool IsOwnOutput(string path)
        {
            try
            {
                string dir = Path.GetDirectoryName(path);
                return File.Exists(Path.Combine(dir, "ozet_rapor.txt")) || File.Exists(Path.Combine(dir, "gunluk.txt"));
            }
            catch { return false; }
        }

        public static List<string> Find(IEnumerable<string> roots, Reporter r, List<string> ostFound)
        {
            var result = new List<string>();
            foreach (string root in roots)
            {
                var stack = new Stack<string>();
                stack.Push(root);
                while (stack.Count > 0)
                {
                    r.Check();
                    string dir = stack.Pop();
                    r.Status = "PST aranıyor: " + dir;
                    try
                    {
                        foreach (string f in Directory.EnumerateFiles(dir, "*.?st"))
                        {
                            if (f.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)) result.Add(f);
                            else if (ostFound != null && f.EndsWith(".ost", StringComparison.OrdinalIgnoreCase)) ostFound.Add(f);
                        }
                    }
                    catch { }
                    bool isRoot = Path.GetPathRoot(dir) == dir;
                    try
                    {
                        foreach (string sub in Directory.EnumerateDirectories(dir))
                        {
                            string name = Path.GetFileName(sub);
                            if (SkipAnywhere.Contains(name) || (isRoot && SkipAtRoot.Contains(name))) continue;
                            try
                            {
                                if ((File.GetAttributes(sub) & FileAttributes.ReparsePoint) != 0) continue;
                            }
                            catch { continue; }
                            stack.Push(sub);
                        }
                    }
                    catch { }
                }
            }
            return result;
        }
    }
}
