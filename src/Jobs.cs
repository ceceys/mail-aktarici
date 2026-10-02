using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailAktarici
{
    sealed class ExportJob
    {
        public string TargetParent;
        public string ContinueDir;               // dolu ise bu yarım aktarım klasörüne devam edilir
        public bool Outlook = true;
        public List<string> StoreIds;            // null: uygun olan tüm hesaplar
        public List<string> StoreNameFilter;     // komut satırı: adı bunlardan birini içeren hesaplar
        public bool ScanDisks;
        public List<string> ScanRoots;           // null: tüm sabit diskler
        public List<string> PstFiles = new List<string>();
        public ServerConfig Server;
        public ExportOptions Opt = new ExportOptions();
        public Func<string, bool> Confirm;       // arayüz: evet/hayır sorusu (null: sormadan devam)
    }

    static class Jobs
    {
        // Dönüş: aktarım klasörünün yolu
        public static string RunExport(ExportJob job, Reporter r)
        {
            r.ResetProgress();
            r.OpenLocal();
            CrashReport.Current = r;
            Package pkg = job.ContinueDir != null ? Package.Continue(job.ContinueDir, r) : Package.Create(job.TargetParent, r);
            r.Info("Mail Aktarıcı " + Program.Version + " · bilgisayar: " + Environment.MachineName + " · kullanıcı: " + Environment.UserName +
                " · Windows " + Environment.OSVersion.Version + (Environment.Is64BitProcess ? " (64 bit)" : " (32 bit)"));
            r.Info("Seçenekler: EML=" + (job.Opt.Eml ? "evet" : "hayır") + " MSG=" + (job.Opt.Msg ? "evet" : "hayır") + " PST=" + (job.Opt.SinglePst ? "evet" : "hayır"));
            string state = "TAMAMLANDI";
            try
            {
                bool needOutlook = job.Outlook || job.ScanDisks || job.PstFiles.Count > 0;
                if (needOutlook)
                {
                    try { RunOutlookPart(job, pkg, r); }
                    catch (OperationCanceledException) { throw; }
                    catch (TargetLostException) { throw; }
                    catch (Exception ex) { r.Error("Outlook bölümü tamamlanamadı: " + Com.Message(ex)); }
                }
                if (job.Server != null)
                {
                    try { ServerExporter.Run(job.Server, pkg, r, job.Opt.LimitPerFolder); }
                    catch (OperationCanceledException) { throw; }
                    catch (TargetLostException) { throw; }
                    catch (Exception ex) { r.Error("Sunucudan indirme tamamlanamadı: " + ex.Message); }
                }
            }
            catch (OperationCanceledException)
            {
                state = "KULLANICI DURDURDU (aktarım yarım kaldı; aynı yere tekrar başlatınca kaldığı yerden devam eder)";
                r.Warn("Aktarım durduruldu.");
            }
            catch (TargetLostException ex)
            {
                state = (ex.Message.Contains("DOLDU") ? "KAYIT DİSKİ DOLDU" : "KAYIT DİSKİ KOPTU") + " (aktarım yarım kaldı; aynı yere tekrar başlatınca kaldığı yerden devam eder)";
                r.Error(ex.Message + (ex.InnerException != null ? " [" + ex.InnerException.Message + "]" : ""));
            }
            if (state == "TAMAMLANDI" && r.Errors > 0) state = "HATALARLA TAMAMLANDI (" + r.Errors + " hata, gunluk.txt'ye bakın)";
            Prof.Report(r);
            r.Status = "Rapor yazılıyor";
            pkg.Finish(state);
            r.Status = state;
            return pkg.Root;
        }

        static void RunOutlookPart(ExportJob job, Package pkg, Reporter r)
        {
            MessageFilter.Register();
            try
            {
                using (var s = OutlookSession.Open(r))
                {
                    var ex = new OutlookExporter(s, pkg, r, job.Opt);
                    var sources = new List<OutlookSource>();
                    var profile = OutlookExporter.ListProfileStores(s.Ns);
                    var profilePaths = new HashSet<string>(profile.Where(p => !string.IsNullOrEmpty(p.FilePath))
                        .Select(p => OutlookSession.NormalizePath(p.FilePath)));
                    if (job.Outlook)
                    {
                        foreach (var p in profile)
                        {
                            bool take;
                            if (job.StoreIds != null) take = job.StoreIds.Contains(p.StoreId);
                            else if (job.StoreNameFilter != null && job.StoreNameFilter.Count > 0)
                                take = job.StoreNameFilter.Any(f => p.Name.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0);
                            else take = p.DefaultSelected;
                            if (take) sources.Add(p);
                            else r.Info("Seçilmedi, atlandı: " + p.Name + " (" + p.Kind + ")");
                        }
                    }

                    var pstPaths = new List<string>();
                    foreach (string p in job.PstFiles)
                        if (!profilePaths.Contains(OutlookSession.NormalizePath(p))) pstPaths.Add(p);
                    if (job.ScanDisks)
                    {
                        r.Info("Disklerde PST dosyası aranıyor...");
                        var osts = new List<string>();
                        var found = PstFinder.Find(job.ScanRoots ?? PstFinder.DefaultRoots(), r, osts);
                        string outRoot = OutlookSession.NormalizePath(pkg.Root);
                        foreach (string f in found)
                        {
                            string n = OutlookSession.NormalizePath(f);
                            if (profilePaths.Contains(n) || n.StartsWith(outRoot) || PstFinder.IsOwnOutput(f)) continue;
                            if (!pstPaths.Any(x => OutlookSession.NormalizePath(x) == n)) pstPaths.Add(f);
                        }
                        r.Info("   Profile bağlı olmayan " + pstPaths.Count + " PST bulundu.");
                        if (osts.Count > 0)
                            r.Info("   Diskte " + osts.Count + " OST dosyası var; OST yalnız kendi Outlook profiliyle okunur, profildeki hesaplar zaten aktarılıyor.");
                    }
                    foreach (string p in pstPaths)
                    {
                        r.Check();
                        try
                        {
                            var src = ex.AttachPst(p);
                            sources.Add(src);
                            r.Info("PST eklendi: " + p);
                        }
                        catch (Exception e) { r.Error("PST açılamadı: " + p + " · " + Com.Message(e)); }
                    }

                    try
                    {
                        r.Status = "Mailler sayılıyor...";
                        long bytes = 0;
                        foreach (var src in sources)
                        {
                            r.Check();
                            long pt = Prof.Start();
                            ex.Count(src);
                            Prof.Stop("sayim", pt);
                            long counted = job.Opt.LimitPerFolder > 0 ? Math.Min(src.Items, job.Opt.LimitPerFolder * 50L) : src.Items;
                            r.AddTotal(counted);
                            bytes += src.Bytes;
                            r.Info("Kaynak: " + src.Name + " · " + src.Kind + " · " + src.Items.ToString("N0") + " mail, " + Util.FormatSize(src.Bytes) +
                                (string.IsNullOrEmpty(src.FilePath) ? "" : " · " + src.FilePath));
                        }
                        CheckSpace(pkg, bytes, job, r);
                        foreach (var src in sources)
                        {
                            r.Check();
                            ex.ExportSource(src);
                            pkg.FlushAll();
                        }
                        if (job.Opt.SinglePst && sources.Count > 0)
                        {
                            try { ex.BuildPsts(sources, pkg.Root); }
                            catch (OperationCanceledException) { throw; }
                            catch (TargetLostException) { throw; }
                            catch (Exception e)
                            {
                                pkg.CheckTarget(e);
                                r.Error("PST dosyası oluşturulamadı: " + Com.Message(e));
                            }
                        }
                    }
                    finally
                    {
                        foreach (var src in sources) ex.Detach(src);
                        ex.Cleanup();
                    }
                }
            }
            finally
            {
                MessageFilter.Revoke();
            }
        }

        // Yer yetmeyecekse aktarım başlamadan sorulur (yarıda kalmasın).
        static void CheckSpace(Package pkg, long bytes, ExportJob job, Reporter r)
        {
            ExportOptions opt = job.Opt;
            long need, free;
            try
            {
                double factor = (opt.Eml ? 1.4 : 0) + (opt.Msg ? 1.2 : 0) + (opt.SinglePst ? 1.1 : 0);
                need = (long)(bytes * factor);
                if (pkg.Resume)
                {
                    long existing = 0;
                    try { existing = new DirectoryInfo(pkg.Root).EnumerateFiles("*.*", SearchOption.AllDirectories).Where(f => !f.Name.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)).Sum(f => f.Length); }
                    catch { }
                    need = Math.Max(0, need - existing);
                }
                free = new DriveInfo(Path.GetPathRoot(pkg.Root)).AvailableFreeSpace;
            }
            catch { return; }
            r.Info("Tahmini gereken yer: " + Util.FormatSize(need) + ", boş yer: " + Util.FormatSize(free));
            if (need <= free) return;
            string msg = "Seçilen biçimler için yaklaşık " + Util.FormatSize(need) + " yer gerekiyor, kayıt diskinde " + Util.FormatSize(free) + " boş.\r\n" +
                "Yer biterse aktarım yarıda kalır.\r\n\r\nÖneri: MSG kopyasını kapatın" +
                (opt.SinglePst ? "; kaynaklar zaten PST dosyasıysa 'Tek PST'yi de kapatıp özgün PST dosyalarını (Outlook kapalıyken) doğrudan kopyalayın" : "") +
                ".\r\n\r\nYine de başlatılsın mı?";
            r.Warn("Diskte yer yetmeyebilir! Gereken yaklaşık " + Util.FormatSize(need) + ", boş " + Util.FormatSize(free) + ".");
            if (job.Confirm != null && !job.Confirm(msg)) throw new OperationCanceledException();
        }
    }
}
