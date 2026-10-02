using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace MailAktarici
{
    // Tek ekran, kaydırmasız: solda kaynaklar, sağda kayıt yeri + biçim + günlük, altta sabit işlem çubuğu.
    sealed class MainForm : Form
    {
        readonly Reporter reporter = new Reporter();
        readonly ConcurrentQueue<string> logQueue = new ConcurrentQueue<string>();
        readonly System.Windows.Forms.Timer uiTimer = new System.Windows.Forms.Timer();
        readonly List<Label> wrapLabels = new List<Label>();
        Thread worker;
        bool closeAfterWork;
        bool storesLoaded, pstScanned;
        string lastPackage;

        CheckBox chkOutlook, chkPst, chkServer, chkEml, chkPstOut, chkMsg;
        CheckedListBox lstStores, lstPst;
        Button btnListStores, btnScan, btnAddPst, btnServer, btnBrowse, btnStart, btnStop, btnOpen;
        TextBox txtTarget, txtLog;
        Label lblStatus, lblCloud, lblStoreHint, lblPstHint, lblServerInfo;
        ProgressBar progress;
        Panel storesFrame, pstFrame;
        ServerConfig serverCfg;   // parola yalnız bellekte durur

        public MainForm()
        {
            Text = "Mail Aktarıcı " + Program.Version + " · by " + Program.Author;
            Font = Theme.Base;
            ForeColor = Theme.Text;
            BackColor = Theme.Bg;
            using (var g = CreateGraphics()) Theme.Scale = g.DpiX / 96f;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1100, 690);
            MinimumSize = new Size(1000, 680);
            Icon = Theme.AppIcon(32);
            BuildUi();
            reporter.Sink = line => logQueue.Enqueue(line);
            CrashReport.Current = reporter;
            uiTimer.Interval = 200;
            uiTimer.Tick += (s, e) => PumpUi();
            uiTimer.Start();
            txtTarget.Text = Util.DesktopPath();
            UpdateCloudWarning();
            AppendLog("Hazır. Kaynakları seçip 'Aktarmayı Başlat'a basın. Hiçbir mail silinmez veya değiştirilmez.");
            if (!OutlookSession.IsClassicOutlookInstalled())
            {
                chkOutlook.Checked = false;
                chkPst.Checked = false;
                Shown += (s, e) => chkServer.Checked = true;
                AppendLog("Bu bilgisayarda klasik Outlook yok. Mailleri sunucudan indirmek için 'Sunucudan doğrudan indir' bölümünü doldurun.");
            }
            else if (OutlookSession.ClassicProfileCount() == 0)
            {
                // Klasik Outlook kurulu ama hiç kullanılmamış: hesaplar yeni Outlook'ta.
                chkOutlook.Checked = false;
                chkPst.Checked = false;
                Shown += (s, e) => OfferServer("Bu bilgisayarda klasik Outlook'ta hiç hesap yok; hesaplar yeni Outlook'ta görünüyor.");
            }
            else if (OutlookSession.NewOutlookPreferred())
                AppendLog("Not: bu bilgisayarda 'Yeni Outlook' anahtarı açık. Program klasik Outlook profilini arka planda okumayı deneyecek.");
        }

        // ---------------- Arayüz ----------------

        void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = Padding.Empty, Margin = Padding.Empty };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            Controls.Add(root);
            root.Controls.Add(new HeaderBand(), 0, 0);

            // Gövde: iki sütun
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = Theme.Bg, Padding = new Padding(16, 14, 16, 2), Margin = Padding.Empty };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 54));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 46));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(body, 0, 1);

            var left = BuildSourcesCard();
            left.Margin = new Padding(0, 0, 7, 12);
            body.Controls.Add(left, 0, 0);

            var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = Theme.Bg, Margin = new Padding(7, 0, 0, 0), Padding = Padding.Empty };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            right.Controls.Add(BuildOutputCard(), 0, 0);
            right.Controls.Add(BuildLogCard(), 0, 1);
            body.Controls.Add(right, 1, 0);

            root.Controls.Add(BuildActionBar(), 0, 2);
            root.Controls.Add(BuildFooter(), 0, 3);

            Resize += (s, e) => FitWrapLabels();
            Shown += (s, e) => FitWrapLabels();
        }

        // Kart içeriği: tek sütunlu tablo
        static TableLayoutPanel CardContent(Card card, bool fill)
        {
            var t = new TableLayoutPanel { Dock = fill ? DockStyle.Fill : DockStyle.Top, ColumnCount = 1, AutoSize = !fill, AutoSizeMode = AutoSizeMode.GrowAndShrink, BackColor = Theme.Surface };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            card.Controls.Add(t);
            return t;
        }

        // Solda seçenek, sağda düğmeler
        static TableLayoutPanel SplitRow(Control left, params Control[] right)
        {
            var t = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, BackColor = Theme.Surface, Margin = Padding.Empty };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.Controls.Add(left, 0, 0);
            var f = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, BackColor = Theme.Surface, Margin = new Padding(6, 2, 0, 0) };
            foreach (var c in right) { c.Margin = new Padding(6, 0, 0, 0); f.Controls.Add(c); }
            t.Controls.Add(f, 1, 0);
            return t;
        }

        Label Wrap(Label l)
        {
            wrapLabels.Add(l);
            return l;
        }

        void FitWrapLabels()
        {
            foreach (var l in wrapLabels)
            {
                if (l.Parent == null) continue;
                int w = l.Parent.ClientSize.Width - l.Margin.Horizontal - 4;
                if (w > 100) l.MaximumSize = new Size(w, 0);
            }
        }

        Card BuildSourcesCard()
        {
            var card = new Card();
            var c = CardContent(card, false);
            c.Controls.Add(new StepHeader(1, "Nereden alınacak?", "Outlook hesapları, diskteki eski arşivler ya da doğrudan sunucu"));

            var optOutlook = new OptionRow("Outlook'taki tüm hesaplar", "IMAP, POP, Exchange / Microsoft 365 ve profile bağlı PST dosyaları", true);
            chkOutlook = optOutlook.Box;
            btnListStores = Theme.Secondary("Hesapları göster", (s, e) => ListStores());
            c.Controls.Add(SplitRow(optOutlook, btnListStores));
            lstStores = new CheckedListBox { CheckOnClick = true, IntegralHeight = false, HorizontalScrollbar = true, BorderStyle = BorderStyle.None, BackColor = Theme.Surface };
            storesFrame = Theme.Framed(lstStores, 74);
            storesFrame.Visible = false;
            storesFrame.Margin = new Padding(18, 3, 0, 2);
            c.Controls.Add(storesFrame);
            lblStoreHint = Wrap(Theme.Hint("Listelemeden başlatırsanız uygun hesapların hepsi alınır."));
            lblStoreHint.Margin = new Padding(18, 1, 0, 0);
            c.Controls.Add(lblStoreHint);

            c.Controls.Add(Theme.Separator());

            var optPst = new OptionRow("Diskteki diğer PST arşivleri", "Outlook'a bağlı olmayan eski .pst dosyaları", true);
            chkPst = optPst.Box;
            btnScan = Theme.Secondary("Diskleri tara", (s, e) => ScanDisks());
            btnAddPst = Theme.Secondary("PST ekle…", (s, e) => AddPst());
            c.Controls.Add(SplitRow(optPst, btnScan, btnAddPst));
            lstPst = new CheckedListBox { CheckOnClick = true, IntegralHeight = false, HorizontalScrollbar = true, BorderStyle = BorderStyle.None, BackColor = Theme.Surface };
            pstFrame = Theme.Framed(lstPst, 56);
            pstFrame.Visible = false;
            pstFrame.Margin = new Padding(18, 3, 0, 2);
            c.Controls.Add(pstFrame);
            lblPstHint = Wrap(Theme.Hint("Taramazsanız tarama aktarım sırasında yapılır."));
            lblPstHint.Margin = new Padding(18, 1, 0, 0);
            c.Controls.Add(lblPstHint);

            c.Controls.Add(Theme.Separator());

            var optServer = new OptionRow("Sunucudan doğrudan indir (IMAP / POP3)", "Outlook yoksa ya da yeni Outlook kullanılıyorsa · parola gerekir", false);
            chkServer = optServer.Box;
            lblServerInfo = optServer.Description;
            btnServer = Theme.Secondary("Sunucu bilgileri…", (s, e) => EditServer());
            chkServer.CheckedChanged += (s, e) => { if (chkServer.Checked && serverCfg == null) EditServer(); };
            c.Controls.Add(SplitRow(optServer, btnServer));
            return card;
        }

        // Yeni Outlook: yolları anlatıp sunucudan indirme penceresini açmayı önerir.
        void OfferServer(string reason)
        {
            OutlookSession.NewOutlookDetected = false;
            AppendLog(reason);
            var ans = MessageBox.Show(this, reason + "\r\n\r\n" + NewOutlookException.Ways +
                "\r\n\r\nSunucudan indirme ayarlarını şimdi açayım mı?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (ans == DialogResult.Yes) chkServer.Checked = true;
        }

        // Sunucu ayarları ayrı küçük pencerede; ana ekranda yalnız özet satırı görünür.
        bool EditServer()
        {
            using (var d = new ServerDialog(serverCfg, reporter))
            {
                if (d.ShowDialog(this) == DialogResult.OK)
                {
                    serverCfg = d.Result;
                    chkServer.Checked = true;
                }
            }
            if (serverCfg == null) chkServer.Checked = false;
            lblServerInfo.Text = serverCfg == null
                ? "Outlook yoksa ya da yeni Outlook kullanılıyorsa · parola gerekir"
                : serverCfg.Protocol + " · " + serverCfg.Host + ":" + serverCfg.Port + " · " + serverCfg.User;
            return serverCfg != null;
        }

        Card BuildOutputCard()
        {
            var card = new Card { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var c = CardContent(card, false);
            c.Controls.Add(new StepHeader(2, "Nereye, hangi biçimde?", "Seçilen yerde 'Mailler_BİLGİSAYARADI_tarih' klasörü açılır"));

            txtTarget = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 0) };
            txtTarget.TextChanged += (s, e) => UpdateCloudWarning();
            btnBrowse = Theme.Secondary("Değiştir…", (s, e) => BrowseTarget());
            var row = new TableLayoutPanel { ColumnCount = 3, AutoSize = true, Dock = DockStyle.Fill, BackColor = Theme.Surface, Margin = new Padding(0, 0, 0, 2) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.Controls.Add(Theme.Label("Kayıt yeri"), 0, 0);
            row.Controls.Add(txtTarget, 1, 0);
            btnBrowse.Margin = new Padding(6, 0, 0, 0);
            row.Controls.Add(btnBrowse, 2, 0);
            c.Controls.Add(row);
            lblCloud = Wrap(Theme.Hint(""));
            lblCloud.ForeColor = Theme.Warn;
            c.Controls.Add(lblCloud);

            c.Controls.Add(Theme.Separator());

            var eml = new OptionRow("EML dosyaları  (önerilir)", "Her mail ayrı dosya · her programda açılır, her hesaba yüklenebilir", true);
            var pst = new OptionRow("Tek PST dosyası  (önerilir)", "Outlook'ta 'İçeri Aktar' ile yeni hesaba tek hamlede alınır", true);
            var msg = new OptionRow("MSG dosyaları", "Outlook'un kendi biçimi · isteğe bağlı, ek yer kaplar", false);
            chkEml = eml.Box;
            chkPstOut = pst.Box;
            chkMsg = msg.Box;
            c.Controls.Add(eml);
            c.Controls.Add(pst);
            c.Controls.Add(msg);
            return card;
        }

        Card BuildLogCard()
        {
            var card = new Card { Margin = new Padding(0, 0, 0, 12) };
            var c = CardContent(card, true);
            c.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            c.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            c.Controls.Add(new StepHeader(3, "Günlük", null), 0, 0);
            txtLog = new TextBox
            {
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, WordWrap = true,
                Font = Theme.Mono, BorderStyle = BorderStyle.None, BackColor = Theme.LogBg, ForeColor = Color.FromArgb(55, 65, 81)
            };
            var frame = Theme.Framed(txtLog, 60);
            frame.Margin = new Padding(0, 2, 0, 0);
            c.Controls.Add(frame, 0, 1);
            return card;
        }

        // Pencerenin altında sabit duran işlem çubuğu
        Control BuildActionBar()
        {
            var bar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, BackColor = Theme.Surface, Padding = new Padding(16, 10, 16, 10), Margin = Padding.Empty };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bar.Paint += (s, e) => { using (var p = new Pen(Theme.Border)) e.Graphics.DrawLine(p, 0, 0, bar.Width, 0); };

            btnStart = Theme.Primary("Aktarmayı Başlat", (s, e) => StartExport());
            btnStop = Theme.Secondary("Durdur", (s, e) => StopWork());
            btnStop.Enabled = false;
            btnOpen = Theme.Secondary("Klasörü aç", (s, e) => OpenPackage());
            btnOpen.Enabled = false;
            var actions = new FlowLayoutPanel { AutoSize = true, BackColor = Theme.Surface, WrapContents = false, Margin = new Padding(0, 0, 0, 8) };
            btnStop.Margin = btnOpen.Margin = new Padding(8, 3, 0, 0);
            actions.Controls.Add(btnStart);
            actions.Controls.Add(btnStop);
            actions.Controls.Add(btnOpen);
            bar.Controls.Add(actions, 0, 0);

            lblStatus = new Label
            {
                AutoSize = false, AutoEllipsis = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Theme.Muted,
                Margin = new Padding(16, 0, 0, 8), Text = "Başlamaya hazır. Yarıda kalırsa aynı yere tekrar başlatın: kaldığı yerden devam eder."
            };
            bar.Controls.Add(lblStatus, 1, 0);

            progress = new ProgressBar { Dock = DockStyle.Fill, Height = 8, Maximum = 1000, Margin = Padding.Empty };
            bar.Controls.Add(progress, 0, 1);
            bar.SetColumnSpan(progress, 2);
            return bar;
        }

        Control BuildFooter()
        {
            var bar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = Color.FromArgb(232, 236, 241), Padding = new Padding(16, 0, 16, 0), Margin = Padding.Empty };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.Controls.Add(new Label { Text = "Mail Aktarıcı " + Program.Version + "  ·  MIT lisansı", AutoSize = true, ForeColor = Theme.Muted, Font = Theme.Small, Anchor = AnchorStyles.Left, Margin = new Padding(0, 9, 0, 0) }, 0, 0);
            var right = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Anchor = AnchorStyles.Right, BackColor = bar.BackColor, Margin = new Padding(0, 8, 0, 0) };
            right.Controls.Add(new Label { Text = "by " + Program.Author, AutoSize = true, Font = Theme.Bold, ForeColor = Theme.Text, Margin = new Padding(0, 0, 10, 0) });
            right.Controls.Add(Link("LinkedIn", Program.LinkedInUrl));
            right.Controls.Add(new Label { Text = "·", AutoSize = true, ForeColor = Theme.Muted, Margin = new Padding(4, 0, 4, 0) });
            right.Controls.Add(Link("GitHub", Program.GitHubUrl));
            bar.Controls.Add(right, 1, 0);
            return bar;
        }

        static LinkLabel Link(string text, string url)
        {
            var l = new LinkLabel { Text = text, AutoSize = true, LinkColor = Theme.Accent, ActiveLinkColor = Theme.AccentDark, VisitedLinkColor = Theme.Accent, LinkBehavior = LinkBehavior.HoverUnderline, Margin = Padding.Empty };
            l.LinkClicked += (s, e) => Theme.OpenUrl(url);
            return l;
        }

        // ---------------- Günlük ve ilerleme ----------------

        void AppendLog(string line)
        {
            logQueue.Enqueue(DateTime.Now.ToString("HH:mm:ss") + "  " + line);
        }

        void PumpUi()
        {
            string line;
            var sb = new System.Text.StringBuilder();
            int n = 0;
            while (n < 500 && logQueue.TryDequeue(out line))
            {
                sb.Append(line).Append("\r\n");
                n++;
            }
            if (sb.Length > 0)
            {
                if (txtLog.TextLength > 400000) txtLog.Text = txtLog.Text.Substring(txtLog.TextLength - 200000);
                txtLog.AppendText(sb.ToString());
            }
            if (worker != null)
            {
                long total = reporter.Total, done = reporter.Done;
                progress.Style = total > 0 ? ProgressBarStyle.Continuous : ProgressBarStyle.Marquee;
                if (total > 0) progress.Value = (int)Math.Min(1000, done * 1000 / total);
                string st = reporter.Status;
                lblStatus.ForeColor = Theme.Text;
                lblStatus.Text = (total > 0 ? "%" + Math.Min(100, done * 100 / total) + "   ·   " + done.ToString("N0") + " / " + total.ToString("N0") + "   ·   " : "") + st;
            }
        }

        // ---------------- Arka plan işleri ----------------

        bool StartWorker(string what, Action job, Action done)
        {
            if (worker != null) return false;
            reporter.CancelRequested = false;
            reporter.ResetProgress();
            SetBusy(true);
            reporter.Status = what;
            worker = new Thread(() =>
            {
                try { job(); }
                catch (OperationCanceledException) { reporter.Warn("Durduruldu."); }
                catch (Exception ex) { reporter.Error(ex.Message); }
                finally
                {
                    try
                    {
                        BeginInvoke(new Action(() =>
                        {
                            worker = null;
                            SetBusy(false);
                            progress.Style = ProgressBarStyle.Continuous;
                            PumpUi();
                            if (done != null) done();
                            if (closeAfterWork) Close();
                        }));
                    }
                    catch { }
                }
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.IsBackground = true;
            worker.Start();
            return true;
        }

        void SetBusy(bool busy)
        {
            foreach (Control c in new Control[] { btnStart, btnListStores, btnScan, btnAddPst, btnBrowse, btnServer, chkOutlook, chkPst, chkServer, chkEml, chkPstOut, chkMsg, lstStores, lstPst, txtTarget })
                c.Enabled = !busy;
            btnStop.Enabled = busy;
        }

        void Idle(string text)
        {
            lblStatus.ForeColor = Theme.Muted;
            lblStatus.Text = text ?? "Başlamaya hazır.";
        }

        void StopWork()
        {
            if (worker == null) return;
            reporter.CancelRequested = true;
            reporter.Status = "Durduruluyor, eldeki mail bitince duracak...";
            btnStop.Enabled = false;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (worker != null)
            {
                var ans = MessageBox.Show(this, "Aktarım sürüyor. Durdurup çıkılsın mı? (O ana kadar kopyalananlar klasörde kalır.)",
                    "Mail Aktarıcı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                e.Cancel = true;
                if (ans == DialogResult.Yes)
                {
                    closeAfterWork = true;
                    StopWork();
                }
                return;
            }
            base.OnFormClosing(e);
        }

        static void FitList(Panel frame, CheckedListBox list, int maxHeight)
        {
            int h = Math.Max(Theme.S(36), Math.Min(Theme.S(maxHeight), list.Items.Count * list.ItemHeight + 6));
            frame.Height = h;
            frame.Visible = list.Items.Count > 0;
        }

        void ListStores()
        {
            var found = new List<OutlookSource>();
            StartWorker("Outlook hesapları okunuyor...", () =>
            {
                MessageFilter.Register();
                try
                {
                    using (var s = OutlookSession.Open(reporter))
                    {
                        var ex = new OutlookExporter(s, null, reporter, null);
                        foreach (var st in OutlookExporter.ListProfileStores(s.Ns))
                        {
                            reporter.Check();
                            reporter.Status = "Sayılıyor: " + st.Name;
                            ex.Count(st);
                            Com.Release(st.Store);
                            st.Store = null;
                            found.Add(st);
                            reporter.Info("Bulundu: " + st);
                        }
                        ex.Cleanup();
                    }
                }
                finally { MessageFilter.Revoke(); }
            }, () =>
            {
                lstStores.Items.Clear();
                foreach (var st in found) lstStores.Items.Add(st, st.DefaultSelected);
                storesLoaded = found.Count > 0;
                FitList(storesFrame, lstStores, 74);
                lblStoreHint.Text = found.Count > 0
                    ? found.Count + " hesap, " + found.Sum(x => x.Items).ToString("N0") + " mail, " + Util.FormatSize(found.Sum(x => x.Bytes)) + " · istemediğinizin işaretini kaldırın"
                    : "Outlook'ta hesap bulunamadı.";
                Idle(null);
                if (OutlookSession.NewOutlookDetected) OfferServer("Bu bilgisayarda yeni Outlook kullanılıyor.");
            });
        }

        void ScanDisks()
        {
            var found = new List<string>();
            StartWorker("Disklerde PST aranıyor...", () =>
            {
                found.AddRange(PstFinder.Find(PstFinder.DefaultRoots(), reporter, null));
            }, () =>
            {
                pstScanned = true;
                foreach (string p in found)
                {
                    if (lstPst.Items.Cast<object>().Any(x => string.Equals(x.ToString(), PstLabel(p), StringComparison.OrdinalIgnoreCase))) continue;
                    lstPst.Items.Add(PstLabel(p), !PstFinder.IsOwnOutput(p));
                }
                FitList(pstFrame, lstPst, 56);
                lblPstHint.Text = found.Count + " PST bulundu · Outlook'a zaten bağlı olanlar iki kez alınmaz";
                reporter.Info("PST taraması bitti: " + found.Count + " dosya.");
                Idle(null);
            });
        }

        static string PstLabel(string path)
        {
            long size = 0;
            try { size = new FileInfo(path).Length; }
            catch { }
            return path + "   (" + Util.FormatSize(size) + ")";
        }

        static string PathFromLabel(string label)
        {
            int i = label.LastIndexOf("   (", StringComparison.Ordinal);
            return i > 0 ? label.Substring(0, i) : label;
        }

        void AddPst()
        {
            using (var d = new OpenFileDialog { Filter = "Outlook veri dosyası (*.pst)|*.pst", Multiselect = true, Title = "PST dosyası seçin" })
            {
                if (d.ShowDialog(this) != DialogResult.OK) return;
                foreach (string f in d.FileNames) lstPst.Items.Add(PstLabel(f), true);
                pstScanned = true;
                chkPst.Checked = true;
                FitList(pstFrame, lstPst, 56);
                lblPstHint.Text = "Listede işaretli PST dosyaları aktarılır.";
            }
        }

        void BrowseTarget()
        {
            using (var d = new FolderBrowserDialog { Description = "Aktarım klasörünün oluşturulacağı yeri seçin (ör. Masaüstü ya da flash bellek)", SelectedPath = txtTarget.Text })
            {
                if (d.ShowDialog(this) == DialogResult.OK) txtTarget.Text = d.SelectedPath;
            }
        }

        void UpdateCloudWarning()
        {
            string p = txtTarget.Text;
            lblCloud.Text = Util.IsCloudSynced(p)
                ? "Dikkat: bu klasör OneDrive/bulut ile eşitleniyor; mailler buluta yüklenir. Başka bir yer (ör. flash bellek) seçin."
                : "";
            lblCloud.Visible = lblCloud.Text.Length > 0;
        }

        void StartExport()
        {
            if (!chkOutlook.Checked && !chkPst.Checked && !chkServer.Checked)
            {
                MessageBox.Show(this, "En az bir kaynak seçin.", Text);
                return;
            }
            if (!chkEml.Checked && !chkPstOut.Checked && !chkMsg.Checked && !chkServer.Checked)
            {
                MessageBox.Show(this, "En az bir biçim seçin (EML, PST ya da MSG).", Text);
                return;
            }
            string target = txtTarget.Text.Trim();
            try
            {
                Directory.CreateDirectory(target);
                string probe = Path.Combine(target, ".yazma_testi_" + Guid.NewGuid().ToString("N"));
                File.WriteAllText(probe, "x");
                File.Delete(probe);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Kayıt yerine yazılamıyor: " + ex.Message, Text);
                return;
            }
            if (Util.IsCloudSynced(target) &&
                MessageBox.Show(this, "Kayıt yeri bulutla (OneDrive vb.) eşitleniyor; mailler buluta da yüklenir. Yine de devam edilsin mi?",
                    Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            var job = new ExportJob { TargetParent = target, Outlook = chkOutlook.Checked };
            var unfinished = Package.FindUnfinished(target);
            if (unfinished.Count > 0)
            {
                var ans = MessageBox.Show(this,
                    "Bu bilgisayar için yarım kalmış bir aktarım bulundu:\r\n" + unfinished[0] +
                    "\r\n\r\nKaldığı yerden devam edilsin mi?\r\n\r\nEvet: aynı klasöre devam eder, sağlam yazılmış mailleri yeniden yazmaz.\r\n" +
                    "Hayır: yeni bir klasörle baştan başlar.",
                    Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (ans == DialogResult.Cancel) return;
                if (ans == DialogResult.Yes) job.ContinueDir = unfinished[0];
            }
            job.Confirm = question =>
            {
                bool ok = false;
                Invoke(new Action(() => ok = MessageBox.Show(this, question, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes));
                return ok;
            };
            if (chkOutlook.Checked && storesLoaded)
            {
                job.StoreIds = lstStores.CheckedItems.Cast<OutlookSource>().Select(s => s.StoreId).ToList();
                if (job.StoreIds.Count == 0) job.Outlook = false;
            }
            if (chkPst.Checked)
            {
                if (pstScanned) job.PstFiles = lstPst.CheckedItems.Cast<object>().Select(x => PathFromLabel(x.ToString())).ToList();
                else job.ScanDisks = true;
            }
            if (chkServer.Checked)
            {
                if (serverCfg == null && !EditServer()) return;
                job.Server = serverCfg;
            }
            job.Opt.Eml = chkEml.Checked;
            job.Opt.SinglePst = chkPstOut.Checked;
            job.Opt.Msg = chkMsg.Checked;

            txtLog.Clear();
            string package = null;
            StartWorker("Başlıyor...", () => { package = Jobs.RunExport(job, reporter); }, () =>
            {
                if (package == null) return;
                lastPackage = package;
                btnOpen.Enabled = true;
                lblStatus.Text = reporter.Status;
                progress.Value = progress.Maximum;
                if (closeAfterWork) return;
                if (OutlookSession.NewOutlookDetected)
                {
                    OfferServer("Bu bilgisayarda yeni Outlook kullanılıyor; Outlook'taki hesaplar alınamadı.");
                    return;
                }
                string msg = "Aktarım bitti: " + reporter.Status + "\r\n\r\nKlasör:\r\n" + package +
                    "\r\n\r\nAyrıntılar klasördeki 'ozet_rapor.txt' dosyasında." +
                    (reporter.LocalLogPath != null ? "\r\nGünlüğün bu bilgisayardaki kopyası: " + reporter.LocalLogPath : "") +
                    "\r\n\r\nKlasör açılsın mı?";
                if (MessageBox.Show(this, msg, "Mail Aktarıcı", MessageBoxButtons.YesNo,
                        reporter.Errors > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information) == DialogResult.Yes)
                    OpenPackage();
            });
        }

        void OpenPackage()
        {
            if (lastPackage == null || !Directory.Exists(lastPackage)) return;
            try { Process.Start("explorer.exe", "\"" + lastPackage + "\""); }
            catch { }
        }
    }
}
