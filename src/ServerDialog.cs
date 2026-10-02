using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace MailAktarici
{
    // "Sunucudan doğrudan indir" ayarları: ana ekran sabit kalsın diye ayrı küçük pencerede.
    sealed class ServerDialog : Form
    {
        readonly Reporter reporter;
        ComboBox cmbPreset, cmbProto, cmbSec;
        TextBox txtHost, txtUser, txtPass;
        NumericUpDown numPort;
        CheckBox chkIgnoreCert;
        Button btnTest, btnOk, btnCancel;
        Label lblNote;

        public ServerConfig Result { get; private set; }

        public ServerDialog(ServerConfig current, Reporter reporter)
        {
            this.reporter = reporter;
            Text = "Sunucu bilgileri";
            Font = Theme.Base;
            BackColor = Theme.Surface;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(18, 14, 18, 14);
            Icon = Theme.AppIcon(32);
            Build();
            LoadPresets();
            if (current != null)
            {
                cmbProto.SelectedItem = current.IsPop ? "POP3" : "IMAP";
                cmbSec.SelectedIndex = current.Sec == Security.SslTls ? 0 : current.Sec == Security.StartTls ? 1 : 2;
                txtHost.Text = current.Host;
                numPort.Value = current.Port;
                txtUser.Text = current.User;
                txtPass.Text = current.Pass;
                chkIgnoreCert.Checked = current.IgnoreCert;
            }
        }

        void Build()
        {
            var t = new TableLayoutPanel { ColumnCount = 4, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, BackColor = Theme.Surface };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));

            var head = new StepHeader(0, "Sunucudan doğrudan indir", "IMAP ya da POP3 · Outlook yoksa ya da yeni Outlook kullanılıyorsa");
            head.Width = 520;
            head.Margin = new Padding(0, 0, 0, 10);
            t.Controls.Add(head, 0, 0);
            t.SetColumnSpan(head, 4);

            cmbPreset = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
            cmbPreset.SelectedIndexChanged += (s, e) => ApplyPreset();
            cmbProto = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
            cmbProto.Items.AddRange(new object[] { "IMAP", "POP3" });
            cmbProto.SelectedIndex = 0;
            cmbSec = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
            cmbSec.Items.AddRange(new object[] { "SSL/TLS", "STARTTLS", "Şifresiz" });
            cmbSec.SelectedIndex = 0;
            cmbProto.SelectedIndexChanged += (s, e) => SuggestPort();
            cmbSec.SelectedIndexChanged += (s, e) => SuggestPort();
            txtHost = new TextBox { Dock = DockStyle.Fill };
            numPort = new NumericUpDown { Minimum = 1, Maximum = 65535, Value = 993, Dock = DockStyle.Fill };
            txtUser = new TextBox { Dock = DockStyle.Fill };
            txtPass = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            chkIgnoreCert = new CheckBox { Text = "Sertifika hatalarını yoksay", AutoSize = true };

            int r = 1;
            Add(t, Theme.Label("Hazır ayar"), 0, r); Add(t, cmbPreset, 1, r, 3); r++;
            Add(t, Theme.Label("Protokol"), 0, r); Add(t, cmbProto, 1, r); Add(t, Theme.Label("Güvenlik"), 2, r); Add(t, cmbSec, 3, r); r++;
            Add(t, Theme.Label("Sunucu"), 0, r); Add(t, txtHost, 1, r); Add(t, Theme.Label("Port"), 2, r); Add(t, numPort, 3, r); r++;
            Add(t, Theme.Label("Kullanıcı"), 0, r); Add(t, txtUser, 1, r, 3); r++;
            Add(t, Theme.Label("Parola"), 0, r); Add(t, txtPass, 1, r, 3); r++;
            Add(t, chkIgnoreCert, 1, r, 3); r++;
            lblNote = Theme.Hint("Parola hiçbir yere kaydedilmez. Gmail, Yahoo ve iCloud normal parola yerine 'uygulama şifresi' ister.");
            lblNote.MaximumSize = new Size(500, 0);
            Add(t, lblNote, 0, r, 4); r++;
            foreach (Control c in t.Controls)
            {
                if (c is StepHeader) continue;
                c.Margin = new Padding(0, 3, 8, 3);
                if (c is Label && c != lblNote) c.Anchor = AnchorStyles.Left;
            }

            btnTest = Theme.Secondary("Bağlantıyı test et", (s, e) => Test());
            btnOk = Theme.Primary("Tamam", (s, e) => Accept());
            btnCancel = Theme.Secondary("Vazgeç", (s, e) => { DialogResult = DialogResult.Cancel; Close(); });
            var bar = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill, BackColor = Theme.Surface, Margin = new Padding(0, 12, 0, 0) };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.Controls.Add(btnTest, 0, 0);
            var right = new FlowLayoutPanel { AutoSize = true, WrapContents = false, BackColor = Theme.Surface, Margin = Padding.Empty };
            btnCancel.Margin = new Padding(0, 3, 8, 0);
            right.Controls.Add(btnCancel);
            right.Controls.Add(btnOk);
            bar.Controls.Add(right, 1, 0);
            Add(t, bar, 0, r, 4);

            Controls.Add(t);
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        static void Add(TableLayoutPanel t, Control c, int col, int row, int span = 1)
        {
            t.Controls.Add(c, col, row);
            if (span > 1) t.SetColumnSpan(c, span);
        }

        void LoadPresets()
        {
            cmbPreset.Items.Add("(elle doldur)");
            try { foreach (var a in AccountSources.FromOutlookRegistry()) cmbPreset.Items.Add(a); }
            catch { }
            foreach (var a in AccountSources.Presets()) cmbPreset.Items.Add(a);
            cmbPreset.SelectedIndex = 0;
        }

        void ApplyPreset()
        {
            var a = cmbPreset.SelectedItem as AccountPreset;
            if (a == null) return;
            cmbProto.SelectedItem = a.Protocol;
            cmbSec.SelectedIndex = a.Sec == Security.SslTls ? 0 : a.Sec == Security.StartTls ? 1 : 2;
            txtHost.Text = a.Host;
            numPort.Value = a.Port;
            if (!string.IsNullOrEmpty(a.User)) txtUser.Text = a.User;
            lblNote.Text = a.Note ?? "Parola hiçbir yere kaydedilmez.";
        }

        void SuggestPort()
        {
            int[] known = { 993, 143, 995, 110 };
            if (!known.Contains((int)numPort.Value)) return;
            Security sec = cmbSec.SelectedIndex == 0 ? Security.SslTls : cmbSec.SelectedIndex == 1 ? Security.StartTls : Security.None;
            numPort.Value = ServerConfig.DefaultPort((string)cmbProto.SelectedItem == "POP3", sec);
        }

        ServerConfig Read(bool requireUser)
        {
            if (txtHost.Text.Trim().Length == 0) { MessageBox.Show(this, "Sunucu adresini yazın.", Text); return null; }
            if (requireUser && (txtUser.Text.Trim().Length == 0 || txtPass.Text.Length == 0))
            {
                MessageBox.Show(this, "Kullanıcı adı ve parolayı yazın.", Text);
                return null;
            }
            return new ServerConfig
            {
                Protocol = (string)cmbProto.SelectedItem,
                Host = txtHost.Text.Trim(),
                Port = (int)numPort.Value,
                Sec = cmbSec.SelectedIndex == 0 ? Security.SslTls : cmbSec.SelectedIndex == 1 ? Security.StartTls : Security.None,
                User = txtUser.Text.Trim().Length > 0 ? txtUser.Text.Trim() : null,
                Pass = txtPass.Text,
                IgnoreCert = chkIgnoreCert.Checked
            };
        }

        void Accept()
        {
            var c = Read(true);
            if (c == null) return;
            Result = c;
            DialogResult = DialogResult.OK;
            Close();
        }

        void Test()
        {
            var c = Read(false);
            if (c == null) return;
            btnTest.Enabled = btnOk.Enabled = false;
            UseWaitCursor = true;
            var th = new Thread(() =>
            {
                string result;
                try { result = ServerExporter.TestConnection(c, reporter); }
                catch (Exception ex) { result = "Bağlantı başarısız: " + ex.Message; }
                if (reporter != null) reporter.Info(result);
                try
                {
                    BeginInvoke(new Action(() =>
                    {
                        UseWaitCursor = false;
                        btnTest.Enabled = btnOk.Enabled = true;
                        MessageBox.Show(this, result, "Bağlantı testi", MessageBoxButtons.OK,
                            result.StartsWith("Bağlantı başarısız") ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                    }));
                }
                catch { }
            });
            th.IsBackground = true;
            th.Start();
        }
    }
}
