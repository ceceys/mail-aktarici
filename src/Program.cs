using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MailAktarici
{
    static class Program
    {
        public const string Version = "1.1.1";
        public const string Author = "cecey";
        public const string LinkedInUrl = "https://www.linkedin.com/in/cuma-ali-dirik/";
        public const string GitHubUrl = "https://github.com/ceceys";

        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int pid);

        [STAThread]
        static int Main(string[] args)
        {
            // Hiçbir hata programı sessizce kapatmasın: sebep Masaüstüne yazılır.
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                string p = CrashReport.Write(e.ExceptionObject as Exception, "arka plan");
                if (Environment.UserInteractive && args.Length == 0)
                {
                    try
                    {
                        MessageBox.Show("Mail Aktarıcı beklenmeyen bir hatayla kapanmak zorunda.\r\n\r\nSebebi şu dosyaya yazıldı:\r\n" + (p ?? "(yazılamadı)") +
                            "\r\n\r\n" + (e.ExceptionObject is Exception ? ((Exception)e.ExceptionObject).Message : ""), "Mail Aktarıcı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch { }
                }
            };
            if (args.Length == 0)
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) =>
                {
                    string p = CrashReport.Write(e.Exception, "arayüz");
                    MessageBox.Show("Beklenmeyen bir hata oldu, program çalışmaya devam ediyor.\r\n\r\n" + e.Exception.Message +
                        "\r\n\r\nAyrıntı: " + (p ?? "(yazılamadı)"), "Mail Aktarıcı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                };
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
                return 0;
            }
            AttachConsole(-1);
            try { Console.OutputEncoding = new UTF8Encoding(false); }
            catch { }
            Console.WriteLine();
            try { return Cli.Run(args); }
            catch (Exception ex)
            {
                Console.WriteLine("HATA: " + ex.Message);
                return 2;
            }
        }
    }

    // Komut satırı: toplu kullanım ve otomatik testler için.
    static class Cli
    {
        static readonly HashSet<string> Flags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "--outlook-yok", "--pst-tara", "--msg", "--eml-yok", "--pst-yok", "--sertifika-yoksay"
        };

        const string Help =
@"Mail Aktarıcı " + Program.Version + @" · by " + Program.Author + @"
" + Program.LinkedInUrl + @"  ·  " + Program.GitHubUrl + @"

Parametresiz çalıştırılırsa pencere açılır. Komut satırı:

  MailAktarici.exe listele
      Outlook'taki hesapları ve veri dosyalarını listeler.

  MailAktarici.exe disa-aktar [seçenekler]
      --hedef KLASÖR          Aktarım klasörünün oluşacağı yer (varsayılan: Masaüstü)
      --devam KLASÖR          Yarım kalmış bu Mailler_... klasörüne kaldığı yerden devam et
      --hesap AD              Yalnız adında AD geçen Outlook hesapları (tekrarlanabilir)
      --outlook-yok           Outlook profilindeki hesapları alma
      --pst DOSYA             Bu PST dosyasını da aktar (tekrarlanabilir)
      --pst-tara              Disklerde PST ara ve aktar
      --tarama-dizini KLASÖR  PST aramasını bu klasörle sınırla (tekrarlanabilir)
      --msg                   .msg kopyası da üret
      --eml-yok               .eml üretme
      --pst-yok               Tek PST dosyası üretme
      --sinir N               Her klasörden en fazla N mail (deneme için)
      Sunucudan indirme:
      --sunucu HOST --protokol imap|pop3 --port N --guvenlik ssl|starttls|yok
      --kullanici AD --parola P (ya da MAILAKTARICI_PAROLA ortam değişkeni) [--sertifika-yoksay]

  MailAktarici.exe test-baglanti --sunucu HOST [--protokol ..] [--port ..] [--guvenlik ..] [--kullanici .. --parola ..]

Çıkış kodu: 0 sorunsuz, 1 hatalarla bitti, 2 çalışamadı.";

        public static int Run(string[] args)
        {
            string cmd = args[0].ToLowerInvariant();
            var o = Parse(args.Skip(1).ToArray());
            var r = new Reporter();
            r.Sink = line => Console.WriteLine(line);
            CrashReport.Current = r;
            switch (cmd)
            {
                case "listele": return List(r);
                case "disa-aktar": return Export(o, r);
                case "test-baglanti": return Test(o, r);
                default:
                    Console.WriteLine(Help);
                    return cmd == "yardim" || cmd == "/?" || cmd == "-h" || cmd == "--help" ? 0 : 2;
            }
        }

        static Dictionary<string, List<string>> Parse(string[] a)
        {
            var d = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < a.Length; i++)
            {
                string k = a[i];
                if (!k.StartsWith("--")) throw new ArgumentException("Beklenmeyen parametre: " + k);
                string v = "";
                if (!Flags.Contains(k))
                {
                    if (i + 1 >= a.Length) throw new ArgumentException(k + " için değer eksik");
                    v = a[++i];
                }
                List<string> list;
                if (!d.TryGetValue(k, out list)) { list = new List<string>(); d[k] = list; }
                list.Add(v);
            }
            return d;
        }

        static string One(Dictionary<string, List<string>> o, string k, string def)
        {
            List<string> l;
            return o.TryGetValue(k, out l) && l.Count > 0 ? l[l.Count - 1] : def;
        }

        static bool Has(Dictionary<string, List<string>> o, string k) { return o.ContainsKey(k); }

        static List<string> Many(Dictionary<string, List<string>> o, string k)
        {
            List<string> l;
            return o.TryGetValue(k, out l) ? l : new List<string>();
        }

        static ServerConfig ServerFrom(Dictionary<string, List<string>> o)
        {
            string host = One(o, "--sunucu", null);
            if (host == null) return null;
            var c = new ServerConfig { Host = host };
            c.Protocol = One(o, "--protokol", "imap").ToUpperInvariant() == "POP3" ? "POP3" : "IMAP";
            switch (One(o, "--guvenlik", "ssl").ToLowerInvariant())
            {
                case "starttls": c.Sec = Security.StartTls; break;
                case "yok": c.Sec = Security.None; break;
                default: c.Sec = Security.SslTls; break;
            }
            c.Port = int.Parse(One(o, "--port", ServerConfig.DefaultPort(c.IsPop, c.Sec).ToString()));
            c.User = One(o, "--kullanici", null);
            c.Pass = One(o, "--parola", Environment.GetEnvironmentVariable("MAILAKTARICI_PAROLA"));
            c.IgnoreCert = Has(o, "--sertifika-yoksay");
            return c;
        }

        static int List(Reporter r)
        {
            MessageFilter.Register();
            try
            {
                using (var s = OutlookSession.Open(r))
                {
                    var ex = new OutlookExporter(s, null, r, null);
                    foreach (var st in OutlookExporter.ListProfileStores(s.Ns))
                    {
                        ex.Count(st);
                        Console.WriteLine((st.DefaultSelected ? "[x] " : "[ ] ") + st);
                        if (!string.IsNullOrEmpty(st.FilePath)) Console.WriteLine("      " + st.FilePath);
                    }
                    ex.Cleanup();
                }
            }
            catch (Exception e) { Console.WriteLine("Outlook: " + e.Message); }
            finally
            {
                MessageFilter.Revoke();
            }
            Console.WriteLine();
            Console.WriteLine("Outlook profilindeki IMAP/POP sunucu bilgileri:");
            foreach (var a in AccountSources.FromOutlookRegistry())
                Console.WriteLine("  " + a.Label + "  ->  " + a.Host + ":" + a.Port + " " + a.Sec + " kullanıcı=" + a.User);
            return 0;
        }

        static int Export(Dictionary<string, List<string>> o, Reporter r)
        {
            var job = new ExportJob
            {
                TargetParent = One(o, "--hedef", Util.DesktopPath()),
                ContinueDir = Has(o, "--devam") ? Path.GetFullPath(One(o, "--devam", "")) : null,
                Outlook = !Has(o, "--outlook-yok"),
                StoreNameFilter = Many(o, "--hesap"),
                ScanDisks = Has(o, "--pst-tara"),
                PstFiles = Many(o, "--pst").Select(Path.GetFullPath).ToList(),
                Server = ServerFrom(o)
            };
            var roots = Many(o, "--tarama-dizini");
            if (roots.Count > 0) job.ScanRoots = roots.Select(Path.GetFullPath).ToList();
            job.Opt.Msg = Has(o, "--msg");
            job.Opt.Eml = !Has(o, "--eml-yok");
            job.Opt.SinglePst = !Has(o, "--pst-yok");
            job.Opt.LimitPerFolder = int.Parse(One(o, "--sinir", "0"));
            Directory.CreateDirectory(job.TargetParent);
            string root = Jobs.RunExport(job, r);
            Console.WriteLine();
            Console.WriteLine("KLASOR=" + root);
            try { Console.WriteLine(File.ReadAllText(Path.Combine(root, "ozet_rapor.txt"))); }
            catch { }
            return r.Errors > 0 ? 1 : 0;
        }

        static int Test(Dictionary<string, List<string>> o, Reporter r)
        {
            var c = ServerFrom(o);
            if (c == null) { Console.WriteLine("--sunucu gerekli"); return 2; }
            Console.WriteLine(ServerExporter.TestConnection(c, r));
            return 0;
        }
    }
}
