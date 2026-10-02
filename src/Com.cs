using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace MailAktarici
{
    // Outlook nesne modeline geç bağlama (late binding) yardımcıları.
    // PIA/interop DLL gerekmez; hangi Outlook sürümü kuruluysa onunla konuşur.
    static class Com
    {
        public static object Get(object o, string name)
        {
            if (o == null) return null;
            try { return o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, null); }
            catch { return null; }
        }

        public static string Str(object o, string name)
        {
            object v = Get(o, name);
            return v == null ? null : v.ToString();
        }

        public static int Int(object o, string name, int def)
        {
            object v = Get(o, name);
            if (v == null) return def;
            try { return Convert.ToInt32(v); }
            catch { return def; }
        }

        public static bool Bool(object o, string name)
        {
            object v = Get(o, name);
            return v is bool && (bool)v;
        }

        public static void Set(object o, string name, object value)
        {
            o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, new[] { value });
        }

        public static object Call(object o, string name, params object[] args)
        {
            try { return o.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, o, args); }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException != null) throw ex.InnerException;
                throw;
            }
        }

        public static object Item(object coll, int index)
        {
            return Call(coll, "Item", index);
        }

        public static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o))
            {
                try { Marshal.ReleaseComObject(o); }
                catch { }
            }
        }

        public static string Tag(string hex)
        {
            return "http://schemas.microsoft.com/mapi/proptag/0x" + hex;
        }

        // PropertyAccessor.GetProperties: tek COM çağrısıyla birden çok MAPI özelliği.
        // Olmayan özellik için dizide hata kodu (negatif int) döner; bunları null yaparız.
        public static object[] Props(object obj, object[] tags)
        {
            object pa = Get(obj, "PropertyAccessor");
            if (pa == null) return new object[tags.Length];
            try
            {
                object[] v = Call(pa, "GetProperties", new object[] { tags }) as object[];
                if (v == null) return new object[tags.Length];
                for (int i = 0; i < v.Length; i++)
                    if (v[i] is int && (int)v[i] < 0) v[i] = null;
                return v;
            }
            catch
            {
                var v = new object[tags.Length];
                for (int i = 0; i < tags.Length; i++)
                {
                    try
                    {
                        v[i] = Call(pa, "GetProperty", tags[i]);
                        if (v[i] is int && (int)v[i] < 0) v[i] = null;
                    }
                    catch { v[i] = null; }
                }
                return v;
            }
            finally { Release(pa); }
        }

        public static object Prop(object obj, string tag)
        {
            return Props(obj, new object[] { tag })[0];
        }

        public static string S(object[] v, int i)
        {
            object o = v != null && i < v.Length ? v[i] : null;
            string s = o as string;
            return string.IsNullOrEmpty(s) ? null : s;
        }

        public static int I(object[] v, int i)
        {
            object o = v != null && i < v.Length ? v[i] : null;
            if (o == null) return 0;
            try { return Convert.ToInt32(o); }
            catch { return 0; }
        }

        public static long L(object[] v, int i)
        {
            object o = v != null && i < v.Length ? v[i] : null;
            if (o == null) return 0;
            try { return Convert.ToInt64(o); }
            catch { return 0; }
        }

        // PropertyAccessor tarihleri UTC döner.
        public static DateTime? D(object[] v, int i)
        {
            object o = v != null && i < v.Length ? v[i] : null;
            if (!(o is DateTime)) return null;
            DateTime d = (DateTime)o;
            if (d.Year < 1971 || d.Year > 4000) return null;
            return DateTime.SpecifyKind(d, DateTimeKind.Utc);
        }

        // Nesne modelinin (ReceivedTime gibi) yerel saatli tarihleri.
        public static DateTime? LocalDate(object o, string name)
        {
            object v = Get(o, name);
            if (!(v is DateTime)) return null;
            DateTime d = (DateTime)v;
            if (d.Year < 1971 || d.Year > 4000) return null;
            return DateTime.SpecifyKind(d, DateTimeKind.Local).ToUniversalTime();
        }

        public static string Message(Exception ex)
        {
            var com = ex as COMException;
            string m = ex.Message.Replace("\r", " ").Replace("\n", " ").Trim();
            return com != null ? m + " (0x" + com.ErrorCode.ToString("X8") + ")" : m;
        }
    }

    // Outlook meşgulken ("Çağrı aranan tarafından reddedildi") COM çağrılarını bekleyip yeniden dener.
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    sealed class MessageFilter : IOleMessageFilter
    {
        [DllImport("ole32.dll")]
        static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        public static void Register()
        {
            IOleMessageFilter old;
            CoRegisterMessageFilter(new MessageFilter(), out old);
        }

        public static void Revoke()
        {
            IOleMessageFilter old;
            CoRegisterMessageFilter(null, out old);
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0;
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // SERVERCALL_RETRYLATER (2) veya REJECTED (1): 3 dakikaya kadar 250 ms arayla tekrar dene
            if ((dwRejectType == 2 || dwRejectType == 1) && dwTickCount < 180000) return 250;
            return -1;
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return 2;
        }
    }

    sealed class OutlookSession : IDisposable
    {
        public object App, Ns;

        public static bool NewOutlookDetected;   // arayüz, iş bitince kullanıcıya yolları gösterir

        // Klasik Outlook profilleri (Kaspersky'nin geçici AvpPstTmp profilleri sayılmaz)
        public static int ClassicProfileCount()
        {
            int n = 0;
            foreach (string root in new[] { @"Software\Microsoft\Office\16.0\Outlook\Profiles", @"Software\Microsoft\Office\15.0\Outlook\Profiles" })
            {
                try
                {
                    using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(root))
                        if (k != null) n += k.GetSubKeyNames().Count(x => !x.StartsWith("AvpPstTmp", StringComparison.OrdinalIgnoreCase));
                }
                catch { }
            }
            return n;
        }

        // Klasik Outlook'taki "Yeni Outlook" anahtarı açık mı
        public static bool NewOutlookPreferred()
        {
            foreach (string key in new[] { @"Software\Microsoft\Office\16.0\Outlook\Preferences", @"Software\Policies\Microsoft\Office\16.0\Outlook\Preferences" })
            {
                try
                {
                    using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key))
                    {
                        object v = k == null ? null : k.GetValue("UseNewOutlook");
                        if (v is int && (int)v == 1) return true;
                    }
                }
                catch { }
            }
            return false;
        }

        public static bool IsNewOutlookInstalled()
        {
            try
            {
                string apps = Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Microsoft\WindowsApps\olk.exe");
                return System.IO.File.Exists(apps) || System.IO.Directory.Exists(Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Microsoft\Olk"));
            }
            catch { return false; }
        }

        public static bool IsClassicOutlookInstalled()
        {
            return Type.GetTypeFromProgID("Outlook.Application") != null;
        }

        // Outlook'a bağlanır. Kapalıysa NORMAL şekilde (görünür) açar: arka planda açılırken
        // sorduğu sorular (güvenli mod, profil seçimi, dosya onarımı) görünmez ve COM 2 dakika
        // sonra 0x80080005 ile düşer. Görünür açılışta soru ekrana gelir, program cevabı bekler.
        public static OutlookSession Open(Reporter r)
        {
            Type t = Type.GetTypeFromProgID("Outlook.Application");
            if (t == null)
                throw new ApplicationException("Bu bilgisayarda klasik Outlook (masaüstü Outlook) kurulu değil. " +
                    "Yeni Outlook uygulaması dışarıdan okunamaz; o durumda 'Sunucudan doğrudan indir' seçeneğini kullanın.");
            object app = TryActive();
            if (app == null && Process.GetProcessesByName("OUTLOOK").Length == 0)
            {
                if (ClassicProfileCount() == 0)
                    throw new NewOutlookException("Klasik Outlook'ta hiç hesap tanımlı değil" + (IsNewOutlookInstalled() ? " (hesaplar yeni Outlook'ta)." : "."));
                if (NewOutlookPreferred())
                {
                    // Yeni Outlook anahtarı açık: görünür açılış yeni Outlook'a yönlenir. Klasik profil
                    // duruyorsa arka planda (otomasyon) açılış yine çalışabilir; bir kez denenir.
                    r.Info("Bu bilgisayarda yeni Outlook açık; klasik Outlook arka planda deneniyor...");
                    try { app = Activator.CreateInstance(t); }
                    catch (COMException) { throw new NewOutlookException("Bu bilgisayarda yeni Outlook kullanılıyor."); }
                }
                else StartVisible(r);
            }
            if (app == null) app = WaitReady(t, r);
            object ns = Com.Call(app, "GetNamespace", "MAPI");
            try { Com.Call(ns, "Logon", Type.Missing, Type.Missing, false, false); }
            catch (Exception ex) { r.Warn("Outlook oturum açma uyarısı: " + Com.Message(ex)); }
            string profile = Com.Str(ns, "CurrentProfileName");
            string ver = Com.Str(app, "Version");
            r.Info("Outlook " + ver + " bağlandı, profil: " + profile);
            return new OutlookSession { App = app, Ns = ns };
        }

        static object TryActive()
        {
            try { return Marshal.GetActiveObject("Outlook.Application"); }
            catch { return null; }
        }

        static void StartVisible(Reporter r)
        {
            string exe = null;
            try
            {
                using (var k = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"))
                    if (k != null) exe = k.GetValue(null) as string;
            }
            catch { }
            if (string.IsNullOrEmpty(exe)) exe = "outlook.exe";
            r.Info("Outlook açılıyor (büyük arşiv dosyaları varsa birkaç dakika sürebilir)...");
            Process.Start(new ProcessStartInfo(exe) { UseShellExecute = true, WindowStyle = ProcessWindowStyle.Minimized });
        }

        static object WaitReady(Type t, Reporter r)
        {
            DateTime start = DateTime.UtcNow, until = start.AddMinutes(10);
            string lastPrompt = null;
            int coTries = 0;
            Exception last = null;
            DateTime goneSince = DateTime.MinValue;
            while (DateTime.UtcNow < until)
            {
                r.Check();
                object app = TryActive();
                if (app != null) return app;
                string prompt = OutlookPrompt.Find();
                if (prompt != null)
                {
                    if (prompt != lastPrompt)
                    {
                        r.Warn("Outlook ekranda bir soru soruyor, lütfen cevaplayın: \"" + prompt + "\"");
                        lastPrompt = prompt;
                    }
                    r.Status = "Outlook ekranda soru soruyor, cevap bekleniyor: " + prompt;
                    System.Threading.Thread.Sleep(1500);
                    continue;
                }
                // Çalışan Outlook'ta tabloya kayıt bazen gecikir: ara ara doğrudan bağlanmayı dene.
                if (Process.GetProcessesByName("OUTLOOK").Length > 0 && (DateTime.UtcNow - start).TotalSeconds > 15 && coTries++ % 8 == 0)
                {
                    try { return Activator.CreateInstance(t); }
                    catch (COMException ex) { last = ex; }
                }
                if (Process.GetProcessesByName("OUTLOOK").Length == 0)
                {
                    if (goneSince == DateTime.MinValue) goneSince = DateTime.UtcNow;
                    if ((DateTime.UtcNow - start).TotalSeconds > 20 && (DateTime.UtcNow - goneSince).TotalSeconds > 15)
                    {
                        if (Process.GetProcessesByName("olk").Length > 0 || IsNewOutlookInstalled())
                            throw new NewOutlookException("Klasik Outlook açılıp kapandı, yeni Outlook'a yönleniyor.");
                        throw new ApplicationException("Outlook açılıp hemen kapandı. Outlook'u elle açıp hesabın göründüğünü kontrol edin, sonra tekrar deneyin.");
                    }
                }
                else goneSince = DateTime.MinValue;
                System.Threading.Thread.Sleep(1500);
            }
            if (Process.GetProcessesByName("OUTLOOK").Length > 0)
                throw new ApplicationException("Outlook açık ama programa bağlanmıyor. Outlook 'yönetici olarak' açılmış olabilir: Outlook'u kapatıp " +
                    "bu programı tekrar başlatın (ya da ikisini de aynı şekilde açın). " + (last != null ? Com.Message(last) : ""));
            throw new ApplicationException("Outlook 10 dakika içinde açılmadı. Outlook'u elle açın, tamamen açılmasını bekleyin, sonra tekrar deneyin.");
        }
        public void Dispose()
        {
            Com.Release(Ns);
            Com.Release(App);
            Ns = App = null;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern int GetLongPathName(string shortPath, StringBuilder longPath, int bufferSize);

        public static string NormalizePath(string p)
        {
            if (string.IsNullOrEmpty(p)) return "";
            try
            {
                p = System.IO.Path.GetFullPath(p);
                var sb = new StringBuilder(1024);
                int n = GetLongPathName(p, sb, sb.Capacity);
                if (n > 0 && n < sb.Capacity) p = sb.ToString();
            }
            catch { }
            return p.TrimEnd('\\').ToLowerInvariant();
        }
    }

    // Outlook'un açılışta sorduğu soru pencereleri (güvenli mod, profil seçimi, dosya onarımı...):
    // metni okunur, pencere öne getirilir ki kullanıcı görüp cevaplasın.
    static class OutlookPrompt
    {
        delegate bool EnumProc(IntPtr h, IntPtr l);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc p, IntPtr l);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);

        static string ClassOf(IntPtr h) { var sb = new StringBuilder(64); GetClassName(h, sb, 64); return sb.ToString(); }
        static string TextOf(IntPtr h) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString(); }

        public static string Find()
        {
            var pids = new System.Collections.Generic.HashSet<uint>();
            foreach (var pr in Process.GetProcessesByName("OUTLOOK")) pids.Add((uint)pr.Id);
            if (pids.Count == 0) return null;
            string found = null;
            EnumWindows((h, l) =>
            {
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                if (!pids.Contains(pid) || !IsWindowVisible(h) || ClassOf(h) != "#32770") return true;
                var text = new StringBuilder();
                EnumChildWindows(h, (c, l2) =>
                {
                    if (ClassOf(c) == "Static")
                    {
                        string t = TextOf(c).Replace("&&", "\u0001").Replace("&", "").Replace("\u0001", "&").Trim();
                        // "DAL=on" gibi Windows'un iç işaretleri metin değil
                        if (t.Length > 0 && !System.Text.RegularExpressions.Regex.IsMatch(t, @"^[A-Z]+=\w+$")) text.Append(t).Append(' ');
                    }
                    return true;
                }, IntPtr.Zero);
                string body = text.ToString().Trim();
                found = body.Length > 0 ? body : TextOf(h);
                ShowWindow(h, 9);   // SW_RESTORE
                SetForegroundWindow(h);
                return false;
            }, IntPtr.Zero);
            return found;
        }
    }

    // Yeni Outlook kullanılan bilgisayar: mailler klasik Outlook üzerinden okunamaz.
    sealed class NewOutlookException : ApplicationException
    {
        public const string Ways =
            "Yeni Outlook'taki maillere programın doğrudan ulaşma yolu yok. Üç yol var:\r\n\r\n" +
            "1. Hesap IMAP ya da POP ise: 'Sunucudan doğrudan indir' ile sunucu adı ve parolayı girin.\r\n" +
            "2. Yeni Outlook'un sağ üstündeki 'Yeni Outlook' anahtarını kapatın; açılan klasik Outlook'a hesabı ekleyin, " +
            "eşitlenince bu programı tekrar çalıştırın.\r\n" +
            "3. Hesap Microsoft 365 / Outlook.com ise mailler zaten Microsoft bulutunda; yeni bilgisayarda o hesapla oturum açmak yeterli.";

        public NewOutlookException(string reason) : base(reason + " " + Ways.Replace("\r\n\r\n", " ").Replace("\r\n", " "))
        {
            OutlookSession.NewOutlookDetected = true;
        }
    }
}