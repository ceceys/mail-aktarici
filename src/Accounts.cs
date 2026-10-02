using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace MailAktarici
{
    // "Sunucudan indir" formunu doldurmak için hazır ayarlar.
    // Outlook profilinden yalnız sunucu, port ve kullanıcı adı okunur; parolaya dokunulmaz.
    sealed class AccountPreset
    {
        public string Label, Protocol, Host, User;
        public int Port;
        public Security Sec;
        public string Note;

        public override string ToString() { return Label; }
    }

    static class AccountSources
    {
        const string AccountsKey = "9375CFF0413111d3B88A00104B2A6676";

        public static List<AccountPreset> FromOutlookRegistry()
        {
            var list = new List<AccountPreset>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var profileRoots = new List<string>();
            foreach (string v in new[] { "16.0", "15.0" })
                profileRoots.Add(@"Software\Microsoft\Office\" + v + @"\Outlook\Profiles");
            profileRoots.Add(@"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles");

            foreach (string rootPath in profileRoots)
            {
                using (RegistryKey root = Registry.CurrentUser.OpenSubKey(rootPath))
                {
                    if (root == null) continue;
                    foreach (string profile in root.GetSubKeyNames())
                    {
                        using (RegistryKey acc = root.OpenSubKey(profile + "\\" + AccountsKey))
                        {
                            if (acc == null) continue;
                            foreach (string sub in acc.GetSubKeyNames())
                            {
                                using (RegistryKey k = acc.OpenSubKey(sub))
                                {
                                    if (k == null) continue;
                                    foreach (string proto in new[] { "IMAP", "POP3" })
                                    {
                                        string host = ReadString(k, proto + " Server");
                                        if (string.IsNullOrEmpty(host)) continue;
                                        string user = ReadString(k, proto + " User") ?? ReadString(k, "Email");
                                        int ssl = ReadInt(k, proto + " Use SSL", 0);
                                        int port = ReadInt(k, proto + " Port", 0);
                                        Security sec = ssl == 1 ? Security.SslTls : ssl == 2 ? Security.StartTls : Security.None;
                                        if (ssl == 3) sec = port == 993 || port == 995 ? Security.SslTls : Security.StartTls;
                                        bool pop = proto == "POP3";
                                        if (port == 0) port = ServerConfig.DefaultPort(pop, sec);
                                        string label = "Outlook: " + (ReadString(k, "Account Name") ?? user) + " (" + proto + ")";
                                        if (!seen.Add(label)) continue;
                                        list.Add(new AccountPreset
                                        {
                                            Label = label, Protocol = pop ? "POP3" : "IMAP", Host = host, User = user, Port = port, Sec = sec
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return list;
        }

        static string ReadString(RegistryKey k, string name)
        {
            object v = k.GetValue(name);
            if (v == null) return null;
            var b = v as byte[];
            string s = b != null ? Encoding.Unicode.GetString(b) : v.ToString();
            s = s.TrimEnd('\0').Trim();
            return s.Length == 0 ? null : s;
        }

        static int ReadInt(RegistryKey k, string name, int def)
        {
            object v = k.GetValue(name);
            if (v is int) return (int)v;
            var b = v as byte[];
            if (b != null && b.Length >= 4) return BitConverter.ToInt32(b, 0);
            return def;
        }

        public static List<AccountPreset> Presets()
        {
            return new List<AccountPreset>
            {
                new AccountPreset { Label = "Gmail (IMAP)", Protocol = "IMAP", Host = "imap.gmail.com", Port = 993, Sec = Security.SslTls,
                    Note = "Gmail normal parolayı kabul etmez: Google hesabında 'Uygulama şifresi' oluşturup onu yazın." },
                new AccountPreset { Label = "Yandex (IMAP)", Protocol = "IMAP", Host = "imap.yandex.com", Port = 993, Sec = Security.SslTls,
                    Note = "Yandex'te IMAP erişimi ayarlardan açık olmalı; uygulama şifresi gerekebilir." },
                new AccountPreset { Label = "Yahoo (IMAP)", Protocol = "IMAP", Host = "imap.mail.yahoo.com", Port = 993, Sec = Security.SslTls,
                    Note = "Yahoo için uygulama şifresi gerekir." },
                new AccountPreset { Label = "iCloud (IMAP)", Protocol = "IMAP", Host = "imap.mail.me.com", Port = 993, Sec = Security.SslTls,
                    Note = "iCloud için uygulama özel parolası gerekir." },
                new AccountPreset { Label = "Outlook.com / Microsoft 365", Protocol = "IMAP", Host = "outlook.office365.com", Port = 993, Sec = Security.SslTls,
                    Note = "Microsoft parola ile IMAP girişine artık izin vermiyor. Bu hesaplar için klasik Outlook'u kullanın (üstteki Outlook seçeneği)." },
            };
        }
    }
}
