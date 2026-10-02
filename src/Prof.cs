using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace MailAktarici
{
    // Süre ölçümü: MAILAKTARICI_PROFIL ortam değişkeni varsa adım adım süreleri toplar.
    static class Prof
    {
        public static readonly bool On = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("MAILAKTARICI_PROFIL"));
        static readonly Dictionary<string, long> ticks = new Dictionary<string, long>();
        static readonly Dictionary<string, long> counts = new Dictionary<string, long>();

        public static long Start() { return On ? Stopwatch.GetTimestamp() : 0; }

        public static void Stop(string name, long start)
        {
            if (!On) return;
            long d = Stopwatch.GetTimestamp() - start;
            lock (ticks)
            {
                long t;
                ticks.TryGetValue(name, out t);
                ticks[name] = t + d;
                counts.TryGetValue(name, out t);
                counts[name] = t + 1;
            }
        }

        public static void Report(Reporter r)
        {
            if (!On) return;
            lock (ticks)
            {
                foreach (var kv in ticks.OrderByDescending(k => k.Value))
                {
                    double ms = kv.Value * 1000.0 / Stopwatch.Frequency;
                    r.Info("PROFIL " + kv.Key.PadRight(14) + ((long)ms).ToString("N0").PadLeft(9) + " ms  " + counts[kv.Key].ToString("N0").PadLeft(7) + " kez  " + (ms / counts[kv.Key]).ToString("0.0") + " ms/kez");
                }
            }
        }
    }
}
