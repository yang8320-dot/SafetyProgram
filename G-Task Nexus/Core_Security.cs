using System;
using System.IO;
using System.Linq;
using System.Text;

namespace GTaskNexus
{
    public static class CoreSecurity
    {
        private const string Prefix = "nexus_";
        private const string Ext = ".bin";
        private const byte Key = 0x7A;

        public static string GetSecurePath() => Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, Prefix + "*" + Ext).FirstOrDefault();

        public static void SaveSecureData(string rawContent)
        {
            string oldFile = GetSecurePath();
            if (oldFile != null) File.Delete(oldFile);
            string newPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{Prefix}{Guid.NewGuid().ToString("N").Substring(0, 10)}{Ext}");
            byte[] data = Encoding.UTF8.GetBytes(rawContent);
            for (int i = 0; i < data.Length; i++) data[i] ^= Key;
            File.WriteAllBytes(newPath, data);
        }

        public static string LoadSecureData()
        {
            string path = GetSecurePath();
            if (path == null) return string.Empty;
            byte[] data = File.ReadAllBytes(path);
            for (int i = 0; i < data.Length; i++) data[i] ^= Key;
            return Encoding.UTF8.GetString(data);
        }
    }
}
