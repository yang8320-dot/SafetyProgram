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
        private const byte Key = 0x7A; // XOR 混淆金鑰

        // 取得目前資料夾中的隨機設定檔路徑
        public static string GetSecurePath()
        {
            return Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, Prefix + "*" + Ext).FirstOrDefault();
        }

        // 儲存加密資料 (如 API Credentials)
        public static void SaveSecureData(string rawContent)
        {
            // 刪除舊檔案
            string oldFile = GetSecurePath();
            if (oldFile != null) File.Delete(oldFile);

            // 產生新亂數檔名
            string newPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
                $"{Prefix}{Guid.NewGuid().ToString("N").Substring(0, 10)}{Ext}");

            byte[] data = Encoding.UTF8.GetBytes(rawContent);
            for (int i = 0; i < data.Length; i++) data[i] ^= Key; // 加密處理

            File.WriteAllBytes(newPath, data);
        }

        // 讀取並解密資料
        public static string LoadSecureData()
        {
            string path = GetSecurePath();
            if (path == null) return string.Empty;

            byte[] data = File.ReadAllBytes(path);
            for (int i = 0; i < data.Length; i++) data[i] ^= Key; // 解密處理

            return Encoding.UTF8.GetString(data);
        }
    }
}
