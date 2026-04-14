/* * 功能：程式進入點
 * 對應選單名稱：系統引導
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Windows.Forms;

namespace MiniImageStudio {
    static class Program {
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
