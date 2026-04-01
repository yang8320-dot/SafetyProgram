using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks; // 這是為了 async Task
using Google.Apis.Auth.OAuth2;
using Google.Apis.Tasks.v1;
using GData = Google.Apis.Tasks.v1.Data; // 設定別名，解決 Task 命名衝突
using Google.Apis.Util.Store;

namespace GTaskNexus
{
    public class GoogleTaskService
    {
        private static readonly string[] Scopes = { TasksService.Scope.Tasks };
        private const string AppName = "G-Task Nexus Portable";

        private async Task<TasksService> InitializeServiceAsync()
        {
            string jsonContent = CoreSecurity.LoadSecureData();
            if (string.IsNullOrEmpty(jsonContent)) 
                throw new Exception("找不到 API 憑證，請先於選單設定。");

            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonContent)))
            {
                // 使用本地目錄存儲 Token
                string tokenPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "auth_token");
                var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes, 
                    "user", 
                    CancellationToken.None, 
                    new FileDataStore(tokenPath, true));

                return new TasksService(new Google.Apis.Services.BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = AppName
                });
            }
        }

        /// <summary>
        /// 獲取所有任務 (明確指定使用 GData.Task)
        /// </summary>
        public async Task<IList<GData.Task>> GetAllTasksAsync()
        {
            var service = await InitializeServiceAsync();
            var listRequest = service.Tasks.List("@default");
            var result = await listRequest.ExecuteAsync();
            
            // 如果 Items 為 null，回傳空列表
            return result.Items ?? new List<GData.Task>();
        }

        /// <summary>
        /// 新增任務並同步回 Google
        /// </summary>
        public async Task AddTaskAsync(string title)
        {
            var service = await InitializeServiceAsync();
            var newTask = new GData.Task { Title = title };
            await service.Tasks.Insert(newTask, "@default").ExecuteAsync();
        }
    }
}
