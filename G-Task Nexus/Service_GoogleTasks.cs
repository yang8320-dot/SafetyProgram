using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Tasks.v1;
using Google.Apis.Tasks.v1.Data;
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
                    Scopes, "user", CancellationToken.None, 
                    new FileDataStore(tokenPath, true));

                return new TasksService(new Google.Apis.Services.BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = AppName
                });
            }
        }

        // 獲取所有任務
        public async Task<IList<Google.Apis.Tasks.v1.Data.Task>> GetAllTasksAsync()
        {
            var service = await InitializeServiceAsync();
            var list = await service.Tasks.List("@default").ExecuteAsync();
            return list.Items ?? new List<Google.Apis.Tasks.v1.Data.Task>();
        }

        // 新增任務並同步回 Google
        public async Task AddTaskAsync(string title)
        {
            var service = await InitializeServiceAsync();
            var newTask = new Google.Apis.Tasks.v1.Data.Task { Title = title };
            await service.Tasks.Insert(newTask, "@default").ExecuteAsync();
        }
    }
}
