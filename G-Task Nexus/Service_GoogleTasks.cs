using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Tasks.v1;
using GData = Google.Apis.Tasks.v1.Data;
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
            if (string.IsNullOrEmpty(jsonContent)) throw new Exception("請先設定 API 憑證內容。");

            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonContent)))
            {
                string tokenPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "auth_token");
                var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes, "user", CancellationToken.None, new FileDataStore(tokenPath, true));

                return new TasksService(new Google.Apis.Services.BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = AppName });
            }
        }

        public async Task<IList<GData.Task>> GetAllTasksAsync()
        {
            var service = await InitializeServiceAsync();
            var listRequest = service.Tasks.List("@default");
            listRequest.ShowCompleted = true; 
            listRequest.ShowHidden = true;
            var list = await listRequest.ExecuteAsync();
            return list.Items ?? new List<GData.Task>();
        }

        public async Task AddTaskAsync(string title)
        {
            var service = await InitializeServiceAsync();
            var newTask = new GData.Task { Title = title };
            await service.Tasks.Insert(newTask, "@default").ExecuteAsync();
        }

        public async Task UpdateTaskStatusAsync(string taskId, bool isCompleted)
        {
            var service = await InitializeServiceAsync();
            var task = await service.Tasks.Get("@default", taskId).ExecuteAsync();
            task.Status = isCompleted ? "completed" : "needsAction";
            await service.Tasks.Update(task, "@default", taskId).ExecuteAsync();
        }

        // 升級：同時更新標題與到期日 (Google API 限制只能存日期，忽略時間)
        public async Task UpdateTaskDetailsAsync(string taskId, string newTitle, DateTime? dueDate)
        {
            var service = await InitializeServiceAsync();
            var task = await service.Tasks.Get("@default", taskId).ExecuteAsync();
            task.Title = newTitle;
            
            if (dueDate.HasValue)
            {
                // 轉換為 Google 規定的 RFC 3339 格式，並將時間部分歸零
                task.Due = dueDate.Value.ToString("yyyy-MM-dd'T'00:00:00.000'Z'");
            }
            else
            {
                task.Due = null; // 清除到期日
            }

            await service.Tasks.Update(task, "@default", taskId).ExecuteAsync();
        }

        public async Task DeleteTaskAsync(string taskId)
        {
            var service = await InitializeServiceAsync();
            await service.Tasks.Delete("@default", taskId).ExecuteAsync();
        }
    }
}
