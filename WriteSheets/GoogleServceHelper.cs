using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteSheets
{
    internal class GoogleServceHelper
    {
        public SheetsService Service { get; set; } // Инициализация класса для работы с таблицой
        const string APP_NAME = "TestTable";   // имя таблицы
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        /// <summary>
        /// Инициализация при создании
        /// </summary>
        public GoogleServceHelper() 
        {
            InitializeService();
        }

        /// <summary>
        /// Инициализация подключения к API Google
        /// </summary>
        private void InitializeService()
        {
            var credential = GetCredentialsFromFile();
            Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APP_NAME
            });
        }

        /// <summary>
        /// Чтение json токена авторизации в API Google
        /// </summary>
        /// <returns></returns>
        private GoogleCredential GetCredentialsFromFile()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }
            return credential;
        }
    }
}
