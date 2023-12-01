using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.ReplyMarkups;
using WriteSheets.Bot;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace WriteSheets
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Создаем экзмепляр бота
            BotLogic bot = new(); 
            // Запускаем бота
            bot.StartBot();
        }
    }
}