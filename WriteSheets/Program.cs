using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace WriteSheets
{
    internal class Program
    {
        const string sheets_id = "1GlQAIkh7oP6UtDfNSbRXDfQ6TtiHQEDzNtDYRSX9RQI";
        const string sheets_name = "Items";
        static GoogleServceHelper helper= new GoogleServceHelper();
        static SpreadsheetsResource.ValuesResource valuesResource = helper.Service.Spreadsheets.Values;
        static ITelegramBotClient bot = new TelegramBotClient("6426754474:AAEXBSxN3aeTXgn98_5MRD-G5-G-BdzNVlA");

        static void Main(string[] args)
        {
            StartBot();
        }

        private static void StartBot()
        {
            Console.WriteLine("Запущен бот " + bot.GetMeAsync().Result.FirstName);

            var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { }, // receive all update types
            };
            bot.StartReceiving(
                HandleUpdateAsync,
                HandleErrorAsync,
                receiverOptions,
                cancellationToken
            );
            Console.ReadLine();
        }

        public static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            // Некоторые действия
            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(update));
            if (update.Type == Telegram.Bot.Types.Enums.UpdateType.Message)
            {
                var message = update.Message;
                if (message.Text.ToLower() == "/help")
                {
                    await botClient.SendTextMessageAsync(message.Chat, "Строка для добавления в таблицу должна иметь следующий вид:" +
                        "\nДата, Министерство, Местоположение АРМ, Номер АРМ, Статус АРМ, ФИО исполнителя/исполнителей, Описание");
                    return;
                }
                await WriteStringToGoogleSheets(message.Text);
            }
        }

        public static async Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            // Некоторые действия
            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(exception));
        }

        private static async Task WriteStringToGoogleSheets(string message)
        {
            lock(valuesResource)
            {
                Items item = MessageParser(message);
                PutData(item);
            }
        }

        private static Items MessageParser(string message)
        {
            string[] items;
            char[] separators = { ',' , ';', '\\' };

            items = message.Split(separators);

            Items line = new Items();
            int count = 0;
            foreach (var e in items)
            {
               
                Console.WriteLine($">>>>>> {count} " + e);
                count++;
            }

            try
            {
                if (items.Length == 7)
                {
                    line.Date = items[0];
                    line.Address = items[1];
                    line.Cabinet = items[2];
                    line.NumberOfPC = items[3];
                    line.Status = items[4];
                    line.NameOfComplete = items[5];
                    line.Caption = items[6];
                }
                else
                {
                    line.Date = items[0];
                    line.Address = items[1];
                    line.Cabinet = items[2];
                    line.NumberOfPC = items[3];
                    line.Status = items[4];
                    line.NameOfComplete = items[5];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return line;
        }

        /// <summary>
        /// Получение всех данных из таблицы
        /// </summary>
        static void GetData()
        {
            var range = $"{sheets_name}!A:G";

            var request = valuesResource.Get(sheets_id, range);
            var response = request.Execute();
            var values = response.Values;

            var items = ItemsMapper.MapFromRangeData(values);

            foreach (var e in items )
            {
                Console.WriteLine($"{e.Date, 10} | {e.Address, 15} | {e.Cabinet, 10} | {e.NumberOfPC, 10} | {e.Status, 10} | {e.Caption, 10} | {e.NameOfComplete, 10} ");
            }
        }

        /// <summary>
        /// Вставка строки данных в таблицу
        /// </summary>
        static void PutData(Items item)
        {
            var range = $"{sheets_name}!A:G";

            var valueRange = new ValueRange
            {
                Values = ItemsMapper.MapToRangeData(item)
            };

            var appendRequest = valuesResource.Append(valueRange, sheets_id, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();

            Console.WriteLine($"Complete input string: | {item.Date} | {item.Address} | {item.Cabinet} | {item.NumberOfPC} | {item.Status} | {item.NameOfComplete} | {item.Caption} ");
        }
    }
}