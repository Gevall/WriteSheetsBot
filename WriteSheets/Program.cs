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
        const string sheets_id = "1C_xjO7wjLiaeemnDmHlVNld-QOPYW2uHdePjVT9jwVo";
        const string sheets_name = "Items";
        static GoogleServceHelper helper= new GoogleServceHelper();
        static SpreadsheetsResource.ValuesResource valuesResource = helper.Service.Spreadsheets.Values;
        static ITelegramBotClient bot = new TelegramBotClient("6426754474:AAEXBSxN3aeTXgn98_5MRD-G5-G-BdzNVlA");

        static void Main(string[] args)
        {
            //StartBot();
            GetData();
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
                if (message.Sticker != null)
                {
                    await botClient.SendTextMessageAsync(message.Chat, "Не слать стикеры!! Метод не реализован!");
                    return;
                }
                if(message.Text.ToLower() is String)
                {
                    if (message.Text.ToLower() == "/help")
                    {
                        await botClient.SendTextMessageAsync(message.Chat, "Строка для добавления в таблицу должна иметь следующий вид:" +
                            "\nДата, Министерство, Местоположение АРМ, Номер АРМ, Статус АРМ, ФИО исполнителя/исполнителей, Описание");
                        return;
                    }
                    if (message.Text.ToLower() != null)
                    {
                        await WriteStringToGoogleSheets(message.Text);
                        return;
                    }
                }
                await botClient.SendTextMessageAsync(message.Chat, "Только текст!!111 Остальное не реализовано!");
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
                Items? item = MessageParser(message);
                if (item != null)
                {
                    PutData(item);
                }
                else Console.WriteLine("Хуйню какую то прислали");
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
                if (items.Length > 5)
                {
                    if (items.Length == 7)
                    {
                        //line.Date = items[0];
                        line.Address = items[1];
                        line.Ministry = items[2];
                        line.Cabinet = items[3];
                        line.NameOfEmployee = items[4];
                        line.NumberOfSSD = items[5];
                        line.NumberOfPC = items[6];
                        line.Status = items[7];
                        line.StatusDescription = items[8];
                        line.WorkDate = DateTime.MinValue.ToString();
                        line.AvailabilityOfSecurity = items[10];
                        line.NameOfComplete = items[11];
                        line.WriteOfJournal = items[12] == "+" ? "ИСТИНА" : "ЛОЖЬ";  
                        line.NeedToByAdapter = items[13] == "+" ? "ИСТИНА" : "ЛОЖЬ";
                        line.Caption = items[14];
                    }
                    else
                    {
                        //line.Date = items[0];
                        line.Address = items[1];
                        line.Cabinet = items[2];
                        line.NumberOfPC = items[3];
                        line.Status = items[4];
                        line.NameOfComplete = items[5];
                    }
                }
                else return null;
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
            var range = $"{sheets_name}!A:O";

            var request = valuesResource.Get(sheets_id, range);
            var response = request.Execute();
            var values = response.Values;

            var items = ItemsMapper.MapFromRangeData(values);

            foreach (var e in items )
            {
                Console.WriteLine($" {e.Address, 15} |" +
                    $" {e.Ministry,10} |" +
                    $" {e.Cabinet, 10} |" +
                    $" {e.NameOfEmployee, 10} |" +
                    $" {e.NumberOfSSD, 10} |" +
                    $" {e.NumberOfPC, 10} |" +
                    $" {e.Status, 10} |" +
                    $" {e.StatusDescription,10} |" +
                    $" {e.WorkDate,10} |" +
                    $" {e.AvailabilityOfSecurity,10} |" +
                    $" {e.NameOfComplete,10} |" +
                    $" {e.WriteOfJournal,10} |" +
                    $" {e.NeedToByAdapter,10} |" +
                    $" {e.Caption,10} |"
                    );
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

            Console.WriteLine($"Complete input string: {item.Address,15} |" +
                                                    $" {item.Ministry,10} |" +
                                                    $" {item.Cabinet,10} |" +
                                                    $" {item.NameOfEmployee,10} |" +
                                                    $" {item.NumberOfSSD,10} |" +
                                                    $" {item.NumberOfPC,10} |" +
                                                    $" {item.Status,10} |" +
                                                    $" {item.StatusDescription,10} |" +
                                                    $" {item.WorkDate,10} |" +
                                                    $" {item.AvailabilityOfSecurity,10} |" +
                                                    $" {item.NameOfComplete,10} |" +
                                                    $" {item.WriteOfJournal,10} |" +
                                                    $" {item.NeedToByAdapter,10} |" +
                                                    $" {item.Caption,10} |"
                                                    );
        }
    }
}