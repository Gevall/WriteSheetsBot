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
        //static List<Items> ItemsFromTable = GetData();


        static void Main(string[] args)
        {
            
            //GetData();
            StartBot();
        }

        /// <summary>
        /// Метода запуска бота
        /// </summary>
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
                            "\n 1) Адрес," +
                            "\n 2) Министерство," +
                            "\n 3) Кабинет," +
                            "\n 4) ФИО чей АРМ," +
                            "\n 5) Номер SSD," +
                            "\n 6) Номер АРМ," +
                            "\n 7) Статус," +
                            "\n 8) Причина если нельзя перевести на РедОС (\"-\" если всё ок)," +
                            "\n 9) Наличи СЗИ/Аттестации," +
                            "\n 10) ФИО Сотрудника," +
                            "\n 11) Запись в журнале (+ или -)," +
                            "\n 12) Необходимость докупить переходник (+ или -)," +
                            "\n 13) Описание (если есть)" +
                            "\n Все поля кроме \"Описания\" Обязательны! Если в поле нельзя ввести данные пишите: \"-\"" +
                            "\n Символами разделителями между полями являются символы \n, - Запятая \n; - Точка с запятой \n\\ - Обратный слеш  " +
                            "\n Если не разделите поля символами-разделителями то получите хуйню!");
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
            char[] separators = { ',', ';', '\\' };

            items = message.Split(separators);

            Items line = new Items();

            foreach (var e in items)
            {
                Console.WriteLine($">>>>>>>>> {e} >>>> {e.Length}");
            }

            try
            {
                        //line.Number = "123"; 
                        line.Address = items[0].Trim() == "-" ? null : items[0].Trim();
                        line.Ministry = items[1].Trim() == "-" ? null : items[1].Trim();
                        line.Cabinet = items[2].Trim() == "-" ? null : items[2].Trim();
                        line.NameOfEmployee = items[3].Trim() == "-" ? null : items[3].Trim();
                        line.NumberOfSSD = items[4].Trim() == "-" ? null : items[4].Trim();
                        line.NumberOfPC = items[5].Trim() == "-" ? null : items[5].Trim();
                        line.Status = items[6].Trim();
                        line.StatusDescription = items[7].Trim() == "-" ? null : items[7].Trim() ;
                        line.WorkDate = DateTime.Today.ToShortDateString();
                        line.AvailabilityOfSecurity = items[8] == "-" ? null : items[8].Trim();
                        line.NameOfComplete = items[9].Trim() == "-" ? null : items[9].Trim();
                        line.WriteOfJournal = items[10].Trim() == "+" ? "Да" : "Нет";  
                        line.NeedToByAdapter = items[11].Trim() == "+" ? "Да" : "Нет";
                        line.Caption = items[12].Trim();
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

            //WriteRandomString(items);
            //WriteRandomString(items);

            //return items;
        }

        static void WriteRandomString(List<Items> e)
        {
            Random random = new();
            int count = random.Next(e.Count);
            Console.WriteLine($"   {e[count].Address,15} |" +
                                $" {e[count].Ministry,10} |" +
                                $" {e[count].Cabinet,10} |" +
                                $" {e[count].NameOfEmployee,10} |" +
                                $" {e[count].NumberOfSSD,10} |" +
                                $" {e[count].NumberOfPC,10} |" +
                                $" {e[count].Status,10} |" +
                                $" {e[count].StatusDescription,10} |" +
                                $" {e[count].WorkDate,10} |" +
                                $" {e[count].AvailabilityOfSecurity,10} |" +
                                $" {e[count].NameOfComplete,10} |" +
                                $" {e[count].WriteOfJournal,10} |" +
                                $" {e[count].NeedToByAdapter,10} |" +
                                $" {e[count].Caption,10} |"
                                );
        }

        /// <summary>
        /// Вставка строки данных в таблицу
        /// </summary>
        static void PutData(Items item)
        {
            var range = $"{sheets_name}!A:O";

            var valueRange = new ValueRange
            {
                Values = ItemsMapper.MapToRangeData(item)
            };


            var appendRequest = valuesResource.Append(valueRange, sheets_id, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.Execute();
            #region
            Console.WriteLine($"Complete input string: " +
                                                    $" {item.Address,15} |" +
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
            #endregion

        }
    }
}