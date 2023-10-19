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
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace WriteSheets
{
    internal class Program
    {
        const string sheets_id = "1C_xjO7wjLiaeemnDmHlVNld-QOPYW2uHdePjVT9jwVo";
        const string sheets_name = "Items";
        static GoogleServceHelper helper = new GoogleServceHelper();
        static SpreadsheetsResource.ValuesResource valuesResource = helper.Service.Spreadsheets.Values;
        static ITelegramBotClient bot = new TelegramBotClient("6426754474:AAEXBSxN3aeTXgn98_5MRD-G5-G-BdzNVlA");
        static Dictionary<long, Employes> users = new();
        static bool startProgram = false;

        static void Main(string[] args)
        {
            StartBot();
        }

        /// <summary>
        /// Метода запуска бота
        /// </summary>
        private static void StartBot()
        {
            if (System.IO.File.Exists("users.json") || startProgram)
            {
                readFromFileWorkersList();
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
                managmentConsole();
                //Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Файл с базой пользователей остутствует. Все равно запустить программу? (yes/no)");
                char input = Char.Parse(Console.ReadLine());
                if(input == 'y' || input == 'д') 
                {
                    startProgram = true;
                    StartBot();
                }
                else 
                {
                    Console.WriteLine("Программа закрыта. Проверьте файл и запустите повторно");
                }
            }
        }

        /// <summary>
        /// Логика бота
        /// </summary>
        /// <param name="botClient"></param>
        /// <param name="update"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            // Логика бота
            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(update));
            if (update.Type == Telegram.Bot.Types.Enums.UpdateType.Message)
            {
                var message = update.Message;
                if (message.Sticker != null)
                {
                    await botClient.SendTextMessageAsync(message.Chat, "Не слать стикеры!! Метод не реализован!");
                    return;
                }
                if (update.Message.Text is String)
                {
                    if (message.Text.ToLower() == "/help")
                    {
                        if (checkRegistation(update))
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
                                "\n 10) Запись в журнале (+ или -)," +
                                "\n 11) Необходимость докупить переходник (+ или -)," +
                                "\n 12) Описание (если есть)" +
                                "\n Все поля кроме \"Описания\" Обязательны! Если в поле нельзя ввести данные пишите: \"-\"" +
                                "\n Символами разделителями между полями являются символы \n, - Запятая \n; - Точка с запятой \n\\ - Обратный слеш  " +
                                "\n Если не разделите поля символами-разделителями то получите хуйню!");
                            return;
                        }
                        else
                        {
                            messateToRegister(update, botClient);
                            return;
                        }
                    }
                    if (message.Text.ToLower() == "/registration")
                    {
                        if (checkRegistation(update))
                        {
                            await botClient.SendTextMessageAsync(message.Chat, "Ты уже зарегистрирован! Иди работай!");
                            return;
                        }
                        else
                        {
                            await botClient.SendTextMessageAsync(message.Chat, "Введите ваше ФИО полностью", replyMarkup: new ForceReplyMarkup { Selective = true });
                            return;
                        }
                    }
                    if (message.Text == "/plsDelMyReg")
                    {
                        await botClient.SendTextMessageAsync(message.Chat, deleteEmploye(update));
                        return;
                    }
                    if (update.Message.ReplyToMessage != null && message.ReplyToMessage.Text.Contains("Введите ваше ФИО полностью"))
                    {
                        //users.Add(update.Message.Chat.Id, );
                        await botClient.SendTextMessageAsync(message.Chat, registation(update, message.Text));
                        return;
                    }
                    if (message.Text.ToLower() != null)
                    {
                        await writeStringToGoogleSheets(message.Text, update, botClient);
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

        /// <summary>
        /// Метод записи строки в таблицу
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        private static async Task writeStringToGoogleSheets(string message, Update update, ITelegramBotClient botclient)
        {
            bool chekerAddLine;
            lock (valuesResource)
            {
                Items? item = messageParser(message, update);
                if (item != null)
                {
                    putData(item);
                    chekerAddLine = true;
                }
                else
                {
                    Console.WriteLine("Хуйню какую то прислали");
                    chekerAddLine= false;
                }
            }

            if (chekerAddLine)
            {
                await botclient.SendTextMessageAsync(update.Message.Chat.Id, "Строка успешно добавлена");

            }
            else
            {
                await botclient.SendTextMessageAsync(update.Message.Chat.Id, "Неверный формат отправляемых данных!" +
                                                                             "\nПроверь отправляемую строку или воспользуйся помощью - /help");
            }
        }

        /// <summary>
        /// Консоль управления для вывода информации по боту
        /// </summary>
        private static void managmentConsole()
        {
            bool checker = true;
            while(checker)
            {
                if (Console.ReadLine().ToLower() == "show reg")
                {
                    foreach (var e in users)
                    {
                        Console.WriteLine($"id: {e.Key}, ФИО: {e.Value.lastname} {e.Value.firstname} {e.Value.patronymic}");
                    }
                }
                if (Console.ReadLine().ToLower() == "exit")
                {
                    checker = false;
                }
            }
        }

        /// <summary>
        /// Метод регистрации пользователя
        /// </summary>
        /// <param name="update"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        private static string registation(Update update, string message)
        {
            string[] data;
            char[] separatots = { ' ' };
            Employes newEmploye = new();

            data = message.Split(separatots);

            try
            {
                newEmploye.firstname = data[1];
                newEmploye.lastname = data[0];
                newEmploye.patronymic = data[2];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            users.Add(update.Message.Chat.Id, newEmploye);
            writeToFileRegistredWorkers();
            return "Вы успешно зарегистрированы! Можете отправлять данные";
        }

        /// <summary>
        /// Парсер сообщений в табличный вид
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        private static Items messageParser(string message, Update update)
        {
            string[] items;
            char[] separators = { ',', ';', '\\' };

            items = message.Split(separators);

            Items line = new Items();

            try
            {
                if (items.Length >= 10) 
                {
                    line.Address = items[0].Trim() == "-" ? null : items[0].Trim();
                    line.Ministry = items[1].Trim() == "-" ? null : items[1].Trim();
                    line.Cabinet = items[2].Trim() == "-" ? null : items[2].Trim();
                    line.NameOfEmployee = items[3].Trim() == "-" ? null : items[3].Trim();
                    line.NumberOfSSD = items[4].Trim() == "-" ? null : items[4].Trim();
                    line.NumberOfPC = items[5].Trim() == "-" ? null : items[5].Trim();
                    line.Status = items[6].Trim();
                    line.StatusDescription = items[7].Trim() == "-" ? null : items[7].Trim();
                    line.WorkDate = DateTime.Today.ToShortDateString();
                    line.AvailabilityOfSecurity = items[8].Trim() == "-" ? null : items[8].Trim();
                    line.Worker = insertWorkerData(update);
                    line.WriteOfJournal = items[9].Trim() == "+" ? "Да" : "Нет";
                    line.NeedToByAdapter = items[10].Trim() == "+" ? "Да" : "Нет";
                    line.Caption = items[11].Trim();
                }
                else
                {
                    return null;
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
        static void getData()
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

        /// <summary>
        /// Метод удаления регистрации пользователя
        /// </summary>
        /// <param name="update"></param>
        private static string deleteEmploye(Update update)
        {
            users.Remove(update.Message.Chat.Id);
            writeToFileRegistredWorkers();
            return "Ваша учетная запись удалена! Можете заного зарегистрироваться (если хотите) \"/registation\" ";
        }

        /// <summary>
        /// Метод преобразования ФИО сотурдника в строку для добавляния в Google таблицу
        /// </summary>
        /// <param name="update"></param>
        /// <returns></returns>
        private static string insertWorkerData(Update update)
        {
            string userDate;
            Employes worker = users[update.Message.Chat.Id];
            userDate = $"{worker.lastname} {worker.firstname} {worker.patronymic}";
            return userDate;
        }

        /// <summary>
        /// Проверка регистрации пользователя
        /// </summary>
        /// <param name="update"></param>
        /// <returns>Возвращает true - если пользователь зарегистрирвоан. Иначе - false</returns>
        private static bool checkRegistation(Update update)
        {
            long id = update.Message.Chat.Id;
            if (users.Keys.Contains(id)) return true;
            else return false;
        }

        /// <summary>
        /// Сообщение пользователю если он не зарегистирован
        /// </summary>
        /// <param name="update"></param>
        /// <param name="botClient"></param>
        private static async void messateToRegister(Update update, ITelegramBotClient botClient)
        {
            await botClient.SendTextMessageAsync(update.Message.Chat, "Вы не зарегистрированы! Для начала работы с ботом введите команду" +
                "\"/registation\" и зарегистрируйтесь!");
        }
        
        /// <summary>
        /// Серялизация зарегистрированных сторудников
        /// </summary>
        private static void writeToFileRegistredWorkers()
        {
            string json = JsonConvert.SerializeObject(users);
            lock (users)
            {
                System.IO.File.WriteAllText("users.json", json, System.Text.Encoding.UTF8);
            }
        }

        /// <summary>
        /// Десярилизация из json
        /// </summary>
        private static void readFromFileWorkersList()
        {
            try
            {
                string json = System.IO.File.ReadAllText("users.json");
                users = JsonConvert.DeserializeObject<Dictionary<long, Employes>>(json);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        static void _writeRandomString(List<Items> e)
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
                                $" {e[count].Worker,10} |" +
                                $" {e[count].WriteOfJournal,10} |" +
                                $" {e[count].NeedToByAdapter,10} |" +
                                $" {e[count].Caption,10} |"
                                );
        }

        /// <summary>
        /// Вставка строки данных в таблицу
        /// </summary>
        static void putData(Items item)
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
                                                    $" {item.Worker,10} |" +
                                                    $" {item.WriteOfJournal,10} |" +
                                                    $" {item.NeedToByAdapter,10} |" +
                                                    $" {item.Caption,10} |"
                                                    );
            #endregion

        }
    }
}