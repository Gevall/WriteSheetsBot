using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace WriteSheets
{
    internal class Program
    {
        const string sheets_id = "1GlQAIkh7oP6UtDfNSbRXDfQ6TtiHQEDzNtDYRSX9RQI";
        const string sheets_name = "Items";
        static GoogleServceHelper helper= new GoogleServceHelper();
        static SpreadsheetsResource.ValuesResource valuesResource = helper.Service.Spreadsheets.Values;

        static void Main(string[] args)
        {
            Console.WriteLine("Hello world");

            GetData();
            Console.WriteLine("\n//////////////////////////////////////////////////////////////////////////////////////////////\n");
            PutData();
            Console.WriteLine("\n//////////////////////////////////////////////////////////////////////////////////////////////\n");
            GetData();
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
        static void PutData()
        {
            var range = $"{sheets_name}!A:G";
            Items item = new()
            {
                Date = DateTime.Now.ToString(),
                Address = "ЗАГС",
                Cabinet = "каб 412",
                NumberOfPC = "222222",
                Status = "Не готово",
                Caption = "Нет питания HDD",
                NameOfComplete = "Оцебрик"
            };

            var valueRange = new ValueRange
            {
                Values = ItemsMapper.MapToRangeData(item)
            };

            var appendRequest = valuesResource.Append(valueRange, sheets_id, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();

            Console.WriteLine("Complete input");
        }
    }
}