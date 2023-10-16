using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteSheets
{
    internal class ItemsMapper
    {
        /// <summary>
        /// Преобразование таблицы в List<Items>
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public static List<Items> MapFromRangeData(IList<IList<object>> values)
        {
            var items = new List<Items>();
            foreach (var value in values)
            {
                #region
                //Console.WriteLine(value[1]);
                //Console.WriteLine(value[2]);
                //Console.WriteLine(value[3]);
                //Console.WriteLine(value[4]);
                //Console.WriteLine(value[5]);
                //Console.WriteLine(value[6]);
                //Console.WriteLine(value[7]);
                //Console.WriteLine(value[8]);
                //Console.WriteLine(value[9]);
                //Console.WriteLine(value[10]);
                //Console.WriteLine(value[11]);
                //Console.WriteLine(value[12]);
                //Console.WriteLine(value[13]);
                ////Console.WriteLine(value[14]);
                //Console.WriteLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                #endregion
                try
                {
                    Items item = new()
                    {
                        CountOfLine =  value[0].ToString(),
                        Address = value[1] == null ? null : value[1].ToString(),
                        Ministry = value[2] == null ? null : value[2].ToString(),
                        Cabinet = value[3] == null ? null : value[3].ToString(),
                        NameOfEmployee = value[4] == null ? null : value[4].ToString(),
                        NumberOfSSD = value[5] == null ? null : value[5].ToString(),
                        NumberOfPC = value[6] == null ? null : value[6].ToString(),
                        Status = value[7] == null ? null : value[7].ToString(),
                        StatusDescription = value[8] == null ? null : value[8].ToString(),
                        WorkDate = value[9] == null ? null : value[9].ToString(),
                        AvailabilityOfSecurity = value[10] == null ? null : value[10].ToString(),
                        NameOfComplete = value[11] == null ? null : value[11].ToString(),
                        WriteOfJournal = value[12] == null ? null : value[12].ToString(),
                        NeedToByAdapter = value[13] == null ? null : value[13].ToString(),
                        Caption = value[14] == null ? null : value[14].ToString(),
                    };
                    items.Add(item);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return items;
        }

        /// <summary>
        /// конвертирование строки с новыми данными в табличную строку
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static IList<IList<object>> MapToRangeData(Items item)
        {
            var objectList = new List<object>() { 
                                                    item.Address, 
                                                    item.Ministry, 
                                                    item.Cabinet, 
                                                    item.NameOfEmployee, 
                                                    item.NumberOfSSD, 
                                                    item.NumberOfPC,
                                                    item.Status,
                                                    item.StatusDescription,
                                                    item.WorkDate,
                                                    item.AvailabilityOfSecurity,
                                                    item.NameOfComplete,
                                                    item.WriteOfJournal,
                                                    item.NeedToByAdapter,
                                                    item.Caption
                                                };
            var rangeData = new List<IList<object>> { objectList };
            return rangeData;
        }
    }
}
