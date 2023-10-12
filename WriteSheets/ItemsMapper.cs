﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteSheets
{
    internal class ItemsMapper
    {
        public static List<Items> MapFromRangeData(IList<IList<object>> values)
        {
            var items = new List<Items>();
            foreach (var value in values)
            {
                Items item = new()
                {
                    Date = value[0].ToString(),
                    Address = value[1].ToString(),
                    Cabinet = value[2].ToString(),
                    NumberOfPC = value[3].ToString(),
                    Status = value[4].ToString(),
                    Caption = value[5].ToString(),
                    NameOfComplete = value[6].ToString()
                };
                items.Add(item);
            }
            return items;
        }
        public static IList<IList<object>> MapToRangeData(Items item)
        {
            var objectList = new List<object>() { item.Date, item.Address, item.Cabinet, item.NumberOfPC, item.Status, item.Caption, item.NameOfComplete };
            var rangeData = new List<IList<object>> { objectList };
            return rangeData;
        }
    }
}