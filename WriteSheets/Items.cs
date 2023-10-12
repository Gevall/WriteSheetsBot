using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteSheets
{
    internal class Items
    {
        public Items() { }

        public string Date { get; set; }    // Дата выполнения работ
        public string Address { get; set; } // Министерство в котором выполнялись работы
        public string Cabinet { get; set; } // Местоположение АРМ в министерстве
        public string NumberOfPC { get; set; } // номер АРМ
        public string Status { get; set; } // Статус готовности АРМ
        public string NameOfComplete { get; set; } // Имя выполнившего
        public string Caption { get; set; } // Описание (если есть)

    }
}
