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

        public string? Number { get; set; } // Порядковый номер
        public string? Address { get; set; }    // Адрес проведения работ
        public string? Ministry { get; set; } // Министерство в котором выполнялись работы
        public string? Cabinet { get; set; } // номер кабинета 
        public string? NameOfEmployee { get; set; } // ФИО Сотрудника чей АРМ
        public string? NumberOfSSD { get; set; } // Номер SSD диска установленного в АРМ
        public string? NumberOfPC { get; set; } // номер АРМ
        public string? Status { get; set; } // Статус готовности АРМ
        public string? StatusDescription { get; set; } // Причина почему перевод невозможен
        public string? WorkDate { get; set; } // Дата выполнения работ
        public string? AvailabilityOfSecurity { get; set; } //Наличие СЗИ/Аттестации на АРМ
        public string? NameOfComplete { get; set; } // Имя выполнившего
        public string? WriteOfJournal { get; set; } // Запись в журнале
        public string? NeedToByAdapter { get; set; } // Необходимость докупить переходник
        public string? Caption { get; set; } // Примечание (если есть)

    }
}
