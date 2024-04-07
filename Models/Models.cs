using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon_case_3.Models
{
    /// <summary>
    /// Модель товара
    /// </summary>
    internal class Product
    {
        /// <summary>
        /// Код товара
        /// </summary>
        public int Code { get; set; } = 0;
        /// <summary>
        /// Наименование товара
        /// </summary>
        public string Name { get; set; } = "";
        /// <summary>
        /// Единица измерения (товара)
        /// </summary>
        public string Unit { get; set; } = "";
        /// <summary>
        /// Цена за единицу товара
        /// </summary>
        public decimal Price { get; set; } = 0;
    }


    /// <summary>
    /// Модель клиента
    /// </summary>
    internal class Client
    {
        /// <summary>
        /// Код клиента
        /// </summary>
        public int Code { get; set; } = 0;
        /// <summary>
        /// Наименование организации
        /// </summary>
        public string OrganizationName { get; set; } = "";
        /// <summary>
        /// Адрес организации
        /// </summary>
        public string Address { get; set; } = "";
        /// <summary>
        /// Контактное лицо (ФИО)
        /// </summary>
        public string ContactPerson { get; set; } = "";
    }


    /// <summary>
    /// Модель заявки
    /// </summary>
    internal class Order
    {
        /// <summary>
        /// Код заявки
        /// </summary>
        public int OrderCode { get; set; } = 0;
        /// <summary>
        /// Код товара
        /// </summary>
        public int ProductCode { get; set; } = 0;
        /// <summary>
        /// Код клиента
        /// </summary>
        public int ClientCode { get; set; } = 0;
        /// <summary>
        /// Номер заявки
        /// </summary>
        public int OrderNumber { get; set; } = 0;
        /// <summary>
        /// Требуемое количество товара
        /// </summary>
        public int Quantity { get; set; } = 0;
        /// <summary>
        /// Дата размещения заявки
        /// </summary>
        public DateTime Date { get; set; } = DateTime.UtcNow;
    }
}
