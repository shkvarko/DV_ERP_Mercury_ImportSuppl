using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace ERPMercury.WebAPI.Data.Classes
{
    #region Класс "Состояние заказа"
    /// <summary>
    /// Состояние заказа
    /// </summary>
    public class OrderStatus
    {
        /// <summary>
        /// Код состояния заказа
        /// </summary>
        public string id { get; set; }

        /// <summary>
        /// Наименоване состояния заказа
        /// </summary>
        public string name { get; set; }

        /// <summary>
        /// цвет состояния
        /// </summary>
        public string color { get; set; }

        /// <summary>
        /// признак "жирного" шрифта
        /// </summary>
        public bool fontIsBold { get; set; }

        /// <summary>
        /// признак "зачеркивания" шрифта
        /// </summary>
        public bool fontIsCrossed { get; set; }
    }
    #endregion

    #region Класс "Тип заказа"
    /// <summary>
    /// Тип заказа
    /// </summary>
    public class OrderType
    {
        /// <summary>
        /// Идентификатор типа заказа
        /// </summary>
        public string id { get; set; }

        /// <summary>
        /// Наименоване типа заказа
        /// </summary>
        public string name { get; set; }
    }
    #endregion

    #region Класс "Итоговая сумма"
    /// <summary>
    /// Итоговая сумма
    /// </summary>
    public class TotalPrice
    {
        /// <summary>
        /// Сумма без скидки
        /// </summary>
        public decimal allPrice { get; set; }

        /// <summary>
        /// Сумма скидки
        /// </summary>
        public decimal allDiscount { get; set; }

        /// <summary>
        /// Сумма со скидкой
        /// </summary>
        public decimal totalPrice
        {
            get
            {
                return allPrice - allDiscount;
            }
        }
    }
    #endregion

    #region Класс "Цена"
    public class Price
    {
        /// <summary>
        /// Необходимое округление для цены
        /// </summary>
        public decimal priceRounding { get; set; }

        /// <summary>
        /// Цена без скидки
        /// </summary>
        public decimal price { get; set; }

        /// <summary>
        /// Скидка на цену в процентах
        /// </summary>
        public decimal discount { get; set; }

        /// <summary>
        /// Цена со скидкой и необходимым округлением
        /// </summary>
        public decimal discountPrice
        {
            get
            {
                return Math.Round(price * (1 - discount / 100) / priceRounding) * priceRounding;
            }
        }

        public Price()
        {
            this.price = 0;
            this.discount = 0;
            this.priceRounding = 1;
        }

    }
    #endregion

    #region Класс "Позиция в заказе"
    /// <summary>
    /// Единица продукции
    /// </summary>
    public class OrderItem
    {
        /// <summary>
        /// Идентификатор позиции в заказе
        /// </summary>
        public string id { get; set; }

        /// <summary>
        /// Товар
        /// </summary>
        public string productId { get; set; }

        /// <summary>
        /// Наименование товара
        /// </summary>
        public string productName { get; set; }

        /// <summary>
        /// артикул товара
        /// </summary>
        public string productArticle { get; set; }

        /// <summary>
        /// Цена в национальной валюте
        /// </summary>
        public Price price { get; set; }

        /// <summary>
        /// Заказанное количество
        /// </summary>
        public int orderQuantity { get; set; }

        /// <summary>
        /// Количество 
        /// </summary>
        public int quantity { get; set; }

        /// <summary>
        /// Сумма без скидки в национальной валюте
        /// </summary>
        public decimal allPrice
        {
            get
            {
                return price.price * quantity;
            }
        }

        /// <summary>
        /// Сумма со скидкой в национальной валюте
        /// </summary>
        public decimal totalPrice
        {
            get
            {
                return price.discountPrice * quantity;
            }
        }

        /// <summary>
        /// Сумма скидки в национальной валюте
        /// </summary>
        public decimal discountAllPrice
        {
            get
            {
                return (price.price - price.discountPrice) * quantity;
            }
        }


        public OrderItem()
        {
            this.price = new Price();
        }

    }
    #endregion

    #region Координаты
    /// <summary>
    /// Географические координаты
    /// </summary>
    public class GeoCoordinate
    {
        /// <summary>
        /// Широта
        /// </summary>
        public double? latitude { get; set; }

        /// <summary>
        /// Долгота
        /// </summary>
        public double? longitude { get; set; }
    }
    #endregion

    #region Класс "Заказ"
    /// <summary>
    /// Заказ
    /// </summary>
    public class Order : General
    {

        /// <summary>
        /// код "родительского" заказа ( null - заказ сам является "родителем"
        /// </summary>
        public string parentId { get; set; }

        /// <summary>
        /// Идентификатор подразделения
        /// </summary>
        public string departId { get; set; }

        /// <summary>
        /// Идентификатор клиента
        /// </summary>
        public string clientId { get; set; }

        /// <summary>
        /// Идентификатор РТТ
        /// </summary>
        public string rttId { get; set; }

        /// <summary>
        /// Идентификатор адреса доставки
        /// </summary>
        public string addressId { get; set; }

        /// <summary>
        /// дата создания заказа
        /// </summary>
        public DateTime beginDate { get; set; }

        /// <summary>
        /// Дата доставки
        /// </summary>
        public DateTime deliveryDate { get; set; }

        /// <summary>
        /// номер заказа
        /// </summary>
        public int orderNum { get; set; }

        /// <summary>
        /// версия заказа
        /// </summary>
        public int orderVersion { get; set; }

        /// <summary>
        /// Признак бонусного заказа
        /// </summary>
        public bool bonus { get; set; }

        /// <summary>
        /// Состояние заказа
        /// </summary>
        public OrderStatus orderStatus { get; set; }

        /// <summary>
        /// Идентификатор типа заказа
        /// </summary>
        public int typeId { get; set; }

        /// <summary>
        /// Сумма по заказу в национальной валюте
        /// </summary>
        public TotalPrice totalPrice { get; set; }

        /// <summary>
        /// Сумма по заказу в основной валюте
        /// </summary>
        public TotalPrice currencyTotalPrice { get; set; }

        /// <summary>
        /// Количество товара в заказе
        /// </summary>
        public int quantity { get; set; }

        /// <summary>
        /// Вес продукции в заказе
        /// </summary>
        public decimal weight { get; set; }

        /// <summary>
        /// признак возможности редактировать заказ
        /// </summary>
        public bool orderCanBeModified { get; set; }

        /// <summary>
        /// признак возможности удаления заказа
        /// </summary>
        public bool orderCanBeDeleted { get; set; }

        /// <summary>
        /// Примечание к заказу
        /// </summary>
        public string description { get; set; }

        /// <summary>
        /// Признак "черновик" для заказа
        /// </summary>
        public bool orderIsDraft { get; set; }

        /// <summary>
        /// Продукция по заказу
        /// </summary>
        public IEnumerable<OrderItem> items { get; set; }

        /// <summary>
        /// координаты, откуда был сохранен заказ
        /// </summary>
        public GeoCoordinate orderGeoPoint { get; set; }

        public Order()
        {
            items = new List<OrderItem>();
            orderStatus = new OrderStatus();
        }

    }
    #endregion

    #region Журнал событий
    /// <summary>
    /// класс для заказа с журналом событий
    /// </summary>
    public class OrderLog
    {
        /// <summary>
        /// уникальный код заказа
        /// </summary>
        public string orderId { get; set; }

        /// <summary>
        /// клиент
        /// </summary>
        public string clientName { get; set; }

        /// <summary>
        /// РТТ
        /// </summary>
        public string rttName { get; set; }

        /// <summary>
        /// адрес доставки
        /// </summary>
        public string addressName { get; set; }

        /// <summary>
        /// дата создания
        /// </summary>
        public DateTime orderBeginDate { get; set; }

        /// <summary>
        /// дата доставки
        /// </summary>
        public DateTime orderDeliveryDate { get; set; }

        /// <summary>
        /// количество товаров в заказе
        /// </summary>
        public int orderQuantity { get; set; }

        /// <summary>
        /// сумма по заказу
        /// </summary>
        public decimal orderTotalPrice { get; set; }

        /// <summary>
        /// состояние заказа
        /// </summary>
        public string orderState { get; set; }

        /// тип заказа
        /// </summary>
        public string orderType { get; set; }

        /// <summary>
        /// признак бонусного заказа
        /// </summary>
        public string orderBonus { get; set; }

        /// <summary>
        /// примечание к заказу
        /// </summary>
        public string orderNotes { get; set; }

        /// <summary>
        /// журнал событий по закау
        /// </summary>
        public IEnumerable<OrderLogItem> orderLogs { get; set; }


    }

    /// <summary>
    /// класс для элемента журнала по заказу
    /// </summary>
    public class OrderLogItem
    {
        /// <summary>
        /// дата события
        /// </summary>
        public DateTime orderLogDate { get; set; }

        /// <summary>
        /// категория события
        /// </summary>
        public string orderLogType { get; set; }

        /// <summary>
        /// описание события
        /// </summary>
        public string orderLogDescription { get; set; }
    }
    #endregion
}
