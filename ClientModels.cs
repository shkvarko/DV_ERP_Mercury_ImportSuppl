using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ERPMercury.WebAPI.Data.Classes
{
        /// <summary>
        /// Полная информация о клиенте
        /// </summary>
        public class Client: General
        {
            /// <summary>
            /// наименование клиента
            /// </summary>
            public string name { get; set; }

            /// <summary>
            /// уникальный числовой код клиента
            /// </summary>
            public int code { get; set; }

            /// <summary>
            /// УНП клиента
            /// </summary>
            public string UNP { get; set; }

            /// <summary>
            /// количество РТТ у клиента
            /// </summary>
            public int rttCount { get; set; }

            /// <summary>
            /// Список адресов клиента
            /// </summary>
            public IEnumerable<Address> addresses { get; set; }

            /// <summary>
            /// Список контактов клиента
            /// </summary>
            public IEnumerable<ContactInfo> contacts { get; set; }

            /// <summary>
            /// Список торговых точек клиента
            /// </summary>
            public IEnumerable<Rtt> rtts { get; set; }

            /// <summary>
            /// Список электронных адресов клиента
            /// </summary>
            public IEnumerable<Email> emails { get; set; }

            /// <summary>
            /// Список телефонов клиента
            /// </summary>
            public IEnumerable<Phone> phones { get; set; }

            /// <summary>
            /// Список банковских счетов
            /// </summary>
            public IEnumerable<BankAccount> bankAccounts { get; set; }

            /// <summary>
            /// Кредиторская задолженность клиента
            /// </summary>
            public IEnumerable<ClientDebt> credits { get; set; }

            /// <summary>
            /// Дебеторская задолженность клиента
            /// </summary>
            public IEnumerable<ClientDebt> debits { get; set; }

            /// <summary>
            /// Лимиты отпуска товара
            /// </summary>
            public IEnumerable<ClientLimit> limits { get; set; }

            /// <summary>
            /// признак клиента в должниках
            /// </summary>
            public bool clientInDebt { get; set; }
        }

        public class ClientShort 
        {
            /// <summary>
            /// идентификатор клиента
            /// </summary>
            public string clientId { get; set; }

            /// <summary>
            /// наименование клиента
            /// </summary>
            public string clientName { get; set; }

            /// <summary>
            /// признак клиента в должниках
            /// </summary>
            public bool clientInDebt { get; set; }
        }
}