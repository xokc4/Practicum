using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum.BDExcel
{
    /// <summary>
    /// объект где вся информация про товар
    /// </summary>
    internal class InformationOutput
    {

        public InformationOutput(Clients clients, Orders orders)
        {
            Clients = clients;
            Orders = orders;
        }
        /// <summary>
        /// конструктор для информации про товара
        /// </summary>
        /// <param name="products"></param>
        /// <param name="clients"></param>
        /// <param name="orders"></param>
        public InformationOutput(Products products, Clients clients, Orders orders)
        {
            Products = products;
            Clients = clients;
            Orders = orders;
        }

        public Products Products { get; set; }//  объект продукта
        public Clients Clients { get; set; }  //  объект клиента
        public Orders Orders { get; set; }//объект заказа
    }
}
