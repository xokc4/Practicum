using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum.BDExcel
{/// <summary>
/// объект все бд
/// </summary>
    internal class FullBd
    {
        public FullBd()
        {
        }
        //конструктор объекта
        public FullBd(List<Products> products, List<Orders> orders, List<Clients> clients)
        {
            this.products = products;
            this.orders = orders;
            this.clients = clients;
        }

        public List<Products> products { get; set; }// коллекция продуктов
        public List<Orders> orders { get; set; }//коллекция заказов
        public List<Clients> clients { get; set; }//коллекция клиентов
    }
}
