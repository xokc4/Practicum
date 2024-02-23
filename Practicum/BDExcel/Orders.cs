using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum.BDExcel
{
    /// <summary>
    /// объект для заказов
    /// </summary>
    internal class Orders
    {/// <summary>
    /// конструктор для заказа
    /// </summary>
    /// <param name="orderCode"></param>
    /// <param name="productCode"></param>
    /// <param name="clientCode"></param>
    /// <param name="orderNumber"></param>
    /// <param name="requiredQuantity"></param>
    /// <param name="orderDate"></param>
        public Orders(int orderCode, int productCode, int clientCode, int orderNumber, int requiredQuantity, DateTime orderDate)
        {
            OrderCode = orderCode;
            ProductCode = productCode;
            ClientCode = clientCode;
            OrderNumber = orderNumber;
            RequiredQuantity = requiredQuantity;
            OrderDate = orderDate;
        }

        public int OrderCode { get; set; }//код заявки
        public int ProductCode { get; set; }//код товара
        public int ClientCode { get; set; }//код клиента
        public int OrderNumber { get; set; }//номер заявки
        public int RequiredQuantity { get; set; }//количество
        public DateTime OrderDate { get; set; }//дата заказа
    }
}
