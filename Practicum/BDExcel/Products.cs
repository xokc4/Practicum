using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum.BDExcel
{/// <summary>
/// объект продуктов
/// </summary>
    internal class Products
    {
        public Products(int productCode, string productName, string unitOfMeasure, float price)
        {
            ProductCode = productCode;
            ProductName = productName;
            UnitOfMeasure = unitOfMeasure;
            Price = price;
        }

      public  int ProductCode { get; set; }//код продукта
        public string ProductName { get; set; }//имя продукта
        public string UnitOfMeasure { get; set; }//Ед измерения

        public float Price { get; set; }//цена продукта


    }
}
