using Practicum.BDExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum
{
    /// <summary>
    /// Класс для работы с данными из файла
    /// </summary>
    internal class WorkingWithData
    {
        
        public static List<Orders> ordersBD = new List<Orders>();
        public static List<Clients> clientsBD = new List<Clients>();
        public static List<Products> productsBD = new List<Products>();
        /// <summary>
        /// Метод для нахождения информации про клиентов и заказов по товару
        /// </summary>
        /// <param name="NameSearchProduct">имя продукта</param>
        /// <returns></returns>
        public static List<InformationOutput> CustomerSearch(string NameSearchProduct,FullBd fullBd)
        {
            ordersBD = fullBd.orders;
            clientsBD = fullBd.clients;
            productsBD = fullBd.products;
            List<InformationOutput> InfCliensAndOrders = new List<InformationOutput>();//лист для хронения  ответа
            Products productsOne = productsBD.FirstOrDefault(x => x.ProductName == NameSearchProduct);//Объект продукта кторый нашли с помощью имени
            List<Orders> OrdesSearch = ordersBD.Where(x => x.ProductCode == productsOne.ProductCode).ToList();// Нахождения листа с заказами по айди продукта
            
            foreach(var orders in OrdesSearch)//Цикл коллекции заказов для нахождения клиентов и запись информации в InfCliensAndOrders
            {
                Clients clients =  clientsBD.FirstOrDefault(x=>x.ClientCode==orders.ClientCode);//Клиент
                InfCliensAndOrders.Add(new InformationOutput(productsOne, clients, orders));//добавление
            }
            return InfCliensAndOrders;
        }
        /// <summary>
        /// Метод для нахождения "золотого клиента"
        /// </summary>
        /// <param name="year">год</param>
        /// <param name="month">месяц</param>
        /// <returns></returns>
        public static Clients GoldClients(int year, int month, FullBd fullBd)
        {
            ordersBD = fullBd.orders;
            clientsBD = fullBd.clients;
            productsBD = fullBd.products;
            var goldenCustomer = ordersBD
                .Where(order => order.OrderDate.Year == year && order.OrderDate.Month == month) // Фильтрация по году и месяцу
                .GroupBy(order => order.OrderCode) // Группировка заказов по OrderCode
                .OrderByDescending(group => group.Count()) // Сортировка по количеству заказов в группе (по убыванию)
                .FirstOrDefault(); // Получение первой группы (клиента) с наибольшим количеством заказов

            Orders ordersGold = goldenCustomer?.FirstOrDefault(); // Выбор первого заказа из группы или null, если группа пуста

            if (ordersGold != null) // Проверка на наличие заказа
            {
                Clients clientsGold = clientsBD.FirstOrDefault(x => x.ClientCode == ordersGold.ClientCode); // Нахождение клиента
                return clientsGold;
            }

            return null; // Если нет заказов в указанном месяце и году, то возвращаем null
        }
    }
}
