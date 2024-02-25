using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Practicum;
using Practicum.BDExcel;
using System;
using System.Linq;


class Program
{
    public static FullBd FullBd = new FullBd();
    static void Main(string[] args)
    {
        FunctionalViewUp();


    }
    /// <summary>
    /// Метод для использования всего функционала приложения
    /// </summary>
    public static void FunctionalViewUp()
    {
        bool flag = true;
        Console.WriteLine("Введите пожалуйста абсолютный (полный) путь к файлу");
        string path= Console.ReadLine();//полный путь
        try
        {
            FullBd fullBd = WorkExcel.OpenExcel(path);
            FullBd = fullBd;
        }
        catch(Exception e )
        {
            Console.WriteLine($"Ошибка, {e}");
            flag = false;
        }
        while(flag)//вечный цикл
        {
            Console.WriteLine("Напишите 1 чтобы найти информацию о товаре, напишите 2 чтобы изменить имя клиента, " +
                "напишите 3 чтобы найти золотого клиента и любое число чтобы выйти");
            switch(Console.ReadLine())
            {
                case "1":
                    SearchView();
                break;
                case "2":
                    RedactClientsView(path);
                 break;
                case "3":
                    GoldClientsView();
                break;
                default:
                    Console.WriteLine("Пока");
                    Environment.Exit(0);
                    break;
            }
        }

        FunctionalViewUp();
    }
    /// <summary>
    /// View метод для поиска информации про товар
    /// </summary>
    public static void SearchView()
    {
        Console.WriteLine("Введите наименованию товара , чтобы вывести информацию о клиентах кто купил");
        string NameSearchProduct = Console.ReadLine();// название товара
        try
        {
            List<InformationOutput> ClientsByProduct = WorkingWithData.CustomerSearch(NameSearchProduct, FullBd);// Bakcend Коллекция с ответом
            if (ClientsByProduct.Count != 0)
            {
                foreach (InformationOutput output in ClientsByProduct)// перебор коллекции для вывода информации
                {
                    Console.WriteLine($"Организация: {output.Clients.OrganizationName}, имя клиента: {output.Clients.ContactPerson}," +
                        $" адресс клиента: {output.Clients.Address}. Количество товара: {output.Orders.RequiredQuantity}, " +
                        $"цена продукта: {output.Products.Price}, дата заказа: {output.Orders.OrderDate}");
                }
            }
            else
            {
                Console.WriteLine("Нет такого продукта");
            }
        }
        catch(Exception e)
        {
            Console.WriteLine($"нет такого продукта. Код ошибки {e}");
        }

    }
    /// <summary>
    /// View изменение данных в Excel Файла
    /// </summary>
    public static void RedactClientsView(string path)
    {
        Console.WriteLine("Введите Название организации");
        string NameOrganization =Console.ReadLine();
        Console.WriteLine("Введите новое имя клиента");
        string NameClient = Console.ReadLine();
        try
        {

           bool FlagSave= WorkExcel.SaveExcel(path, NameOrganization, NameClient);
        if (FlagSave)
            {
                Console.WriteLine("Изменение произошло");
            }
            else
            {
                Console.WriteLine("Нет такой организации");
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex}");
        }
    }
    /// <summary>
    /// View метод для нахождение золотого клиента
    /// </summary>
    public static void GoldClientsView()
    {
        try
        {


            Console.WriteLine("Введите год и месяц для определения золотого клиента");
            Console.WriteLine("Введите год");
            int year = Convert.ToInt32(Console.ReadLine());//переменная год
            Console.WriteLine("Введите месяц");
            int month = Convert.ToInt32(Console.ReadLine());// переменная месяц
            Clients clientsGold = WorkingWithData.GoldClients(year, month, FullBd);//Bakcend нахождение золотого клиента
            if (clientsGold != null)
            {
                Console.WriteLine($"Золотой клиент: {clientsGold.OrganizationName}, контактное лицо: {clientsGold.ContactPerson} На {month}/{year}");
            }
            else
            {
                Console.WriteLine("Нет заказов");
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine("Ввели не правильные данные ");
        }
    }
  
}
