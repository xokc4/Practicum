using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Practicum.BDExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practicum
{
    /// <summary>
    /// класс который реализует интерфейс 
    /// </summary>
    internal class WorkExcel
    {
        public static List<Orders> ordersBD = new List<Orders>();
        public static List<Clients> clientsBD = new List<Clients>();
        public static List<Products> productsBD = new List<Products>();
        /// <summary>
        /// Метод который сохраняет изменения 
        /// </summary>
        /// <param name="path">путь</param>
        /// <param name="NameOrganization">имя организации</param>
        /// <param name="UpdateName">обнавленное имя</param>
        public static void SaveExcel(string path, string NameOrganization, string UpdateName)
        {
            ///Чтение файла
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet clientsSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Клиенты");//Нахождение листа по имени
                //условия существования листа
                if (clientsSheet != null)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(clientsSheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Цикл по всем строкам в листе "Клиенты"
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        //  название организации находится во второй ячейке каждой строки
                        Cell organizationCell = row.Elements<Cell>().ElementAtOrDefault(1);
                        string organizationName = GetCellValue(organizationCell, workbookPart);

                        // Если название организации совпадает с искомым
                        if (organizationName == NameOrganization)
                        {
                            // имя клиента находится в 4 ячейке каждой строки
                            Cell clientNameCell = row.Elements<Cell>().ElementAtOrDefault(3);

                            if (clientNameCell != null)
                            {
                                // Меняем значение ячейки с именем клиента
                                clientNameCell.CellValue = new CellValue(UpdateName);
                                clientNameCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }
                        }
                    }

                    workbookPart.Workbook.Save();//Сохраняем изменения 
                }

            }

        }
        /// <summary>
        /// Метод для чтение файла 
        /// </summary>
        /// <param name="path"></param>
        public static FullBd OpenExcel(string path)
        {
            //Открываем поток для чтения информации из файла
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                // Обрабатываем каждый лист в книге.
                foreach (Sheet sheet in workbookPart.Workbook.Descendants<Sheet>())
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    bool firstRow = true;//флаг для первой строчки
                    // Обрабатываем строки и ячейки в текущем листе.
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (firstRow)
                        {
                            firstRow = false;
                            continue;
                        }
                        if (sheet.Name == "Товары")//условие по названию листа
                        {
                            List<string> rowValues = new List<string>(); // Коллекция для хранения значений ячеек 
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = GetCellValue(cell, workbookPart);
                                rowValues.Add(cellValue); // Добавляем значение ячейки в коллекцию
                            }
                            if (rowValues[0] != "") // условие для пустой строчки
                            {
                                //добавляем данные в коллекцию  с продуктами
                                productsBD.Add(new Products(Convert.ToInt32(rowValues[0]), rowValues[1].ToString(), rowValues[2].ToString(), Convert.ToInt64(rowValues[3])));
                            }
                        }
                        if (sheet.Name == "Клиенты")
                        {
                            List<string> rowValues = new List<string>(); // Коллекция для хранения значений ячеек текущей строки

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = GetCellValue(cell, workbookPart);
                                rowValues.Add(cellValue); // Добавляем значение ячейки в коллекцию
                            }
                            if (rowValues[0] != "")
                            {//добавляем данные в коллекцию  с клиентами
                                clientsBD.Add(new Clients(Convert.ToInt32(rowValues[0]), rowValues[1].ToString(), rowValues[2].ToString(), rowValues[3].ToString()));
                            }

                        }
                        if (sheet.Name == "Заявки")
                        {
                            List<string> rowValues = new List<string>(); // Коллекция для хранения значений ячеек текущей строки

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = GetCellValue(cell, workbookPart);
                                rowValues.Add(cellValue); // Добавляем значение ячейки в коллекцию
                            }

                            if (rowValues[0] != "")// условие для пустой строчки
                            {
                                //переменная с датой 
                                DateTime date = ConvertDate(rowValues[5]);
                                //добавляем данные в коллекцию  с заказами
                                ordersBD.Add(new Orders(Convert.ToInt32(rowValues[0]), Convert.ToInt32(rowValues[1]),
                                Convert.ToInt32(rowValues[2]), Convert.ToInt32(rowValues[3]),
                                Convert.ToInt32(rowValues[4]), date));
                            }
                        }
                    }
                }
            }
            FullBd fullBd = new FullBd(productsBD,ordersBD,clientsBD);
            return fullBd;
        }
        /// <summary>
        /// Проверка текста
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                if (sharedStringTablePart != null)
                {
                    SharedStringItem sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.Text));
                    return sharedStringItem.Text.Text;
                }
            }
            else
            {
                return cell.CellValue?.Text.ToString() ?? string.Empty;
            }
            return string.Empty;
        }
        /// <summary>
        /// Метод для преобразования числа в дату
        /// </summary>
        /// <param name="Date"></param>
        /// <returns></returns>
        public static DateTime ConvertDate(string DateNumber)
        {

            DateTime baseDate = new DateTime(1900, 1, 1);

            // Преобразование числа в дату
            DateTime dateValue = baseDate.AddDays(Convert.ToInt32(DateNumber) - 2);
            return dateValue;
        }
    }
}
