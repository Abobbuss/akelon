using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace akelon._data
{
    internal class DataService
    {
        private static List<string> GetDataByColumn(SpreadsheetDocument document, string sheetName, string columnName)
        {
            List<string> columnData = new List<string>();

            WorkbookPart workbookPart = document.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            if (sheet != null)
            {
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                Row headerRow = sheetData.Elements<Row>().FirstOrDefault();

                if (headerRow != null)
                {
                    int columnIndex = GetColumnIndex(workbookPart, headerRow, columnName);

                    if (columnIndex != -1)
                    {
                        foreach (Row row in sheetData.Elements<Row>().Skip(1))
                        {
                            Cell cell = row.Elements<Cell>().ElementAtOrDefault(columnIndex);
                            string cellValue = GetCellValue(workbookPart, cell);
                            columnData.Add(cellValue);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Столбец с названием \"{0}\" не найден.", columnName);
                    }
                }
                else
                {
                    Console.WriteLine("Не удалось найти строку заголовка.");
                }
            }
            else
            {
                Console.WriteLine("Лист с именем \"{0}\" не найден.", sheetName);
            }

            return columnData;
        }

        private static (int columnIndex, SheetData sheetData) GetColumnIndexAndSheetDataByColumnName(WorkbookPart workbookPart, string sheetName, string columnName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            if (sheet != null)
            {
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                Row headerRow = sheetData.Elements<Row>().FirstOrDefault();

                if (headerRow != null)
                {
                    int columnIndex = GetColumnIndex(workbookPart, headerRow, columnName);

                    return (columnIndex, sheetData);
                }
                else
                {
                    Console.WriteLine("Не удалось найти строку заголовка на листе \"{0}\".", sheetName);

                    return (-1, null);
                }
            }
            else
            {
                Console.WriteLine("Лист с именем \"{0}\" не найден.", sheetName);

                return (-1, null);
            }
        }

        private static int GetColumnIndex(WorkbookPart workbookPart, Row headerRow, string columnName)
        {
            int columnIndex = -1;

            foreach (Cell cell in headerRow.Elements<Cell>())
            {
                string headerText = GetCellValue(workbookPart, cell);

                if (headerText == columnName)
                {
                    columnIndex = headerRow.Elements<Cell>().ToList().IndexOf(cell);

                    break;
                }
            }

            return columnIndex;
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int index = int.Parse(value);
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (sharedStringTablePart != null)
                    value = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
            }

            return value;
        }

        private static string GetCellValueByColumnAndRow(WorkbookPart workbookPart, SheetData sheetData, int columnIndex, Row row)
        {
            if (columnIndex != -1 && sheetData != null)
            {
                if (row != null)
                {
                    Cell cell = row.Elements<Cell>().ElementAtOrDefault(columnIndex);

                    return GetCellValue(workbookPart, cell);
                }
                else
                {
                    Console.WriteLine("Строка не найдена на листе.");

                    return null;
                }
            }
            else
            {
                Console.WriteLine("Столбец не найден на листе.");

                return null;
            }
        }

        private static void SetCellValueByColumnAndRow(WorkbookPart workbookPart, string sheetName, string columnName, int rowIndex, string value)
        {
            var columnIndexAndSheetData = GetColumnIndexAndSheetDataByColumnName(workbookPart, sheetName, columnName);
            var columnIndex = columnIndexAndSheetData.columnIndex;
            var sheetData = columnIndexAndSheetData.sheetData;

            if (columnIndex != -1 && sheetData != null)
            {
                Row row = sheetData.Elements<Row>().ElementAtOrDefault(rowIndex);

                if (row != null)
                {
                    Cell cell = row.Elements<Cell>().ElementAtOrDefault(columnIndex);

                    if (cell == null)
                    {
                        cell = new Cell() { CellReference = GetColumnName(columnIndex) + rowIndex.ToString() };
                        row.InsertAt(cell, columnIndex);
                    }

                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);

                    workbookPart.Workbook.Save();
                }
                else
                {
                    Console.WriteLine("Строка с индексом {0} не найдена на листе \"{1}\".", rowIndex, sheetName);
                }
            }
            else
            {
                Console.WriteLine("Столбец с названием \"{0}\" не найден на листе \"{1}\".", columnName, sheetName);
            }
        }

        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static Dictionary<string, string> GetProductDataByProductCode(WorkbookPart workbookPart, SheetData productSheetData, int productCodeIndex, string productCode)
        {
            Dictionary<string, string> productData = new Dictionary<string, string>();

            var productPricePerUnitIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ProductColumns.ListName, ProductColumns.PricePerUnit).columnIndex;
            var productNameIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ProductColumns.ListName, ProductColumns.ProductName).columnIndex;

            foreach (var row in productSheetData.Elements<Row>().Skip(1))
            {
                string currentProductCode = GetCellValueByColumnAndRow(workbookPart, productSheetData, productCodeIndex, row);

                if (currentProductCode == productCode)
                {
                    productData[ProductColumns.ProductName] = GetCellValueByColumnAndRow(workbookPart, productSheetData, productNameIndex, row);
                    productData[ProductColumns.ProductCode] = currentProductCode;
                    productData[ProductColumns.PricePerUnit] = GetCellValueByColumnAndRow(workbookPart, productSheetData, productPricePerUnitIndex, row);

                    break;
                }
            }

            return productData.Count > 0 ? productData : null;
        }

        private static Dictionary<string, string> GetOrderDataByOrderCode(WorkbookPart workbookPart, SheetData orderSheetData, int orderCodeIndex, string orderCode)
        {
            Dictionary<string, string> orderData = new Dictionary<string, string>();

            var productCodeIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.ProductCode).columnIndex;
            var clientCodeIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.ClientCode).columnIndex;
            var quantityIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.RequiredQuantity).columnIndex;
            var dateIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.PlacementDate).columnIndex;

            foreach (var row in orderSheetData.Elements<Row>().Skip(1))
            {
                string currentOrderCode = GetCellValueByColumnAndRow(workbookPart, orderSheetData, orderCodeIndex, row);

                if (currentOrderCode == orderCode)
                {
                    orderData[OrderColumns.ProductCode] = GetCellValueByColumnAndRow(workbookPart, orderSheetData, productCodeIndex, row);
                    orderData[OrderColumns.ClientCode] = GetCellValueByColumnAndRow(workbookPart, orderSheetData, clientCodeIndex, row);
                    orderData[OrderColumns.RequiredQuantity] = GetCellValueByColumnAndRow(workbookPart, orderSheetData, quantityIndex, row);

                    string orderDateString = GetCellValueByColumnAndRow(workbookPart, orderSheetData, dateIndex, row);
                    DateTime? orderDate = null;

                    if (!string.IsNullOrEmpty(orderDateString))
                    {
                        double excelDateValue;

                        if (double.TryParse(orderDateString, out excelDateValue))
                            orderDate = DateTime.FromOADate(excelDateValue);
                        else
                            Console.WriteLine("Неверный формат даты.");
                    }

                    orderData[OrderColumns.PlacementDate] = orderDate?.ToString("dd.MM.yyyy");

                    break;
                }
            }

            return orderData.Count > 0 ? orderData : null;
        }

        private static string GetProductCodeByProductName(WorkbookPart workbookPart, SheetData productSheetData, string productName)
        {
            var productNameIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ProductColumns.ListName, ProductColumns.ProductName).columnIndex;
            var productCodeIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ProductColumns.ListName, ProductColumns.ProductCode).columnIndex;

            foreach (var row in productSheetData.Elements<Row>().Skip(1))
            {
                string currentProductName = GetCellValueByColumnAndRow(workbookPart, productSheetData, productNameIndex, row);

                if (currentProductName == productName)
                    return GetCellValueByColumnAndRow(workbookPart, productSheetData, productCodeIndex, row);
            }

            return null;
        }

        private static string GetClientNameByCode(WorkbookPart workbookPart, SheetData clientSheetData, string clientCode)
        {
            var clientCodeIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ClientColumns.ListName, ClientColumns.ClientCode).columnIndex;

            foreach (var row in clientSheetData.Elements<Row>().Skip(1))
            {
                string currentClientCode = GetCellValueByColumnAndRow(workbookPart, clientSheetData, clientCodeIndex, row);

                if (currentClientCode == clientCode)
                {
                    var clientOrganizationNameIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ClientColumns.ListName, ClientColumns.OrganizationName).columnIndex;

                    return GetCellValueByColumnAndRow(workbookPart, clientSheetData, clientOrganizationNameIndex, row);
                }
            }

            return null;
        }

        private static List<string> GetClientsWithMaxOrders(Dictionary<string, int> clientOrdersCount)
        {
            List<string> goldenClients = new List<string>();
            int maxOrders = 0;

            foreach (var kvp in clientOrdersCount)
            {
                if (kvp.Value > maxOrders)
                {
                    maxOrders = kvp.Value;
                    goldenClients.Clear();
                    goldenClients.Add(kvp.Key);
                }
                else if (kvp.Value == maxOrders)
                {
                    goldenClients.Add(kvp.Key);
                }
            }

            return goldenClients;
        }

        public static void GetOrderInfoByProductName(string filePath, string productName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                var (productNameIndex, productSheetData) = GetColumnIndexAndSheetDataByColumnName(workbookPart, ProductColumns.ListName, ProductColumns.ProductName);
                var (orderProductCodeIndex, orderSheetData) = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.ProductCode);
                var (orderClientIndex, _) = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.ClientCode);
                var (_, clientSheetData) = GetColumnIndexAndSheetDataByColumnName(workbookPart, ClientColumns.ListName, ClientColumns.ClientCode);
                var (_, clientOrganizationNameIndex) = GetColumnIndexAndSheetDataByColumnName(workbookPart, ClientColumns.ListName, ClientColumns.OrganizationName);

                string productCode = GetProductCodeByProductName(workbookPart, productSheetData, productName);
                Dictionary<string, string> productData = GetProductDataByProductCode(workbookPart, productSheetData, productNameIndex, productName);

                if (productData != null)
                {
                    string pricePerUnit = productData[ProductColumns.PricePerUnit];

                    foreach (Row orderRow in orderSheetData.Elements<Row>().Skip(1))
                    {
                        string orderProductCode = GetCellValueByColumnAndRow(workbookPart, orderSheetData, orderProductCodeIndex, orderRow);

                        if (orderProductCode == productCode)
                        {
                            string orderClientCode = GetCellValueByColumnAndRow(workbookPart, orderSheetData, orderClientIndex, orderRow);
                            Dictionary<string, string> orderData = GetOrderDataByOrderCode(workbookPart, orderSheetData, orderProductCodeIndex, orderProductCode);

                            if (orderData != null)
                            {
                                string clientName = GetClientNameByCode(workbookPart, clientSheetData, orderClientCode);

                                if (clientName != null)
                                {
                                    int quantity = int.Parse(orderData[OrderColumns.RequiredQuantity]);
                                    double price = quantity * double.Parse(pricePerUnit);
                                    string orderDateString = orderData[OrderColumns.PlacementDate];

                                    Console.WriteLine($"\nЗаказ на продукт: {productName}" +
                                        $"\nКоличество: {quantity}, Цена за единицу: {pricePerUnit}, Итоговая цена: {price}" +
                                        $" Дата: {orderDateString}, Клиент: {clientName}");
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Продукт с наименованием \"{productName}\" не найден.");

                    Console.WriteLine("\nДоступные продукты для поиска:");
                    List<string> allProducts = GetDataByColumn(spreadsheetDocument, ProductColumns.ListName, ProductColumns.ProductName);

                    foreach (var product in allProducts)
                    {
                        Console.WriteLine(product);
                    }
                }
            }
        }

        public static void ChangeClientContactPerson(string filePath, string organizationName, string newContactPerson)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                var (clientOrganizationNameIndex, clientSheetData) = GetColumnIndexAndSheetDataByColumnName(workbookPart, ClientColumns.ListName, ClientColumns.OrganizationName);

                if (clientOrganizationNameIndex != -1 && clientSheetData != null)
                {
                    bool clientFound = false; 

                    foreach (Row row in clientSheetData.Elements<Row>().Skip(1))
                    {
                        string currentOrganizationName = GetCellValueByColumnAndRow(workbookPart, clientSheetData, clientOrganizationNameIndex, row);

                        if (currentOrganizationName == organizationName)
                        {
                            clientFound = true;

                            var clientContactPersonIndex = GetColumnIndexAndSheetDataByColumnName(workbookPart, ClientColumns.ListName, ClientColumns.ContactPerson).columnIndex;

                            SetCellValueByColumnAndRow(workbookPart, ClientColumns.ListName, ClientColumns.ContactPerson, (int)row.RowIndex.Value - 1, newContactPerson);

                            Console.WriteLine($"Контактное лицо организации \"{organizationName}\" успешно изменено на \"{newContactPerson}\".");
                            
                            break; 
                        }
                    }

                    if (!clientFound)
                    {
                        Console.WriteLine($"\nОрганизация с названием \"{organizationName}\" не найдена.");

                        Console.WriteLine("\nДоступные организации:");
                        List<string> allOrganizations = GetDataByColumn(spreadsheetDocument, ClientColumns.ListName, ClientColumns.OrganizationName);

                        foreach (var organization in allOrganizations)
                        {
                            Console.WriteLine(organization);
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Лист с данными о клиентах \"{ClientColumns.ListName}\" не найден.");
                }
            }
        }

        public static List<string> GetGoldenClients(string filePath, int year, int? month = null)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                var (orderClientCodeIndex, orderSheetData) = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.ClientCode);
                var (orderDateIndex, _) = GetColumnIndexAndSheetDataByColumnName(workbookPart, OrderColumns.ListName, OrderColumns.PlacementDate);

                Dictionary<string, int> clientOrdersCount = new Dictionary<string, int>();

                if (orderClientCodeIndex != -1 && orderSheetData != null && orderDateIndex != -1)
                {
                    foreach (Row row in orderSheetData.Elements<Row>().Skip(1))
                    {
                        string orderClientCode = GetCellValueByColumnAndRow(workbookPart, orderSheetData, orderClientCodeIndex, row);
                        string orderDateString = GetCellValueByColumnAndRow(workbookPart, orderSheetData, orderDateIndex, row);

                        if (!string.IsNullOrEmpty(orderDateString))
                        {
                            double excelDateValue;

                            if (double.TryParse(orderDateString, out excelDateValue))
                            {
                                DateTime orderDate = DateTime.FromOADate(excelDateValue);

                                if (orderDate.Year == year && (!month.HasValue || orderDate.Month == month))
                                {
                                    if (clientOrdersCount.ContainsKey(orderClientCode))
                                        clientOrdersCount[orderClientCode]++;
                                    else
                                        clientOrdersCount[orderClientCode] = 1;
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Не удалось найти необходимые столбцы на листе заказов.");

                    return null;
                }

                List<string> goldenClients = GetClientsWithMaxOrders(clientOrdersCount);

                return goldenClients;
            }
        }
    }
}
