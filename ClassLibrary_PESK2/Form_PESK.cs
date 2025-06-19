using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using Eplan.EplApi.MasterData;
using HtmlAgilityPack;
using System.Net.Http;

namespace ClassLibrary_PESK2
{
    public partial class Form_PESK : Form
    {
        public Form_PESK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /////////// Книга 1 ///////////
            int numRows = 13;
            int numCells = 2;

            IWorkbook workbook = new XSSFWorkbook();    // Создаем книгу Excel (.xlsx)
            ISheet sheet = workbook.CreateSheet("Лист1");   // Создаем лист, строки и ячейки

            // Стиль для жирного шрифта и размера + выравнивания по центру
            IFont boldFont = workbook.CreateFont();
            boldFont.IsBold = true;
            boldFont.FontHeightInPoints = 16; // Устанавливаем размер шрифта (в пунктах)
            ICellStyle boldStyle = workbook.CreateCellStyle();
            boldStyle.SetFont(boldFont);
            ICellStyle centerAlignedStyle = workbook.CreateCellStyle();
            centerAlignedStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            centerAlignedStyle.VerticalAlignment = VerticalAlignment.Center;
            centerAlignedStyle.SetFont(boldFont);

            // Стиль для выравнивания по центру + серая заливка + обводка
            ICellStyle styledCell = workbook.CreateCellStyle();
            styledCell.FillPattern = FillPattern.SolidForeground;
            styledCell.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            styledCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styledCell.VerticalAlignment = VerticalAlignment.Center;
            styledCell.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;      
            styledCell.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;   
            styledCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;    
            styledCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;   
            styledCell.TopBorderColor = IndexedColors.Black.Index; 
            styledCell.BottomBorderColor = IndexedColors.Black.Index; 
            styledCell.LeftBorderColor = IndexedColors.Black.Index; 
            styledCell.RightBorderColor = IndexedColors.Black.Index;

            // Стиль для обводки
            ICellStyle cellOutline = workbook.CreateCellStyle();
            cellOutline.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.TopBorderColor = IndexedColors.Black.Index;
            cellOutline.BottomBorderColor = IndexedColors.Black.Index;
            cellOutline.LeftBorderColor = IndexedColors.Black.Index;
            cellOutline.RightBorderColor = IndexedColors.Black.Index;
            cellOutline.WrapText = true;

            for (int rowNum = 0; rowNum < numRows; rowNum++)
            {
                IRow row = sheet.CreateRow(rowNum);
                for (int cellNum = 0; cellNum < numCells; cellNum++)
                {
                    NPOI.SS.UserModel.ICell cell = row.CreateCell(cellNum);
                    if (rowNum > 5)
                    {
                        cell.CellStyle = cellOutline;
                    }
                }
            }

            string dateString = "Дата: " + DateTime.Now.ToString("dd.MM.yyyy");
            sheet.GetRow(3).GetCell(0).SetCellValue(dateString);

            sheet.GetRow(5).GetCell(0).SetCellValue("Наименование");
            sheet.GetRow(5).GetCell(0).CellStyle = styledCell;
            sheet.GetRow(5).GetCell(1).SetCellValue("Значение");
            sheet.GetRow(5).GetCell(1).CellStyle = styledCell;

            sheet.GetRow(6).GetCell(0).SetCellValue(label1.Text); // Устанавливаем значение ячейки
            sheet.GetRow(6).GetCell(1).SetCellValue(Convert.ToString(comboBox1.SelectedItem));

            sheet.GetRow(7).GetCell(0).SetCellValue(label2.Text);
            sheet.GetRow(7).GetCell(1).SetCellValue(Convert.ToString(comboBox2.SelectedItem));

            sheet.GetRow(8).GetCell(0).SetCellValue(label3.Text);
            sheet.GetRow(8).GetCell(1).SetCellValue(Convert.ToString(comboBox3.SelectedItem));

            sheet.GetRow(9).GetCell(0).SetCellValue(label4.Text);
            sheet.GetRow(9).GetCell(1).SetCellValue(Convert.ToString(comboBox4.SelectedItem));

            sheet.GetRow(10).GetCell(0).SetCellValue(label7.Text);
            sheet.GetRow(10).GetCell(1).SetCellValue(Convert.ToString(comboBox5.SelectedItem));

            sheet.GetRow(11).GetCell(0).SetCellValue(label8.Text);
            sheet.GetRow(11).GetCell(1).SetCellValue(textBox1.Text + " кВт");

            sheet.GetRow(12).GetCell(0).SetCellValue(label9.Text);
            sheet.GetRow(12).GetCell(1).SetCellValue(textBox2.Text + " А");

            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1)); // объединение
            sheet.GetRow(1).GetCell(0).SetCellValue("Характеристики");
            sheet.GetRow(1).GetCell(0).CellStyle = centerAlignedStyle;

            sheet.SetColumnWidth(0, 29 * 256);
            sheet.SetColumnWidth(1, 20 * 256);


            /////////// Книга 2 ///////////





            // Получаем путь к папке "Документы" на рабочем столе
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderName = "Документы";
            string fullFolderPath = Path.Combine(desktopPath, folderName);

            // Проверяем, существует ли папка. Если нет - создаем
            if (!Directory.Exists(fullFolderPath))
            {
                Directory.CreateDirectory(fullFolderPath);
            }

            // Указываем полный путь к файлу
            string fileName = "myworkbook.xlsx";
            string filePath = Path.Combine(fullFolderPath, fileName);

            // Создаем файловый поток для записи данных в файл
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                // Записываем книгу в файловый поток
                workbook.Write(fileStream);
            }


            /////////// Достаем номер детали ///////////
            // Открыть базу данных
            MDPartsDatabase partsDatabase = null;

            try
            {
                MDPartsManagement oPartsManagement = new MDPartsManagement();
                partsDatabase = oPartsManagement.OpenDatabase();

                if (partsDatabase == null)
                {
                    MessageBox.Show("Не удалось открыть базу данных изделий!");
                    return;
                }

                MDPartsDatabaseItemPropertyList filter = new MDPartsDatabaseItemPropertyList();
                filter.ARTICLE_PRODUCTSUBGROUP = "1";
                filter.ARTICLE_PRODUCTGROUP = "5";

                MDPartsDatabaseItemPropertyList properties = new MDPartsDatabaseItemPropertyList();
                MDPart[] SubGroupParts = partsDatabase.GetParts(filter, properties);

                string partsInfo = "";

                if (SubGroupParts != null && SubGroupParts.Length > 0)
                {
                    foreach (MDPart part in SubGroupParts)
                    {
                        // Производитель
                        string manufacturer = part.Properties.ARTICLE_MANUFACTURER;
                        manufacturer = manufacturer.ToLower();
                        if (manufacturer == "veda")
                        {
                            partsInfo += $"Номер изделия: {part.PartNr}, Производитель:  {manufacturer}, ProductSubGroup: {part.ProductSubGroup}\r\n";
                        }
                    }
                }
                else
                {
                    partsInfo = "Не найдены изделия, соответствующие заданным критериям.";
                }

                MessageBox.Show("Изделия:\r\n" + partsInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message + "\r\n" + ex.StackTrace);
            }
            // Закрытие базы данных
            finally
            {
                if (partsDatabase != null)
                {
                    try
                    {
                        partsDatabase.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при закрытии базы данных: {ex.Message}");
                    }
                    finally
                    {
                        partsDatabase.Dispose();
                    }
                }
            }


            MessageBox.Show("Сгенерировано иди проверяй");
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            /////////// Пытаемся парсить сайты ///////////
            string article = "vf-51-p15k-0032-t4-e20-b-h"; // Здесь артикул для поиска.  Можно взять из TextBox, если нужно
            string price = await GetPriceFromEtm(article);

            if (!string.IsNullOrEmpty(price))
            {
                MessageBox.Show($"Цена для артикула {article}: {price}", "Цена");
            }
            else
            {
                MessageBox.Show($"Не удалось получить цену для артикула {article}.", "Ошибка");
            }


            /////////// Книга 2 ///////////
            int numRows = 15;
            int numCells = 6;

            IWorkbook workbook_2 = new XSSFWorkbook();    // Создаем книгу Excel (.xlsx)
            ISheet sheet_2 = workbook_2.CreateSheet("Лист1");   // Создаем лист, строки и ячейки

            // Стиль для жирного шрифта и размера + выравнивания по центру
            IFont boldFont = workbook_2.CreateFont();
            boldFont.IsBold = true;
            boldFont.FontHeightInPoints = 16; // Устанавливаем размер шрифта (в пунктах)
            ICellStyle boldStyle = workbook_2.CreateCellStyle();
            boldStyle.SetFont(boldFont);
            ICellStyle centerAlignedStyle = workbook_2.CreateCellStyle();
            centerAlignedStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            centerAlignedStyle.VerticalAlignment = VerticalAlignment.Center;
            centerAlignedStyle.SetFont(boldFont);
            centerAlignedStyle.WrapText = true;

            // Стиль для выравнивания по центру + серая заливка + обводка
            ICellStyle styledCell = workbook_2.CreateCellStyle();
            styledCell.FillPattern = FillPattern.SolidForeground;
            styledCell.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            styledCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styledCell.VerticalAlignment = VerticalAlignment.Center;
            styledCell.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            styledCell.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            styledCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            styledCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            styledCell.TopBorderColor = IndexedColors.Black.Index;
            styledCell.BottomBorderColor = IndexedColors.Black.Index;
            styledCell.LeftBorderColor = IndexedColors.Black.Index;
            styledCell.RightBorderColor = IndexedColors.Black.Index;
            styledCell.WrapText = true;

            // Стиль для обводки
            ICellStyle cellOutline = workbook_2.CreateCellStyle();
            cellOutline.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellOutline.TopBorderColor = IndexedColors.Black.Index;
            cellOutline.BottomBorderColor = IndexedColors.Black.Index;
            cellOutline.LeftBorderColor = IndexedColors.Black.Index;
            cellOutline.RightBorderColor = IndexedColors.Black.Index;
            cellOutline.WrapText = true;

            for (int rowNum = 0; rowNum < numRows; rowNum++)
            {
                IRow row = sheet_2.CreateRow(rowNum);
                for (int cellNum = 0; cellNum < numCells; cellNum++)
                {
                    NPOI.SS.UserModel.ICell cell = row.CreateCell(cellNum);
                    if (rowNum == 12) { cell.CellStyle = styledCell; }
                    else if (rowNum == 13) { cell.CellStyle = cellOutline; }
                }
            }

            try
            {
                string dateString = "Дата: " + DateTime.Now.ToString("dd.MM.yyyy");
                sheet_2.GetRow(8).GetCell(0).SetCellValue(dateString);

                sheet_2.GetRow(12).GetCell(0).SetCellValue("№ п/п");
                sheet_2.GetRow(12).GetCell(1).SetCellValue("Артикул");
                sheet_2.GetRow(12).GetCell(2).SetCellValue("Наименование");
                sheet_2.GetRow(12).GetCell(3).SetCellValue("Кол-во шт/упак.");
                sheet_2.GetRow(13).GetCell(3).SetCellValue("1");
                sheet_2.GetRow(12).GetCell(4).SetCellValue("Цена в руб.");
                sheet_2.GetRow(12).GetCell(5).SetCellValue("Наличие");

                sheet_2.AddMergedRegion(new CellRangeAddress(1, 1, 0, 5)); // объединение
                //sheet_2.AddMergedRegion(new CellRangeAddress(2, 3, 1, 1));
                //sheet_2.AddMergedRegion(new CellRangeAddress(4, 5, 1, 1));
                sheet_2.AddMergedRegion(new CellRangeAddress(2, 3, 2, 2));
                //sheet_2.AddMergedRegion(new CellRangeAddress(4, 5, 2, 2));
                sheet_2.GetRow(1).GetCell(0).SetCellValue("Коммерческое предложение");
                sheet_2.GetRow(1).GetCell(0).CellStyle = centerAlignedStyle;
                //sheet_2.GetRow(2).GetCell(1).SetCellValue("От:");
                //sheet_2.GetRow(3).GetCell(1).SetCellValue("Кому:");
                //sheet_2.GetRow(2).GetCell(1).CellStyle = cellOutline;
                //sheet_2.GetRow(3).GetCell(1).CellStyle = cellOutline;

                MessageBox.Show($"Цена: {price}", "Цена");
                if (price != "") { sheet_2.GetRow(13).GetCell(4).SetCellValue(price); }
                else { sheet_2.GetRow(13).GetCell(4).SetCellValue("--"); }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}");
                MessageBox.Show(ex.StackTrace); // Выводим трассировку стека для более детальной информации
            }

            sheet_2.SetColumnWidth(0, 4 * 256);
            sheet_2.SetColumnWidth(1, 18 * 256);
            sheet_2.SetColumnWidth(2, 38 * 256);
            sheet_2.SetColumnWidth(3, 11 * 256);
            sheet_2.SetColumnWidth(4, 11 * 256);
            sheet_2.SetColumnWidth(5, 11 * 256);

            // Получаем путь к папке "Документы" на рабочем столе
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderName = "Документы";
            string fullFolderPath = Path.Combine(desktopPath, folderName);

            // Проверяем, существует ли папка. Если нет - создаем
            if (!Directory.Exists(fullFolderPath))
            {
                Directory.CreateDirectory(fullFolderPath);
            }

            // Указываем полный путь к файлу
            string fileName = "myworkbook_2.xlsx";
            string filePath = Path.Combine(fullFolderPath, fileName);

            // Создаем файловый поток для записи данных в файл
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                // Записываем книгу в файловый поток
                workbook_2.Write(fileStream);
            }

            MessageBox.Show("Сгенерировано иди проверяй");
        }



        private async Task<string> GetPriceFromEtm(string article)
        {
            string baseUrl = "https://www.etm.ru/";
            string searchUrl = $"{baseUrl}catalog?searchValue={article}";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Отправляем GET запрос
                    HttpResponseMessage response = await client.GetAsync(searchUrl);
                    response.EnsureSuccessStatusCode(); // Проверка на успешный статус код (200 OK)
                    string htmlContent = await response.Content.ReadAsStringAsync();            

                    // Загружаем HTML в HtmlAgilityPack
                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                    htmlDocument.LoadHtml(htmlContent);

                    //  Здесь нужно определить XPath к элементу с ценой.
                    string priceXPath = "//p[@data-testid='catalog-list-item-price-details-0']";

                    HtmlNode priceNode = htmlDocument.DocumentNode.SelectSingleNode(priceXPath);

                    if (priceNode != null)
                    {
                        return priceNode.InnerText.Trim(); // Возвращаем текст (цену)
                    }
                    else
                    {
                        Console.WriteLine("Не удалось найти элемент с ценой по XPath.");
                        return null;
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                string message = $"Ошибка HTTP запроса: {ex.Message}";
                if (ex.InnerException != null)
                {
                    message += $"\nInner Exception: {ex.InnerException.Message}";
                }
                MessageBox.Show(message);

                //MessageBox.Show("Ошибка HTTP запроса: {ex.Message}");
                //Console.WriteLine($"Ошибка HTTP запроса: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}");
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
                return null;
            }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            /////////// Пытаемся парсить сайты ///////////
            string searchQuery = "VF-51-P15K-0032-T4-E20-B-H";

            try
            {
                string price = await GetPriceFromDrivesRu(searchQuery);

                if (!string.IsNullOrEmpty(price))
                {
                    MessageBox.Show($"Цена: {price}", "Цена товара");
                }
                else
                {
                    MessageBox.Show("Цена не найдена.", "Ошибка");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка");
            }
        }


        private async Task<string> GetPriceFromDrivesRu(string searchQuery)
        {
            string url = $"https://drives.ru/search/?query={searchQuery}";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string htmlContent = await response.Content.ReadAsStringAsync();

                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                    htmlDocument.LoadHtml(htmlContent);

                    // Ищем все элементы, содержащие информацию о товаре
                    HtmlNodeCollection productNodes = htmlDocument.DocumentNode.SelectNodes("//div[@class='s-products__item']");

                    if (productNodes != null)
                    {
                        foreach (HtmlNode productNode in productNodes)
                        {
                            // 1. Получаем название товара
                            HtmlNode nameNode = productNode.SelectSingleNode(".//div[@class='s-products__name']/a");  // Путь к названию товара
                            string productName = nameNode?.InnerText.Trim();

                            // 2. Проверяем, содержит ли название поисковый запрос (номер)
                            if (!string.IsNullOrEmpty(productName) && productName.Contains(searchQuery))
                            {
                                // 3. Если содержит, ищем цену в ЭТОМ элементе
                                HtmlNode priceNode = productNode.SelectSingleNode(".//div[@class='products__pr-price-new']//span[@class='price']");
                                if (priceNode != null)
                                {
                                    string priceText = priceNode.InnerText.Trim();
                                    return priceText;
                                }
                                else
                                {
                                    Console.WriteLine($"Цена не найдена для товара: {productName}");
                                    //  Можно также попробовать найти цену в других местах, если структура сайта отличается в зависимости от товара
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Название товара не содержит поисковый запрос: {productName}");
                            }
                        }
                        // Если прошли все товары и ничего не нашли:
                        Console.WriteLine($"Товар с названием, содержащим '{searchQuery}', не найден.");
                        return null;
                    }
                    else
                    {
                        Console.WriteLine("Не найдено элементов товаров на странице.");
                        return null;
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Ошибка HTTP: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при парсинге: {ex.Message}");
                return null;
            }
        }
    }
}
