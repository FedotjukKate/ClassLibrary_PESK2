﻿using System;
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
using Eplan.EplApi.Base;
using static ClassLibrary_PESK2.Form_PESK;
using Eplan.EplApi.Base.Internal;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static Eplan.EplApi.HEServices.Renumber.Enums;
using System.Text.RegularExpressions;
using System.Globalization;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using static Eplan.EplApi.EServices.IMessage;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using static Eplan.EplApi.Base.ISOCode;
using Eplan.EplApi.HEServices;
using System.Security.Cryptography;
using static Eplan.EplApi.DataModel.Properties;
using ExcelDataReader;



namespace ClassLibrary_PESK2
{
    public partial class Form_PESK : Form
    {
        private List<Device> devices = new List<Device>();
        private List<Filter> filters = new List<Filter>();
        private Device foundDevice;
        private bool flagNormal = true;

        public Form_PESK()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
            comboBox8.SelectedIndex = 0;
            comboBox9.SelectedIndex = 0;
            comboBox10.SelectedIndex = 0;
            comboBox11.SelectedIndex = 0;

            textBox25.Text = "C:\\Users\\Katya\\Desktop\\Учеба\\Спбпу\\Практика ПЭСК\\Нужное\\Переферийные и защитные устройства характеристики.xlsx";
            textBox26.Text = "C:\\Users\\Katya\\Desktop\\Учеба\\Спбпу\\Практика ПЭСК\\Нужное\\Рекомендуемые силовые опции.xlsx";

            checkBox1.Checked = true;
            comboBox6.Enabled = !checkBox1.Checked;
            checkBox1.CheckedChanged += (sender, e) => comboBox6.Enabled = !checkBox1.Checked;

            // Загрузка данных о ПЧ из базы
            LoadDataFromDatabase();
            //DisplayDeviceInfo();
            UpdateDeviceList();

            // События изменения
            comboBox1.SelectedIndexChanged += ParameterChanged;
            comboBox3.SelectedIndexChanged += ParameterChanged;
            comboBox5.SelectedIndexChanged += ParameterChanged;
            textBox1.TextChanged += ParameterChanged;
            textBox2.TextChanged += ParameterChanged;
            checkBox1.CheckedChanged += ParameterChanged;
        }

        //// Полезные кнопки ////
        private void button1_Click(object sender, EventArgs e)
        {
            string article = "";
            string discribe = "";
            string series = comboBox1.SelectedItem.ToString();
            string power = textBox1.Text.Trim();
            string current = textBox2.Text.Trim();
            string voltage = comboBox5.SelectedItem.ToString();

            filters.Clear();

            //Ищем ПЧ
            if (checkBox1.Checked)
            {
                // Проверка
                if (string.IsNullOrEmpty(power) || string.IsNullOrEmpty(current) || string.IsNullOrEmpty(voltage))
                {
                    MessageBox.Show("Пожалуйста, заполните все поля: мощность, ток и напряжение.", "Предупреждение");
                    return;
                }

                if (!float.TryParse(power, out float powerInt))
                {
                    MessageBox.Show("Некорректный формат мощности. Введите число.", "Предупреждение");
                    return;
                }

                if (!float.TryParse(current, out float currentInt))
                {
                    MessageBox.Show("Некорректный формат тока. Введите число.", "Предупреждение");
                    return;
                }

                // Идеальное
                if (flagNormal)
                {
                    foundDevice = devices.FirstOrDefault(d =>
                    d.Number.Contains(series) &&
                    d.Power == power &&
                    d.NormalCurrent == current &&
                    d.Voltage == voltage);
                }
                else
                {
                    foundDevice = devices.FirstOrDefault(d =>
                    d.Number.Contains(series) &&
                    d.Power == power &&
                    d.OverloadCurrent == current &&
                    d.Voltage == voltage);
                }

                try
                {
                    // Ближайшее
                    if (foundDevice == null)
                    {
                        if (flagNormal) { foundDevice = FindDeviceNormal(series, voltage, powerInt, currentInt); }
                        else { foundDevice = FindDeviceOverload(series, voltage, powerInt, currentInt); }
                    }

                    if (foundDevice != null)
                    {
                        article = foundDevice.Number;
                        discribe = foundDevice.Discribe;
                        MessageBox.Show($"Найден номер устройства: {foundDevice.Number}");
                    }
                    else
                    {
                        MessageBox.Show("Устройство с указанными параметрами не найдено.", "Ошибка");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка в заполнении базы данных: {ex.Message}", "Ошибка!");
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(comboBox6.SelectedItem.ToString()))
                {
                    article = comboBox6.SelectedItem.ToString();
                    foundDevice = devices.FirstOrDefault(d => d.Number == article);
                    discribe = foundDevice.Discribe;
                    MessageBox.Show($"Найден номер устройства: {article}");
                }
                else
                {
                    MessageBox.Show("Выберите устройство из списка или поставьте автоподбор.", "Предупреждение");
                }
            }



            //Ищем Предохранитель
            if (foundDevice != null)
            {
                Protection foundProtect = new Protection();
                string protectCurrent = FindProtectionCurrent(series, foundDevice.Voltage, foundDevice.Power);
                if (!string.IsNullOrEmpty(protectCurrent))
                {
                    MessageBox.Show($"Ток: {protectCurrent}");
                    foundProtect = FindProtection(protectCurrent);
                    if (foundProtect != null)
                    {
                        string deviceInfo = $"ПЧ\r\n Номер: {foundDevice.Number}, Напряжение: {foundDevice.Voltage},\r\n Мощность: {foundDevice.Power}," +
                            $" Ток выс перегрузки: {foundDevice.OverloadCurrent}, Ток норм перегрузки: {foundDevice.NormalCurrent}\r\n";
                        string protectInfo = $"Защита\r\n Номер: {foundProtect.Number}, Ток: {foundProtect.Current}\r\n";
                        MessageBox.Show("Информация об устройствах:\r\n" + deviceInfo + protectInfo);
                    }
                }
                else
                {
                    MessageBox.Show("Ток не найден или произошла ошибка.");
                }
            }

            //Ищем Фильтры
            if (foundDevice != null)
            {
                string input_filter = null;
                string input_throttle = null;
                string output_filter = null;
                string output_throttle = null;

                if (comboBox10.Text.Trim() != "Нет" || comboBox11.Text.Trim() != "Нет")
                {
                    input_filter = FindFilterNumber(1, series, foundDevice.Voltage, foundDevice.Power);
                    input_throttle = FindFilterNumber(2, series, foundDevice.Voltage, foundDevice.Power);
                    output_throttle = FindFilterNumber(3, series, foundDevice.Voltage, foundDevice.Power);
                    output_filter = FindFilterNumber(4, series, foundDevice.Voltage, foundDevice.Power);

                    FindFilter(input_filter, input_throttle, output_filter, output_throttle);

                    string deviceInfo = $"ПЧ\r\n Номер: {foundDevice.Number}, Напряжение: {foundDevice.Voltage},\r\n Мощность: {foundDevice.Power}," +
                            $" Ток выс перегрузки: {foundDevice.OverloadCurrent}, Ток норм перегрузки: {foundDevice.NormalCurrent}\r\n";
                    string filterInfo = "";
                    foreach (Filter filter in filters)
                    {
                        filterInfo += $"Фильтр Номер: {filter.Number}\r\n";
                    }
                    MessageBox.Show("Информация об устройствах:\r\n" + deviceInfo + filterInfo);
                }
            }


            if (!string.IsNullOrEmpty(textBox21.Text) && !string.IsNullOrEmpty(textBox22.Text)
                && !string.IsNullOrEmpty(textBox23.Text) && !string.IsNullOrEmpty(textBox24.Text))
            {
                button2.Enabled = true;
                button3.Enabled = true;

                string projectName = textBox21.Text;

                // путь к папке для хранения
                string documentsPath = textBox22.Text;
                string projectDirectory = Path.Combine(documentsPath, projectName);
                // Путь к проекту полный
                string projectPath = Path.Combine(projectDirectory, projectName + ".elk");

                // Проверяем, существует ли папка. Если нет - создаем
                if (!Directory.Exists(projectDirectory))
                {
                    Directory.CreateDirectory(projectDirectory);
                }

                // EPLAN
                try
                {
                    string templatesPath = textBox23.Text;

                    using (SafetyPoint safetyPoint = SafetyPoint.Create())
                    {
                        // Cоздание проекта
                        ProjectManager projectManager = new ProjectManager();
                        Eplan.EplApi.DataModel.Project oProject = projectManager.CreateProject(projectPath, templatesPath);

                        //Create_Pages(oProject);

                        if (oProject != null)
                        {
                            MessageBox.Show($"Проект EPLAN \"{projectName}\" успешно создан в папке: {projectDirectory}", "Успешно");
                        }
                        else
                        {
                            MessageBox.Show($"Не удалось создать проект EPLAN \"{projectName}\"", "Ошибка");
                        }
                        safetyPoint.Commit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при создании проекта EPLAN: {ex.Message}", "Ошибка!");
                }
            }
            else
            {
                MessageBox.Show("Заполните все поля вкладки 'Инженер'. ", "Предупреждение");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Кнопка 2", "Предупреждение");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Кнопка 3", "Предупреждение");
            button4.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Кнопка 4", "Предупреждение");
        }

        //// Кнопки поиска ////
        private void button21_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog of = new FolderBrowserDialog();
            of.Description = "Выберите папку для сохранения проекта:";
            if (of.ShowDialog() == DialogResult.OK)
            {
                string folderPath = of.SelectedPath;
                textBox22.Text = folderPath;
                textBox22.SelectionStart = textBox21.Text.Length;
                textBox22.SelectionLength = 0;
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Файлы EPLAN|*.ept;*.epb;*.zw9";
            if (of.ShowDialog() == DialogResult.OK)
            {
                string filePath = of.FileName;
                textBox23.Text = filePath;
                textBox23.SelectionStart = textBox1.Text.Length;
                textBox23.SelectionLength = 0;
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            var of = new FolderPicker();
            of.Multiselect = false;
            if (of.ShowDialog(IntPtr.Zero) == true)
            {
                string folderPath = of.ResultPath;
                textBox24.Text = folderPath;
                textBox24.SelectionStart = textBox21.Text.Length;
                textBox24.SelectionLength = 0;
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Файлы Excel|*.xls;*.xlsx;*.xlsm";
            if (of.ShowDialog() == DialogResult.OK)
            {
                string filePath = of.FileName;
                textBox25.Text = filePath;
                textBox25.SelectionStart = textBox25.Text.Length;
                textBox25.SelectionLength = 0;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Файлы Excel|*.xls;*.xlsx;*.xlsm";
            if (of.ShowDialog() == DialogResult.OK)
            {
                string filePath = of.FileName;
                textBox26.Text = filePath;
                textBox26.SelectionStart = textBox26.Text.Length;
                textBox26.SelectionLength = 0;
            }
        }

        //// Бесполезные вещи ////
        /////private void button2_Click(object sender, EventArgs e)
        //{
        //    string projectName = "NewEplanProject";

        //    // путь к папке "Документы"
        //    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //    string documentsPath = Path.Combine(desktopPath, "Документы");
        //    string projectDirectory = Path.Combine(documentsPath, projectName);

        //    // Путь к проекту полный
        //    string projectPath = Path.Combine(projectDirectory, projectName + ".elk");


        //    // Проверяем, существует ли папка. Если нет - создаем
        //    if (!Directory.Exists(projectDirectory))
        //    {
        //        Directory.CreateDirectory(projectDirectory);
        //    }

        //    // EPLAN
        //    try
        //    {
        //        string templatesPath = "$(MD_TEMPLATES)\\IEC_bas003.zw9";  

        //        using (SafetyPoint safetyPoint = SafetyPoint.Create())
        //        {
        //            // Cоздание проекта
        //            ProjectManager projectManager = new ProjectManager();
        //            Project oProject = projectManager.CreateProject(projectPath, templatesPath);

        //            if (oProject != null)
        //            {
        //                MessageBox.Show($"Проект EPLAN \"{projectName}\" успешно создан в папке: {projectDirectory}", "Успешно");
        //            }
        //            else
        //            {
        //                MessageBox.Show($"Не удалось создать проект EPLAN \"{projectName}\"", "Ошибка");
        //            }
        //            safetyPoint.Commit();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Ошибка при создании проекта EPLAN: {ex.Message}", "Ошибка!");
        //    }
        //}


        //private async void button0_Click(object sender, EventArgs e)
        //{
        //    string article = "";
        //    string discribe = "";
        //    string series = comboBox1.SelectedItem.ToString();
        //    string power = textBox1.Text.Trim();
        //    string current = textBox2.Text.Trim();
        //    string voltage = comboBox5.SelectedItem.ToString();

        //    if (checkBox1.Checked)
        //    {
        //        // Проверка
        //        if (string.IsNullOrEmpty(power) || string.IsNullOrEmpty(current) || string.IsNullOrEmpty(voltage))
        //        {
        //            MessageBox.Show("Пожалуйста, заполните все поля: мощность, ток и напряжение.", "Предупреждение");
        //            return;
        //        }

        //        if (!float.TryParse(power, out float powerInt))
        //        {
        //            MessageBox.Show("Некорректный формат мощности. Введите число.", "Предупреждение");
        //            return;
        //        }

        //        if (!float.TryParse(current, out float currentInt))
        //        {
        //            MessageBox.Show("Некорректный формат тока. Введите число.", "Предупреждение");
        //            return;
        //        }

        //        // Идеальное
        //        foundDevice = devices.FirstOrDefault(d =>
        //            d.Number.Contains(series) &&
        //            d.Power == power &&
        //            d.Current == current &&
        //            d.Voltage == voltage);

        //        // Ближайшее
        //        if (foundDevice == null)
        //        {
        //            foundDevice = FindDevice(series, voltage, powerInt, currentInt);
        //        }

        //        if (foundDevice != null)
        //        {
        //            article = foundDevice.Number;
        //            discribe = foundDevice.Discribe;
        //            //MessageBox.Show($"Найден номер устройства: {foundDevice.Number}");
        //        }
        //        else
        //        {
        //            MessageBox.Show("Устройство с указанными параметрами не найдено.", "Ошибка");
        //        }
        //    }
        //    else
        //    {
        //        if (!string.IsNullOrEmpty(comboBox6.SelectedItem.ToString()))
        //        {
        //            article = comboBox6.SelectedItem.ToString();
        //            foundDevice = devices.FirstOrDefault(d => d.Number == article);
        //            discribe = foundDevice.Discribe;
        //            //MessageBox.Show($"Найден номер устройства: {article}");
        //        }
        //        else
        //        {
        //            MessageBox.Show("Выберите устройство из списка или поставьте автоподбор.", "Предупреждение");
        //        }
        //    }

        //    /////////// Пытаемся парсить сайты ///////////
        //    string price = await GetPriceFromDrivesRu(article);
        //    if (!string.IsNullOrEmpty(article))
        //    {
        //        if (!string.IsNullOrEmpty(price))
        //        {
        //            //MessageBox.Show($"Цена для артикула {article}: {price}", "Цена");
        //        }
        //        else
        //        {
        //            price = "Нет на складе";
        //            MessageBox.Show($"Не удалось получить цену для артикула {article}.", "Предупреждение");
        //        }

        //        /////////// Книга 1 ///////////
        //        try
        //        {
        //            CreateBook();  // Функция создания книги 1
        //            MessageBox.Show("Файл 1 успешно сохранен", "Уведомление");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Произошла ошибка при создании Excel-файла: " + ex.Message);
        //        }

        //        /////////// Книга 2 ///////////
        //        try
        //        {
        //            CreateBook_2(price, article, discribe);  // Функция создания книги 2
        //            MessageBox.Show("Файл 2 успешно сохранен", "Уведомление");
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Произошла ошибка при создании Excel-файла: " + ex.Message);
        //        }
        //    }
        //}

        private void Create_Pages(Eplan.EplApi.DataModel.Project oProject)
        {
            //Установка иерархии:
            PagePropertyList name_parts = new PagePropertyList();
            name_parts.DESIGNATION_USERDEFINED.Set("КД");
            //name_parts.DESIGNATION_INSTALLATIONNUMBER.Set("DESIGNATION_INSTALLATIONNUMBER");
            //name_parts.DESIGNATION_FUNCTIONALASSIGNMENT.Set("DESIGNATION_FUNCTIONALASSIGNMENT");
            //name_parts.DESIGNATION_PLANT.Set("DESIGNATION_PLANT");
            //name_parts.DESIGNATION_PLACEOFINSTALLATION.Set("DESIGNATION_PLACEOFINSTALLATION");
            name_parts.DESIGNATION_LOCATION.Set("ШУ-ПЧ");
            name_parts.DESIGNATION_DOCTYPE.Set("ТЛ");

            // Создание страницы
            try { name_parts.PAGE_COUNTER.Set(1); }
            catch (Exception ex) { MessageBox.Show($"Ошибка при создании иерархии: {ex.Message}", "Ошибка!"); }
            Eplan.EplApi.DataModel.Page Page_TL = new Eplan.EplApi.DataModel.Page();
            Page_TL.Create(oProject, DocumentTypeManager.DocumentType.TitlePage, name_parts);
            Page_TL.Properties.PAGE_FORMPLOT.Set("ЕСКД_A4_Титульный_лист_v3.0"); 
            try { Page_TL.Properties[11011] = "Титульный лист"; }
            catch { }

            //name_parts.DESIGNATION_DOCTYPE.Set("СП");
            //try { name_parts.PAGE_COUNTER.Set(1); }
            //catch (Exception ex) { MessageBox.Show($"Ошибка при создании иерархии: {ex.Message}", "Ошибка!"); }
            //Page Page_SP_1 = new Page();
            //Page_SP_1.Create(oProject, DocumentTypeManager.DocumentType.Circuit, name_parts);
            //Page_SP_1.Properties.PAGE_FORMULAR.Set("F28_002_en_US");
            //Page_SP_1.Properties.PAGE_FORMPLOT.Set("GOST_A4_first_page_text");
            //try { Page_SP_1.Properties[11011] = "Ведомость документов"; }
            //catch { }

            //try { name_parts.PAGE_COUNTER.Set(2); }
            //catch (Exception ex) { MessageBox.Show($"Ошибка при создании иерархии: {ex.Message}", "Ошибка!"); }
            //Page Page_SP_2 = new Page();
            //Page_SP_2.Create(oProject, DocumentTypeManager.DocumentType.Circuit, name_parts);
            //Page_SP_2.Properties.PAGE_FORMULAR.Set("F28_002_en_US");
            //Page_SP_2.Properties.PAGE_FORMPLOT.Set("GOST_A4_first_page_text");
            //try { Page_SP_2.Properties[11011] = "Спецификация"; }
            //catch { }

            name_parts.DESIGNATION_DOCTYPE.Set("ЭЗ");
            try { name_parts.PAGE_COUNTER.Set(1); }
            catch (Exception ex) { MessageBox.Show($"Ошибка при создании иерархии: {ex.Message}", "Ошибка!"); }
            Eplan.EplApi.DataModel.Page Page_SP_1 = new Eplan.EplApi.DataModel.Page();
            Page_SP_1.Create(oProject, DocumentTypeManager.DocumentType.Circuit, name_parts);
            //Page_SP_1.Properties.PAGE_FORMPLOT = "ЕСКД_A3_Форма_1_v3.0";
            Page_SP_1.Properties.PAGE_FORMPLOT.Set("ЕСКД_A3_Форма_1_v3.0");
            
            //Page_SP_1.Properties.set_PAGE_FORMPLOT(1, "ЕСКД_A3_Форма_1_v3.0");
            try { Page_SP_1.Properties[11011] = "Ввод 660VAC"; }
            catch { }


            MessageBox.Show("iiiiiiiiiiiiiii");
            //NewPage.NameParts
            //считывать все эти свойства: NewPage.NameParts.DESIGNATION_INSTALLATIONNUMBER.ToString();
        }

        //// Доп функции ////
        private void LoadDataFromDatabase()
        {
            MDPartsDatabase partsDatabase = null;
            try
            {
                MDPartsManagement oPartsManagement = new MDPartsManagement();
                partsDatabase = oPartsManagement.OpenDatabase();

                if (partsDatabase == null)
                {
                    MessageBox.Show("Не удалось открыть базу данных изделий.", "Ошибка");
                    return;
                }

                MDPartsDatabaseItemPropertyList filter = new MDPartsDatabaseItemPropertyList();
                filter.ARTICLE_PRODUCTSUBGROUP = "1";
                filter.ARTICLE_PRODUCTGROUP = "5";

                MDPartsDatabaseItemPropertyList properties = new MDPartsDatabaseItemPropertyList();
                MDPart[] SubGroupParts = partsDatabase.GetParts(filter, properties);

                if (SubGroupParts != null && SubGroupParts.Length > 0)
                {
                    foreach (MDPart part in SubGroupParts)
                    {


                        // Производитель
                        string manufacturer = part.Properties.ARTICLE_MANUFACTURER;
                        manufacturer = manufacturer.ToLower();
                        if (manufacturer == "veda")
                        {
                            if (!string.IsNullOrEmpty(part.PartNr) && part.PartNr.ToUpper().Contains("VF")) 
                            {
                                // Создание объекта Device и заполнение его данными
                                Device device = new Device();
                                device.Number = part.PartNr;
                                device.Voltage = part.Properties.ARTICLE_VOLTAGE;
                                //device.Current = part.Properties.ARTICLE_ELECTRICALCURRENT;
                                device.Discribe = part.Properties.ARTICLE_NOTE.GetDisplayString().GetString(ISOCode.Language.L___);
                                device.OverloadCurrent = part.Properties[22337][1].ToMultiLangString().GetString(0).Replace(".", ",");
                                device.NormalCurrent = part.Properties[22337][2].ToMultiLangString().GetString(0).Replace(".", ",");

                                string powerString = part.Properties.ARTICLE_CHARACTERISTICS.GetDisplayString().GetString(ISOCode.Language.L___);
                                powerString = ExtractPowerValue(powerString);
                                device.Power = powerString;

                                devices.Add(device);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Не найдены ПЧ, соответствующие заданным критериям.", "Предупреждение");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message + "\r\n" + ex.StackTrace, "Ошибка");
            }
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
                        MessageBox.Show($"Ошибка при закрытии базы данных: {ex.Message}", "Ошибка");
                    }
                    finally
                    {
                        partsDatabase.Dispose();
                    }
                }
            }
        }

        private void DisplayDeviceInfo()
        {
            string deviceInfo = "";
            if (devices.Count > 0)
            {
                foreach (Device device in devices)
                {
                    deviceInfo += $"Номер: {device.Number}, Напряжение: {device.Voltage}, Мощность: {device.Power},\r\n" +
                        $" Ток выс перегрузки: {device.OverloadCurrent}, Ток норм перегрузки: {device.NormalCurrent}\r\n";
                }
            }
            else
            {
                deviceInfo = "Нет данных об устройствах.";
            }
            MessageBox.Show("Информация об устройствах:\r\n" + deviceInfo);
        }

        private Device FindDeviceNormal(string series, string voltage, float power, float current)
        {
            MessageBox.Show("Ищем по нормальному");
            // Серия и напряжение
            var filteredDevices = devices.Where(d => d.Number.Contains(series) && d.Voltage == voltage
            && !string.IsNullOrEmpty(d.Power) && !string.IsNullOrEmpty(d.NormalCurrent)).ToList();

            if (!filteredDevices.Any()) { return null; }

            //MessageBox.Show("rrrrrrrrrrrrrrr");
            // По мощности
            var samePowerDevices = filteredDevices.Where(d =>
            {
                if (float.TryParse(d.Power, out float devicePower))
                {
                    return devicePower == power;
                }
                return false; // Игнорируем, если не удалось распарсить мощность
            }).ToList();

            if (samePowerDevices.Any())
            {
                Device bestDeviceByCurrent = samePowerDevices
                    .OrderBy(d =>
                    {
                        if (float.TryParse(d.NormalCurrent, out float deviceCurrent))
                        {
                            float currentDifference = (deviceCurrent >= current)
                                                      ? (deviceCurrent - current)
                                                      : float.MaxValue;
                            return currentDifference;
                        }
                        return float.MaxValue;
                    })
                    .FirstOrDefault();
                if (float.Parse(bestDeviceByCurrent.NormalCurrent) > current)
                {
                    return bestDeviceByCurrent;
                }
            }

            //MessageBox.Show("hhhhhhhhhhhhhh");
            // По току
            var sameCurrentDevices = filteredDevices.Where(d =>
            {
                if (float.TryParse(d.NormalCurrent, out float deviceCurrent))
                {
                    return deviceCurrent == current;
                }
                return false;
            }).ToList();
            if (sameCurrentDevices.Any())
            {
                Device bestDeviceByPower = sameCurrentDevices
                    .OrderBy(d =>
                    {
                        if (float.TryParse(d.Power, out float devicePower))
                        {
                            float powerDifference = (devicePower >= power)
                                                     ? (devicePower - power)
                                                     : float.MaxValue;

                            return powerDifference;
                        }
                        return float.MaxValue;
                    })
                    .FirstOrDefault();
                if (float.Parse(bestDeviceByPower.Power) > power)
                {
                    return bestDeviceByPower;
                }
            }

            //MessageBox.Show("tttttttttttttttttt");
            Device bestDeviceOverall = filteredDevices
                .Where(d =>
                {
                    if (float.TryParse(d.Power, out float devicePower))
                    {
                        return devicePower >= power; // мощность >= power
                    }
                    return false; 
                })
                .OrderBy(d =>
                {
                    if (float.TryParse(d.NormalCurrent, out float deviceCurrent))
                    {
                        // Разница только по току
                        return (deviceCurrent >= current)
                                   ? (deviceCurrent - current)
                                   : float.MaxValue;
                    }
                    return float.MaxValue; 
                })
                .FirstOrDefault();
            if (float.Parse(bestDeviceOverall.NormalCurrent) > current)
            {
                return bestDeviceOverall;
            }
            return null;
        }

        private Device FindDeviceOverload(string series, string voltage, float power, float current)
        {
            MessageBox.Show("Ищем по высокому");
            // Серия и напряжение
            var filteredDevices = devices.Where(d => d.Number.Contains(series) && d.Voltage == voltage
            && !string.IsNullOrEmpty(d.Power) && !string.IsNullOrEmpty(d.OverloadCurrent)).ToList();

            if (!filteredDevices.Any()) { return null; }

            //MessageBox.Show("рррррррр");
            // По мощности
            var samePowerDevices = filteredDevices.Where(d =>
            {
                if (float.TryParse(d.Power, out float devicePower))
                {
                    return devicePower == power;
                }
                return false; // Игнорируем, если не удалось распарсить мощность
            }).ToList();

            if (samePowerDevices.Any())
            {
                Device bestDeviceByCurrent = samePowerDevices
                    .OrderBy(d =>
                    {
                        if (float.TryParse(d.OverloadCurrent, out float deviceCurrent))
                        {
                            float currentDifference = (deviceCurrent >= current)
                                                      ? (deviceCurrent - current)
                                                      : float.MaxValue;
                            return currentDifference;
                        }
                        return float.MaxValue;
                    })
                    .FirstOrDefault();
                if (float.Parse(bestDeviceByCurrent.OverloadCurrent) > current)
                {
                    return bestDeviceByCurrent;
                }
            }

            //MessageBox.Show("дддддддддд");
            // По току
            var sameCurrentDevices = filteredDevices.Where(d =>
            {
                if (float.TryParse(d.OverloadCurrent, out float deviceCurrent))
                {
                    return deviceCurrent == current;
                }
                return false;
            }).ToList();
            if (sameCurrentDevices.Any())
            {
                Device bestDeviceByPower = sameCurrentDevices
                    .OrderBy(d =>
                    {
                        if (float.TryParse(d.Power, out float devicePower))
                        {
                            float powerDifference = (devicePower >= power)
                                                     ? (devicePower - power)
                                                     : float.MaxValue;

                            return powerDifference;
                        }
                        return float.MaxValue;
                    })
                    .FirstOrDefault();
                if (float.Parse(bestDeviceByPower.Power) > power)
                {
                    return bestDeviceByPower;
                }
            }

            //MessageBox.Show("ааааааааааааааааа");
            Device bestDeviceOverall = filteredDevices
                .Where(d =>
                {
                    if (float.TryParse(d.Power, out float devicePower))
                    {
                        return devicePower >= power; // мощность >= power
                    }
                    return false;
                })
                .OrderBy(d =>
                {
                    if (float.TryParse(d.OverloadCurrent, out float deviceCurrent))
                    {
                        // Разница только по току
                        return (deviceCurrent >= current)
                                   ? (deviceCurrent - current)
                                   : float.MaxValue;
                    }
                    return float.MaxValue;
                })
                .FirstOrDefault();
            if (float.Parse(bestDeviceOverall.OverloadCurrent) > current)
            {
                return bestDeviceOverall;
            }
            return null;
        }

        private Protection FindProtection(string current)
        {
            string type = comboBox9.Text.Trim();
            bool find = false;
            Protection protect = new Protection();

            MDPartsDatabase partsDatabase = null;
            try
            {
                MDPartsManagement oPartsManagement = new MDPartsManagement();
                partsDatabase = oPartsManagement.OpenDatabase();

                if (partsDatabase == null)
                {
                    MessageBox.Show("Не удалось открыть базу данных изделий.", "Ошибка");
                    return null;
                }

                MDPartsDatabaseItemPropertyList filter = new MDPartsDatabaseItemPropertyList();
                filter.ARTICLE_PRODUCTGROUP = "6";
                if (type == "Защитный выключатель") { filter.ARTICLE_PRODUCTSUBGROUP = "193"; }
                else if (type == "Плавкий предохранитель") { filter.ARTICLE_PRODUCTSUBGROUP = "192"; }

                MDPartsDatabaseItemPropertyList properties = new MDPartsDatabaseItemPropertyList();
                MDPart[] SubGroupParts = partsDatabase.GetParts(filter, properties);

                if (SubGroupParts != null && SubGroupParts.Length > 0)
                {
                    foreach (MDPart part in SubGroupParts)
                    {
                        // Производитель
                        string manufacturer = part.Properties.ARTICLE_MANUFACTURER;
                        manufacturer = manufacturer.ToLower();
                        if (manufacturer == comboBox2.Text.Trim().ToLower())
                        {
                            if (!string.IsNullOrEmpty(part.PartNr))
                            {
                                string curr = part.Properties.ARTICLE_ELECTRICALCURRENT;
                                if (curr == current)
                                {
                                    protect.Number = part.PartNr;
                                    protect.Current = part.Properties.ARTICLE_ELECTRICALCURRENT;
                                    protect.Discribe = part.Properties.ARTICLE_NOTE.GetDisplayString().GetString(ISOCode.Language.L___);
                                    find = true;
                                }
                            }
                        }
                    }
                    if (!find) { MessageBox.Show("Не найдены защитные устройства, соответствующие заданным критериям.", "Предупреждение"); }
                }
                else
                {
                    MessageBox.Show("Не найдены защитные устройства, соответствующие заданным критериям.", "Предупреждение");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message + "\r\n" + ex.StackTrace, "Ошибка");
                return null;
            }
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
                        MessageBox.Show($"Ошибка при закрытии базы данных: {ex.Message}", "Ошибка");
                    }
                    finally
                    {
                        partsDatabase.Dispose();
                    }
                }
            }
            if (find) { return protect; }
            else { return null; }
        }

        private string FindProtectionCurrent(string series, string voltage, string power)
        {
            // Ищем ток предохранителя в таблице Excel
            if (voltage == "220")
            {
                return null;
            }

            string type = comboBox9.Text.Trim(); 
            string excelFilePath = textBox25.Text;

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                MessageBox.Show("Не указан путь к файлу Excel или файл не существует.", "Ошибка");
                return null;
            }

            try
            {
                // Открываем Excel файл
                using (FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet();
                        DataTable table = result.Tables[0];

                        // Определяем номер столбца
                        string powerColumn = "";
                        string currentColumn = "";

                        switch (series)
                        {
                            case "VF-101":
                                powerColumn = "A";
                                if (voltage == "380")
                                {
                                    currentColumn = (type == "Защитный выключатель") ? "C" : "D";
                                }
                                else if (voltage == "690")
                                {
                                    currentColumn = "F";
                                }
                                break;
                            case "VF-51":
                                if (voltage == "690")
                                {
                                    stream.Close();
                                    return null;
                                }
                                powerColumn = "K";
                                currentColumn = (type == "Защитный выключатель") ? "M" : "N";
                                break;
                            case "VF-11":
                                //powerColumn = "S";
                                //currentColumn = (type == "Защитный выключатель") ? "U" : "V";
                                //break;
                                stream.Close();
                                return null;
                            default:
                                MessageBox.Show($"Неизвестная серия оборудования: {series}", "Предупреждение");
                                stream.Close();
                                return null;
                        }

                        // Cтрока с мощностью
                        int powerColumnIndex = ColumnNameToNumber(powerColumn);
                        int currentColumnIndex = ColumnNameToNumber(currentColumn);

                        for (int row = 4; row < table.Rows.Count; row++) // С 5 строки
                        {
                            DataRow currentRow = table.Rows[row];
                            string powerValue = currentRow[powerColumnIndex]?.ToString().Trim();

                            if (string.Equals(powerValue, power, StringComparison.OrdinalIgnoreCase))
                            {
                                // Ток
                                string currentValue = currentRow[currentColumnIndex]?.ToString().Trim();

                                if (!string.IsNullOrEmpty(currentValue))
                                {
                                    // Убираем "gG-" или "aR-"
                                    if (type == "Плавкий предохранитель")
                                    {
                                        currentValue = currentValue.Replace("gG-", "").Replace("aR-", "").Trim();
                                    }
                                    stream.Close();
                                    return currentValue;
                                }
                                stream.Close();
                                return currentValue;
                            }
                        }

                        MessageBox.Show($"Мощность '{power}' не найдена в таблице для серии '{series}'.", "Предупреждение");
                        stream.Close();
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении файла Excel: {ex.Message}", "Ошибка");
                return null;
            }
        }

        private string FindFilterNumber(int num, string series, string voltage, string power)
        {
            // Ищем фильтр в таблице
            if (voltage == "220")
            {
                return null;
            }

            int type = 0;
            string excelFilePath = textBox26.Text;

            if (num == 1 && comboBox10.Text.Trim() != "Нет" && comboBox10.Text.Trim() != "AC дроссель") { type = 1; }
            else if (num == 2 && comboBox10.Text.Trim() != "Нет" && comboBox10.Text.Trim() != "ЭМС фильтр") { type = 2; }
            else if (num == 3 && comboBox11.Text.Trim() == "Дроссель") { type = 3; }
            else if (num == 4 && comboBox11.Text.Trim() == "Синус-фильтр") { type = 4; }

            if (type == 0) { return null; }

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                MessageBox.Show("Не указан путь к файлу Excel или файл не существует.", "Ошибка");
                return null;
            }

            try
            {
                // Открываем Excel файл
                using (FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet();
                        DataTable table = result.Tables[0];

                        // Определяем номер столбца
                        string powerColumn = "";
                        string filterColumn = "";

                        switch (series)
                        {
                            case "VF-101":
                                powerColumn = "A";
                                if (voltage == "380")
                                {
                                    if (type == 1) { filterColumn = "B"; }
                                    else if (type == 2) { filterColumn = "C"; }
                                    else if (type == 3) { filterColumn = "D"; }
                                    else if (type == 4) { filterColumn = "E"; }
                                }
                                else if (voltage == "690")
                                {
                                    if (type == 1) { filterColumn = "F"; }
                                    else if (type == 2) { filterColumn = "G"; }
                                    else if (type == 3) { filterColumn = "H"; }
                                    else if (type == 4) { filterColumn = "I"; }
                                }
                                break;
                            case "VF-51":
                                if (voltage == "690") 
                                {
                                    stream.Close();
                                    return null;
                                }
                                powerColumn = "N";
                                if (type == 1) { filterColumn = "O"; }
                                else if (type == 2) { filterColumn = "P"; }
                                else if (type == 3) { filterColumn = "Q"; }
                                else if (type == 4) { filterColumn = "R"; }
                                break;
                            case "VF-11":
                                //powerColumn = "S";
                                //currentColumn = (type == "Защитный выключатель") ? "U" : "V";
                                //break;
                                stream.Close();
                                return null;
                            default:
                                MessageBox.Show($"Неизвестная серия оборудования: {series}", "Предупреждение");
                                stream.Close();
                                return null;
                        }

                        // Cтрока с мощностью
                        int powerColumnIndex = ColumnNameToNumber(powerColumn);
                        int filterColumnIndex = ColumnNameToNumber(filterColumn);

                        for (int row = 3; row < table.Rows.Count; row++) // С 4 строки
                        {
                            DataRow filterRow = table.Rows[row];
                            string powerValue = filterRow[powerColumnIndex]?.ToString().Trim();

                            if (string.Equals(powerValue, power, StringComparison.OrdinalIgnoreCase))
                            {
                                // Номер
                                string filterName = filterRow[filterColumnIndex]?.ToString().Trim();
                                stream.Close();
                                return filterName;
                            }
                        }

                        MessageBox.Show($"Мощность '{power}' не найдена в таблице для серии '{series}'.", "Предупреждение");
                        stream.Close();
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении файла Excel: {ex.Message}", "Ошибка");
                return null;
            }
        }

        private void FindFilter(string input_filter, string input_throttle, string output_filter, string output_throttle)
        {
            MDPartsDatabase partsDatabase = null;
            try
            {
                MDPartsManagement oPartsManagement = new MDPartsManagement();
                partsDatabase = oPartsManagement.OpenDatabase();

                if (partsDatabase == null)
                {
                    MessageBox.Show("Не удалось открыть базу данных изделий.", "Ошибка");
                    return;
                }

                if (!string.IsNullOrEmpty(input_filter) || !string.IsNullOrEmpty(output_filter)) 
                {
                    MDPartsDatabaseItemPropertyList filter = new MDPartsDatabaseItemPropertyList();
                    filter.ARTICLE_PRODUCTSUBGROUP = "1";
                    filter.ARTICLE_PRODUCTGROUP = "24";

                    MDPartsDatabaseItemPropertyList properties = new MDPartsDatabaseItemPropertyList();
                    MDPart[] SubGroupParts = partsDatabase.GetParts(filter, properties);

                    if (SubGroupParts != null && SubGroupParts.Length > 0)
                    {
                        foreach (MDPart part in SubGroupParts)
                        {
                            // Производитель
                            string manufacturer = part.Properties.ARTICLE_MANUFACTURER;
                            manufacturer = manufacturer.ToLower();
                            if (manufacturer == "veda" && (part.PartNr == input_filter || part.PartNr == output_filter))
                            {
                                // Создание объекта и заполнение его данными
                                Filter filt = new Filter();
                                filt.Number = part.PartNr;
                                //filt.Current = part.Properties.ARTICLE_ELECTRICALCURRENT;
                                filt.Discribe = part.Properties.ARTICLE_NOTE.GetDisplayString().GetString(ISOCode.Language.L___);
                                filters.Add(filt);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Не найдены фильтры, соответствующие заданным критериям.", "Предупреждение");
                    }
                }

                if (!string.IsNullOrEmpty(input_throttle) || !string.IsNullOrEmpty(output_throttle))
                {
                    MDPartsDatabaseItemPropertyList filter = new MDPartsDatabaseItemPropertyList();
                    filter.ARTICLE_PRODUCTSUBGROUP = "1";
                    filter.ARTICLE_PRODUCTGROUP = "21";

                    MDPartsDatabaseItemPropertyList properties = new MDPartsDatabaseItemPropertyList();
                    MDPart[] SubGroupParts = partsDatabase.GetParts(filter, properties);

                    if (SubGroupParts != null && SubGroupParts.Length > 0)
                    {
                        foreach (MDPart part in SubGroupParts)
                        {
                            // Производитель
                            string manufacturer = part.Properties.ARTICLE_MANUFACTURER;
                            manufacturer = manufacturer.ToLower();
                            if (manufacturer == "veda" && (part.PartNr == input_throttle || part.PartNr == output_throttle))
                            {
                                // Создание объекта и заполнение его данными
                                Filter filt = new Filter();
                                filt.Number = part.PartNr;
                                //filt.Current = part.Properties.ARTICLE_ELECTRICALCURRENT;
                                filt.Discribe = part.Properties.ARTICLE_NOTE.GetDisplayString().GetString(ISOCode.Language.L___);
                                filters.Add(filt);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Не найдены дроссель, соответствующие заданным критериям.", "Предупреждение");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message + "\r\n" + ex.StackTrace, "Ошибка");
            }
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
                        MessageBox.Show($"Ошибка при закрытии базы данных: {ex.Message}", "Ошибка");
                    }
                    finally
                    {
                        partsDatabase.Dispose();
                    }
                }
            }
        }

        private int ColumnNameToNumber(string columnName)
        {
            // Названия столбцов
            int number = 0;
            foreach (char c in columnName.ToUpper())
            {
                number *= 26;
                number += (c - 'A' + 1);
            }
            return number - 1;
        }

        private string ExtractPowerValue(string powerString)
        {
            // Достаем из строки мощность
            string result = "";
            // Используем регулярное выражение
            System.Text.RegularExpressions.Match match = Regex.Match(powerString, @"([\d,\.]+)\s*(кВт|kW|Квт|квт)", RegexOptions.IgnoreCase);

            if (match.Success)
            {
                result = match.Groups[1].Value.Trim();
                result = result.Replace(",", ".");

                if (!float.TryParse(result, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    result = "";
                }
            }
            return result;
        }

        private void ParameterChanged(object sender, EventArgs e)
        {
            // Обновление ComboBox6
            if (!checkBox1.Checked)
            {
                UpdateDeviceList();
            }
        }

        private void UpdateDeviceList()
        {
            comboBox6.Items.Clear();

            string series = comboBox1.SelectedItem?.ToString();
            string voltage = comboBox5.SelectedItem?.ToString();
            string power = textBox1.Text.Trim();
            string current = textBox2.Text.Trim();

            // Фильтр спискса устройств
            IEnumerable<Device> filteredDevices = devices;

            if (!string.IsNullOrEmpty(series))
            {
                filteredDevices = filteredDevices.Where(d => d.Number.Contains(series));
            }
            if (!string.IsNullOrEmpty(voltage))
            {
                filteredDevices = filteredDevices.Where(d => d.Voltage == voltage);
            }
            if (!string.IsNullOrEmpty(power))
            {
                filteredDevices = filteredDevices.Where(d => d.Power == power);
            }
            if (!string.IsNullOrEmpty(current) && flagNormal)
            {
                filteredDevices = filteredDevices.Where(d => d.NormalCurrent == current);
            }
            if (!string.IsNullOrEmpty(current) && !flagNormal)
            {
                filteredDevices = filteredDevices.Where(d => d.OverloadCurrent == current);
            }

            // Добавляем
            foreach (Device device in filteredDevices)
            {
                comboBox6.Items.Add(device.Number);
            }
            if (comboBox6.Items.Count == 0)
            {
                comboBox6.Items.Add("Устройства не найдены");
            }
        }

        public class Device
        {
            // Класс для представления ПЧ
            public string Number { get; set; }
            public string Voltage { get; set; }
            //public string Current { get; set; }
            public string OverloadCurrent { get; set; }
            public string NormalCurrent { get; set; }
            public string Power { get; set; }
            public string Discribe { get; set; }
        }

        public class Protection
        {
            // Класс для представления Защитного устройства
            public string Number { get; set; }
            public string Current { get; set; }
            public string Discribe { get; set; }
        }

        public class Filter
        {
            // Класс для представления Защитного устройства
            public string Number { get; set; }
            //public string Current { get; set; }
            public string Discribe { get; set; }
        }

        // Excel
        //private void CreateBook()
        //{
        //    /////////// Книга 1 ///////////
        //    int numRows = 13;
        //    int numCells = 2;

        //    IWorkbook workbook = new XSSFWorkbook();    // Создаем книгу Excel (.xlsx)
        //    ISheet sheet = workbook.CreateSheet("Лист1");   // Создаем лист, строки и ячейки

        //    // Стили
        //    ICellStyle centerAlignedStyle = CreateCenterAlignedBoldStyle(workbook);
        //    ICellStyle cellOutline = CreateCellOutlineStyle(workbook);
        //    ICellStyle styledCell = CreateStyledCell(workbook);

        //    for (int rowNum = 0; rowNum < numRows; rowNum++)
        //    {
        //        IRow row = sheet.CreateRow(rowNum);
        //        for (int cellNum = 0; cellNum < numCells; cellNum++)
        //        {
        //            NPOI.SS.UserModel.ICell cell = row.CreateCell(cellNum);
        //            if (rowNum > 5)
        //            {
        //                cell.CellStyle = cellOutline;
        //            }
        //        }
        //    }

        //    string dateString = "Дата: " + DateTime.Now.ToString("dd.MM.yyyy");
        //    sheet.GetRow(3).GetCell(0).SetCellValue(dateString);

        //    sheet.GetRow(5).GetCell(0).SetCellValue("Наименование");
        //    sheet.GetRow(5).GetCell(0).CellStyle = styledCell;
        //    sheet.GetRow(5).GetCell(1).SetCellValue("Значение");
        //    sheet.GetRow(5).GetCell(1).CellStyle = styledCell;

        //    sheet.GetRow(6).GetCell(0).SetCellValue(label1.Text); // Устанавливаем значение ячейки
        //    sheet.GetRow(6).GetCell(1).SetCellValue(Convert.ToString(comboBox1.SelectedItem));

        //    sheet.GetRow(7).GetCell(0).SetCellValue(label2.Text);
        //    sheet.GetRow(7).GetCell(1).SetCellValue(Convert.ToString(comboBox2.SelectedItem));

        //    sheet.GetRow(8).GetCell(0).SetCellValue(label3.Text);
        //    sheet.GetRow(8).GetCell(1).SetCellValue(Convert.ToString(comboBox3.SelectedItem));

        //    sheet.GetRow(9).GetCell(0).SetCellValue(label4.Text);
        //    sheet.GetRow(9).GetCell(1).SetCellValue(Convert.ToString(comboBox4.SelectedItem));

        //    sheet.GetRow(10).GetCell(0).SetCellValue(label7.Text);
        //    sheet.GetRow(10).GetCell(1).SetCellValue(Convert.ToString(comboBox5.SelectedItem));

        //    sheet.GetRow(11).GetCell(0).SetCellValue(label8.Text);
        //    sheet.GetRow(11).GetCell(1).SetCellValue(foundDevice.Power + " кВт");

        //    sheet.GetRow(12).GetCell(0).SetCellValue(label9.Text);
        //    sheet.GetRow(12).GetCell(1).SetCellValue(foundDevice.Current + " А");

        //    sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1)); // объединение
        //    sheet.GetRow(1).GetCell(0).SetCellValue("Характеристики");
        //    sheet.GetRow(1).GetCell(0).CellStyle = centerAlignedStyle;

        //    sheet.SetColumnWidth(0, 29 * 256);
        //    sheet.SetColumnWidth(1, 20 * 256);

        //    // Получаем путь к папке "Документы" на рабочем столе
        //    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //    string folderName = "Документы";
        //    string fullFolderPath = Path.Combine(desktopPath, folderName);

        //    // Проверяем, существует ли папка. Если нет - создаем
        //    if (!Directory.Exists(fullFolderPath))
        //    {
        //        Directory.CreateDirectory(fullFolderPath);
        //    }

        //    // Указываем полный путь к файлу
        //    string fileName = "myworkbook.xlsx";
        //    string filePath = Path.Combine(fullFolderPath, fileName);

        //    // Создаем файловый поток для записи данных в файл
        //    using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        //    {
        //        // Записываем книгу в файловый поток
        //        workbook.Write(fileStream);
        //    }
        //}

        //private void CreateBook_2(string price, string article, string discribe)
        //{
        //    /////////// Книга 2 ///////////
        //    int numRows = 16;
        //    int numCells = 7;
        //    IWorkbook workbook_2 = new XSSFWorkbook();    // Создаем книгу Excel (.xlsx)
        //    ISheet sheet_2 = workbook_2.CreateSheet("Лист1");   // Создаем лист, строки и ячейки

        //    // Стили
        //    ICellStyle centerAlignedStyle = CreateCenterAlignedBoldStyle(workbook_2);
        //    ICellStyle cellOutline = CreateCellOutlineStyle(workbook_2);
        //    ICellStyle styledCell = CreateStyledCell(workbook_2);


        //    for (int rowNum = 0; rowNum < numRows; rowNum++)
        //    {
        //        IRow row = sheet_2.CreateRow(rowNum);
        //        if (rowNum == 1)
        //        {
        //            row.HeightInPoints = 21;
        //        }
        //        for (int cellNum = 0; cellNum < numCells; cellNum++)
        //        {
        //            NPOI.SS.UserModel.ICell cell = row.CreateCell(cellNum, CellType.String);

        //            if (rowNum == 12) { cell.CellStyle = styledCell; }
        //            else if (rowNum == 13) { cell.CellStyle = cellOutline; }
        //        }
        //    }

        //    try
        //    {
        //        sheet_2.AddMergedRegion(new CellRangeAddress(1, 1, 1, 5));   // объединение
        //                                                                     //sheet_2.AddMergedRegion(new CellRangeAddress(2, 3, 2, 2));
        //                                                                     //sheet_2.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(4, 5, 2, 2));


        //        sheet_2.GetRow(1).GetCell(1).SetCellValue("Коммерческое предложение");
        //        sheet_2.GetRow(1).GetCell(1).CellStyle = centerAlignedStyle;

        //        sheet_2.GetRow(3).GetCell(1).SetCellValue("От:");
        //        sheet_2.GetRow(4).GetCell(1).SetCellValue("Кому:");
        //        sheet_2.GetRow(3).GetCell(1).CellStyle = cellOutline;
        //        sheet_2.GetRow(4).GetCell(1).CellStyle = cellOutline;
        //        sheet_2.GetRow(3).GetCell(2).CellStyle = cellOutline;
        //        sheet_2.GetRow(4).GetCell(2).CellStyle = cellOutline;

        //        string dateString = "Дата: " + DateTime.Now.ToString("dd.MM.yyyy");
        //        sheet_2.GetRow(8).GetCell(1).SetCellValue(dateString);
        //        //sheet_2.AddMergedRegion(new CellRangeAddress(9, 9, 0, 2));
        //        //sheet_2.GetRow(9).GetCell(1).SetCellValue("Счетом");

        //        sheet_2.GetRow(12).GetCell(0).SetCellValue("№ п/п");
        //        sheet_2.GetRow(13).GetCell(0).SetCellValue("1");
        //        sheet_2.GetRow(12).GetCell(1).SetCellValue("Артикул");
        //        sheet_2.GetRow(13).GetCell(1).SetCellValue(article);
        //        sheet_2.GetRow(12).GetCell(2).SetCellValue("Наименование");
        //        sheet_2.GetRow(13).GetCell(2).SetCellValue(discribe);
        //        sheet_2.GetRow(12).GetCell(3).SetCellValue("Кол-во шт/упак.");
        //        sheet_2.GetRow(13).GetCell(3).SetCellValue("1");
        //        sheet_2.GetRow(12).GetCell(4).SetCellValue("Цена в руб.");
        //        sheet_2.GetRow(12).GetCell(5).SetCellValue("Цена со скидкой в руб.");
        //        sheet_2.GetRow(12).GetCell(6).SetCellValue("Наличие");

        //        if (price != "Нет на складе")
        //        {
        //            string price_num = price.Trim();
        //            //string price_num = price.Replace(",", ".").Trim();
        //            price_num = Regex.Replace(price_num, "[^0-9.,]", "");
        //            sheet_2.GetRow(13).GetCell(4).SetCellValue(price_num);
        //            float discountPrice = float.Parse(price_num);
        //            sheet_2.GetRow(13).GetCell(5).SetCellValue(price_num);
        //            if (!string.IsNullOrEmpty(textBox3.Text))
        //            {
        //                float discount = (float)Convert.ToDouble(textBox3.Text);
        //                discountPrice = discountPrice * (100 - discount) / 100;
        //                string formattedFloat = string.Format("{0:F2}", discountPrice);
        //                sheet_2.GetRow(13).GetCell(5).SetCellValue(formattedFloat);
        //            }
        //        }
        //        else
        //        {
        //            sheet_2.GetRow(13).GetCell(4).SetCellValue("--");
        //            sheet_2.GetRow(13).GetCell(5).SetCellValue("--");
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка");
        //        MessageBox.Show(ex.StackTrace); // Выводим трассировку стека для более детальной информации
        //    }

        //    sheet_2.SetColumnWidth(0, 4 * 256);
        //    sheet_2.SetColumnWidth(1, 30 * 256);
        //    sheet_2.SetColumnWidth(2, 50 * 256);
        //    sheet_2.SetColumnWidth(3, 11 * 256);
        //    sheet_2.SetColumnWidth(4, 11 * 256);
        //    sheet_2.SetColumnWidth(5, 11 * 256);
        //    sheet_2.SetColumnWidth(6, 11 * 256);

        //    // Получаем путь к папке "Документы" на рабочем столе
        //    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //    string folderName = "Документы";
        //    string fullFolderPath = Path.Combine(desktopPath, folderName);

        //    // Проверяем, существует ли папка. Если нет - создаем
        //    if (!Directory.Exists(fullFolderPath))
        //    {
        //        Directory.CreateDirectory(fullFolderPath);
        //    }

        //    // Указываем полный путь к файлу
        //    string fileName = "myworkbook_2.xlsx";
        //    string filePath = Path.Combine(fullFolderPath, fileName);

        //    // Создаем файловый поток для записи данных в файл
        //    using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        //    {
        //        // Записываем книгу в файловый поток
        //        workbook_2.Write(fileStream);
        //    }
        //}


        //private ICellStyle CreateCenterAlignedBoldStyle(IWorkbook workbook)
        //{
        //    // Стиль для жирного шрифта и размера + выравнивания по центру
        //    IFont boldFont = workbook.CreateFont();
        //    boldFont.IsBold = true;
        //    boldFont.FontHeightInPoints = 16; // Устанавливаем размер шрифта (в пунктах)
        //    ICellStyle boldStyle = workbook.CreateCellStyle();
        //    boldStyle.SetFont(boldFont);
        //    ICellStyle centerAlignedStyle = workbook.CreateCellStyle();
        //    centerAlignedStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        //    centerAlignedStyle.VerticalAlignment = VerticalAlignment.Center;
        //    centerAlignedStyle.SetFont(boldFont);
        //    centerAlignedStyle.WrapText = true;
        //    return centerAlignedStyle;
        //}

        //private ICellStyle CreateCellOutlineStyle(IWorkbook workbook)
        //{
        //    // Стиль для обводки + по центру по вертикали
        //    ICellStyle cellOutline = workbook.CreateCellStyle();
        //    cellOutline.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellOutline.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellOutline.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellOutline.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        //    cellOutline.TopBorderColor = IndexedColors.Black.Index;
        //    cellOutline.BottomBorderColor = IndexedColors.Black.Index;
        //    cellOutline.LeftBorderColor = IndexedColors.Black.Index;
        //    cellOutline.RightBorderColor = IndexedColors.Black.Index;
        //    cellOutline.VerticalAlignment = VerticalAlignment.Center;
        //    cellOutline.WrapText = true;
        //    return cellOutline;
        //}

        //private ICellStyle CreateStyledCell(IWorkbook workbook)
        //{
        //    // Стиль для выравнивания по центру + серая заливка + обводка
        //    ICellStyle styledCell = workbook.CreateCellStyle();
        //    styledCell.FillPattern = FillPattern.SolidForeground;
        //    styledCell.FillForegroundColor = IndexedColors.Grey25Percent.Index;
        //    styledCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        //    styledCell.VerticalAlignment = VerticalAlignment.Center;
        //    styledCell.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        //    styledCell.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        //    styledCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        //    styledCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        //    styledCell.TopBorderColor = IndexedColors.Black.Index;
        //    styledCell.BottomBorderColor = IndexedColors.Black.Index;
        //    styledCell.LeftBorderColor = IndexedColors.Black.Index;
        //    styledCell.RightBorderColor = IndexedColors.Black.Index;
        //    styledCell.WrapText = true;
        //    return styledCell;
        //}

        private async Task<string> GetPriceFromDrivesRu(string searchQuery)
        {
            string url = $"https://drives.ru/search/?query={searchQuery}";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string htmlContent = await response.Content.ReadAsStringAsync();

                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                    htmlDocument.LoadHtml(htmlContent);

                    // Элементы, содержащие информацию о товаре
                    HtmlNodeCollection productNodes = htmlDocument.DocumentNode.SelectNodes("//div[@class='s-products__item']");

                    if (productNodes != null)
                    {
                        foreach (HtmlNode productNode in productNodes)
                        {
                            // Название товара
                            HtmlNode nameNode = productNode.SelectSingleNode(".//div[@class='s-products__name']/a"); 
                            string productName = nameNode?.InnerText.Trim();

                            // Проверка
                            if (!string.IsNullOrEmpty(productName) && productName.Contains(searchQuery))
                            {
                                // Цена
                                HtmlNode priceNode = productNode.SelectSingleNode(".//div[@class='products__pr-price-new']//span[@class='price']");
                                if (priceNode != null)
                                {
                                    string priceText = priceNode.InnerText.Trim();
                                    return priceText;
                                }
                                else
                                {
                                    //MessageBox.Show($"Цена не найдена для товара: {productName}", "Предупреждение");
                                }
                            }
                        }
                        MessageBox.Show($"Товар с названием, содержащим '{searchQuery}', не найден.", "Ошибка");
                        return null;
                    }
                    else
                    {
                        MessageBox.Show("Не найдено элементов товаров на странице.", "Ошибка");
                        return null;
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                MessageBox.Show($"Ошибка HTTP: {ex.Message}", "Ошибка");
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при парсинге: {ex.Message}", "Ошибка");
                return null;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string load = comboBox3.SelectedItem.ToString();
            if (load == "Общепром. нагрузка") { flagNormal = false; }
            else {  flagNormal = true; }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem.ToString() == "VF-51")
            {
                comboBox9.SelectedIndex = 0;
                comboBox9.Enabled = false;
            }
            else { comboBox9.Enabled = true; }
        }
    }
}
