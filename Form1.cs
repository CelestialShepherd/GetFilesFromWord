using System;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
//Сторонние пакеты
using IronXL;
using Ionic.Zip;
using System.Collections.Generic;
using System.Xml;
using System.Text.RegularExpressions;

namespace GetFilesFromWord
{
    public partial class Form1 : Form
    {
        //Источник токена отмены
        CancellationTokenSource _tokenSource;

        public Form1()
        {
            InitializeComponent();
        }

        class Fact
        {
            string engWord { get; set; }
            string rusFact { get; set; }
            string engFact { get; set; }
            bool hasImage { get; set; }

            public Fact(string eW, string rF, string eF, bool hI)
            {
                engWord = eW;
                rusFact = rF;
                engFact = eF;
                hasImage = hI;
            }
        }

        //Выбрать путь к файлам фактов
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog1.ShowDialog();
                if (!folderBrowserDialog1.SelectedPath.Equals(""))
                {
                    textBox1.Text = folderBrowserDialog1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }
        
        //Выбрать путь к файлам вывода
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog2.ShowDialog();
                if (!folderBrowserDialog2.SelectedPath.Equals(""))
                {
                    textBox2.Text = folderBrowserDialog2.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

        //Запустить процесс
        private void button3_Click(object sender, EventArgs e)
        {
            ChangeElementsAvailabilityStatus();
            ClearConsole();
            _tokenSource = new CancellationTokenSource();
            CancellationToken cancelToken = _tokenSource.Token;
            GenerateLog("Процесс запущен!");

            string pathSource = textBox1.Text + "\\";
            string pathResult = textBox2.Text + "\\";

            List<string> xmlPaths = new List<string>();

            try
            {
                cancelToken.ThrowIfCancellationRequested();
                //Команды
                //Создаем папок с файлами для вывода
                CreateFolders(pathResult);
                //Получаем список файлов с фактами
                string[] files = Directory.GetFiles(textBox1.Text);
                //Создаем и переносим вспомогательные файлы и файлы изображений по путям
                GenerateLog("Генерация архивов и вспомогательных файлов");
                xmlPaths = LocateWordSourceFiles(pathSource, pathResult, files);
                //Запуск получения интересных фактов
                GetFacts(pathResult, xmlPaths);
                //Команды
                cancelToken.ThrowIfCancellationRequested();
            }
            catch (Exception ex)
            {
                GenerateLog(ex.Message);
            }
        }

        //Остановить процесс
        private void button4_Click(object sender, EventArgs e)
        {
            _tokenSource.Cancel();
            GenerateLog("Процесс прекращен!");
            ChangeElementsAvailabilityStatus();
        }

        //Очистить консоль
        private void button5_Click(object sender, EventArgs e)
        {
            ClearConsole();
        }

        //Генерация логов в консоль
        private void GenerateLog(string message)
        {
            textBox3.Text = "==========================================\r\n\r\n" + message + "\r\n\r\n==========================================\r\n\r\n" + textBox3.Text;
            Application.DoEvents();
        }

        //Изменение статус доступности элементов
        private void ChangeElementsAvailabilityStatus()
        {
            //Путь к файлам
            textBox1.Enabled = !textBox1.Enabled;
            button1.Enabled = !button1.Enabled;
            textBox2.Enabled = !textBox2.Enabled;
            button2.Enabled = !button2.Enabled;
            //Кнопки
            button3.Enabled = !button3.Enabled;
            button4.Enabled = !button4.Enabled;
            button5.Enabled = !button5.Enabled;
        }

        //Очистить консоль
        private void ClearConsole()
        {
            textBox3.Clear();
            textBox3.ClearUndo();
        }

        //Создаем папки для записи результатов
        static void CreateFolders(string path)
        {
            Directory.CreateDirectory(path + "L_AFP");
            Directory.CreateDirectory(path + "L1_E");
            Directory.CreateDirectory(path + "L2_E");
        }

        //Создаем и переносим вспомогательные файлы и файлы изображений по путям
        private List<string> LocateWordSourceFiles(string pathSource, string pathResult, string[] files)
        {
            int counter = 1;

            string fileName = "";
            string picPath = "";
            string zipPath = "";
            string xmlPath = "";

            List<string> xmlPaths = new List<string>();

            foreach (string filePath in files)
            {
                GenerateLog($"{counter}. {filePath}");
                //Получение наименования файла с фактами
                fileName = filePath.Replace(pathSource, "");
                //Получение пути и создание папки в директории с картинками
                picPath = pathResult + "L_AFP\\" + fileName.Replace(".docx", "");
                Directory.CreateDirectory(picPath);
                //Получение пути, создание папки в диркетории с архивами и генерация архива из Word-документа
                Directory.CreateDirectory(pathResult + "ZipFiles\\");
                zipPath = pathResult + "ZipFiles\\" + fileName.Replace(".docx", ".zip");
                if (System.IO.File.Exists(zipPath))
                    System.IO.File.Delete(zipPath);
                System.IO.File.Copy(filePath, zipPath);
                //Получение пути и создании папки в директории с xml-файлами
                xmlPath = pathResult + "XmlFiles\\" + fileName.Replace(".docx", "");
                Directory.CreateDirectory(xmlPath);
                //Перемещение изображений и xml-файлов в обозначенные директории
                using (var zip = ZipFile.Read(zipPath))
                {
                    int totalEntries = zip.Entries.Count;
                    foreach (ZipEntry e in zip.Entries)
                    {
                        if (e.FileName.Contains("word/media/"))
                        {
                            if (System.IO.File.Exists(picPath + "/" + e.FileName))
                                System.IO.File.Delete(picPath + "/" + e.FileName);
                            e.Extract(picPath);
                        }
                        if (e.FileName.Contains("word/document.xml"))
                        {
                            if (System.IO.File.Exists(xmlPath + "/word/document.xml"))
                                System.IO.File.Delete(xmlPath + "/word/document.xml");
                            e.Extract(xmlPath);
                            xmlPaths.Add(xmlPath);
                        }
                    }
                }
                counter++;
            }

            return xmlPaths;
        }

        //Получение фактов
        static void GetFacts(string pathResult, List<string> xmlPaths)
        {
            //if (files.Length != xmlPaths.Count)
            //    throw new Exception("Количество файлов с фактами и вспопомгательных xml-файлов не совпадают");

            /*Вспомогательные переменные*/
            //Строковые
            string xmlText = "";
            string tcXml = "";
            string strTemp = "";
            string strResult = "";
            //Численные
            int indexStartTemp = Int32.MaxValue;
            int indexState = 0;
            int factsCounter = 1;
            //Булевы
            bool hasImage = false;
            bool isAccentWord = false;
            //Массивы
            string[] facts = new string[3] { "", "", "" };
            //Листы
            List<string> stringsList = new List<string>();
            List<Fact> factsList = new List<Fact>();
            //Регулярные выражения
            Regex regexEng = new Regex("([A-Za-z])+([0-9])*$");
            //Excel
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.CreateWorkSheet("Facts");

            for (int i = 0; i < xmlPaths.Count; i++)
            {
                using (StreamReader reader = new StreamReader(xmlPaths[0]))
                {
                    xmlText = reader.ReadToEnd();
                }
                do
                {
                    //Ячейка таблицы
                    if (xmlText.Contains("<w:tc>"))
                    {
                        //Обрезаем между тегами ячейки <w:tc>
                        tcXml = xmlText.Substring(xmlText.IndexOf("<w:tc>"));
                        tcXml = xmlText.Substring(0, xmlText.IndexOf("</w:tc>") + 7);
                        if (tcXml.Contains("<pic:nvPicPr>"))
                            hasImage = true;

                        //Обрезаем между тегами ячейки <w:r>
                        //Ячейка таблицы
                        foreach (string text in tcXml.Split(new string[] { "</w:r>" }, StringSplitOptions.None))
                        {
                            //Строки в ячейке таблицы
                            foreach (string text2 in text.Split(new string[] { "</w:r>" }, StringSplitOptions.None))
                            {
                                if (text2.Equals("") || text2.Equals(null))
                                    continue;
                                else
                                    strResult += " ";

                                foreach (string text3 in text2.Split(new string[] { "</w:t>" }, StringSplitOptions.None))
                                {
                                    //Получение стартового индекса получения текста под тегом <w:t> в ячейке
                                    indexStartTemp = CalculateIndexStart(text3);
                                    if (indexStartTemp == -1)
                                    {
                                        strTemp = "";
                                        break;
                                    }
                                    else
                                    {
                                        strTemp = text3.Substring(indexStartTemp);
                                        strTemp = strTemp.Substring(strTemp.IndexOf(">") + 1);
                                    }
                                    //Условие сложения текста в ячейке
                                    if (strTemp != "")
                                    {
                                        if (strTemp.Contains((char)769))
                                        {
                                            strResult = strResult.Trim();
                                            isAccentWord = true;
                                        }
                                        else if (isAccentWord)
                                        {
                                            strResult = strResult.Trim();
                                            isAccentWord = false;
                                        }
                                    }
                                }
                            }
                        }
                        strResult = strResult.Trim();
                        switch (indexState)
                        {
                            case 0:
                                if (regexEng.IsMatch(strResult) && strResult != "")
                                {
                                    facts[0] = strResult;
                                    indexState++;
                                }
                                break;
                            case 1:
                                if (strResult != "")
                                {
                                    facts[1] = strResult.Replace("  ", " ");
                                    indexState++;
                                }
                                break;
                            case 2:
                                if (strResult != "")
                                {
                                    facts[2] = strResult.Replace("  ", " "); ;
                                    indexState++;
                                }
                                break;
                            default:
                                break;
                        }
                        stringsList.Add(strResult);
                        strResult = "";
                        if (indexState == 3)
                        {
                            factsList.Add(new Fact(facts[0], facts[1], facts[2], hasImage));
                            System.IO.File.WriteAllText(pathResult + "L1_E\\" + facts[0] + ".html", facts[1]);
                            System.IO.File.WriteAllText(pathResult + "L2_E\\" + facts[0] + ".html", facts[2]);
                            /*Запись в Excel*/
                            //Запись данных
                            factsCounter++;
                            workSheet[$"A{factsCounter}"].Value = facts[0];
                            workSheet[$"B{factsCounter}"].Value = hasImage ? 1 : 0;
                            //TODO: Доделать проверку по этому параметру
                            workSheet[$"C{factsCounter}"].Value = 0;
                            workSheet[$"D{factsCounter}"].Value = facts[1];
                            workSheet[$"E{factsCounter}"].Value = facts[2];

                            //Обнуление параметров
                            indexState = 0;
                        }
                        //Завершаем обрезание тега <w:r> ячейки 
                        xmlText = xmlText.Substring(xmlText.IndexOf("</w:tc>") + 7); 
                    }
                    //Завершаем обрезание тега <w:tc> ячейки
                    else
                        break;

                } while (true);
                
                //Сохранение Excel
                workBook.SaveAs($"{pathResult}Result_{i}_" +
                $"{DateTime.Now.Day}_" +
                $"{DateTime.Now.Month}_" +
                $"{DateTime.Now.Year}__" +
                $"{DateTime.Now.Hour}_" +
                $"{DateTime.Now.Minute}_" +
                $"{DateTime.Now.Second}.xlsx");
            }
        }

        static int CalculateIndexStart(string text)
        {
            int index1 = text.IndexOf("<w:t>");
            int index2 = text.IndexOf("<w:t ");

            if (index1 < index2)
            {
                if (index1 == -1)
                    return index2;
                else
                    return index1;
            }
            else if (index2 < index1)
            {
                if (index2 == -1)
                    return index1;
                else
                    return index2;
            }
            else
            {
                return -1;
            }
        }
    }
}
