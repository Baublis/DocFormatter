using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Diagnostics;
using Image = iTextSharp.text.Image;

namespace DocFormatter
{
    
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Переменные
        /// </summary>
        BaseFont BfArial;
        List<string> sourceList = new List<string>();
        string path_PDF;
        string[] paragraphs;
        static int _sectionNumber = 0;
        static int _pictureNumber = 0;
        static int _tableNumber = 0;
        int _sourceNumber = 0;


        /// <summary>
        /// Список контрольных фраз
        /// </summary>
        string[] arrayControlWords =
        {
            "[*номер раздела*]",                //0
            "[*номер рисунка*]",                //1
            "[*номер таблицы*]",                //2
            "[*ссылка на следующий рисунок*]",  //3
            "[*ссылка на предыдущий рисунок*]", //4
            "[*ссылка на таблицу*]",            //5
            "[*таблица ",                       //6
            "[*cписок литературы*]",            //7
            "[*код",                            //8
            "[*рисунок "                        //9
        };

        /// <summary>
        /// Главное окно:
        /// </summary>
        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            BfArial = BaseFont.CreateFont("arialuni.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED, BaseFont.CACHED, Properties.Resources.ARIAL, null);
            InitializeComponent();
        }

		private void Load_TXT(object sender, RoutedEventArgs e)
		{
            DefaultDialogService defaultDialogService = new DefaultDialogService();
            if (defaultDialogService.OpenFileDialogTXT())
            {
                
                paragraphs = File.ReadAllLines(defaultDialogService.FilePath);          
            }
            Save_PDF();
        }

        private void Save_PDF()
        {
            var fileContent = string.Empty;
            DefaultDialogService defaultDialogService = new DefaultDialogService();
            if (defaultDialogService.SaveFileDialogPDF())
            {
                path_PDF = defaultDialogService.FilePath;
                CreatePDF(path_PDF);
                ProcessStart(path_PDF);
            }
        }

        private void CreatePDF(string path)
        {

            FileStream DestinationStream = File.Create(path);
            Document document = new Document(PageSize.A4, 50f, 50f, 50f, 50f);
            PdfWriter.GetInstance(document, DestinationStream);
            document.Open();

            foreach (string paragraph in paragraphs)
            {
                string textParagraph = SearchSource(paragraph);

                var iparagraph = new iTextSharp.text.Paragraph(textParagraph.Replace("  "," "), new Font(BfArial, 12f, Font.NORMAL));
                /// <summary>
                /// Базовый шрифт для парагрофа
                /// </summary>             
                iparagraph.Alignment = Element.ALIGN_JUSTIFIED;

                /// <summary>
                /// Интервал после абзаца
                /// </summary>
                iparagraph.SpacingAfter = 0;

                /// <summary>
                /// Интервал перед абзацом
                /// </summary>
                iparagraph.SpacingBefore = 0;

                /// <summary>
                /// Отступ первой строки
                /// </summary>
                iparagraph.FirstLineIndent = 20f;

                //iparagraph.ExtraParagraphSpace = 10;

                for (int i = 0; i < arrayControlWords.Length; i++)
                {
                    if (paragraph.Contains(arrayControlWords[i]))
                    {
                        switch (i)
                        {
                            case 0:
                                {
                                    _sectionNumber++;
                                    _pictureNumber = 0;
                                    _tableNumber = 0;
                                    if (_sectionNumber != 1)
                                    {
                                        document.NewPage();
                                    }
                                    textParagraph = textParagraph.Replace(arrayControlWords[i], _sectionNumber.ToString());
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 13f, Font.BOLD));          
                                    iparagraph.Alignment = Element.ALIGN_CENTER;
                                    iparagraph.SpacingAfter = 15f;
                                    iparagraph.SpacingBefore = 0;
                                    iparagraph.FirstLineIndent = 0;                                                                 
                                }
                                break;
                            case 1:
                                {
                                    _pictureNumber++;
                                    textParagraph = textParagraph.Replace(arrayControlWords[i], "Рисунок " + _sectionNumber.ToString() + "." + _pictureNumber.ToString() + " - ");
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 12f, Font.ITALIC));
                                    iparagraph.Alignment = Element.ALIGN_CENTER;
                                    iparagraph.SpacingAfter = 12f;
                                    iparagraph.SpacingBefore = 12f;
                                    iparagraph.FirstLineIndent = 0;
                                }
                                break;
                            case 2:
                                {
                                    _tableNumber++;
                                    textParagraph = textParagraph.Replace(arrayControlWords[i], "Таблица " + _sectionNumber.ToString() + "." + _tableNumber.ToString() + " - ");
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 12f, Font.ITALIC));
                                    iparagraph.Alignment = Element.ALIGN_LEFT;
                                    iparagraph.SpacingAfter = 12f;
                                    iparagraph.SpacingBefore = 12f;
                                    iparagraph.FirstLineIndent = 0;
                                }
                                break;
                            case 3:
                                {
                                    textParagraph = textParagraph.Replace(arrayControlWords[i], _sectionNumber.ToString() + "." + (_pictureNumber + 1).ToString());
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 12f, Font.NORMAL));
                                    iparagraph.Alignment = Element.ALIGN_JUSTIFIED;
                                    iparagraph.SpacingAfter = 0;
                                    iparagraph.SpacingBefore = 0;
                                    iparagraph.FirstLineIndent = 0;
                                }
                                break;
                            case 4:
                                {
                                    textParagraph = textParagraph.Replace(arrayControlWords[i], _sectionNumber.ToString() + "." + (_pictureNumber).ToString());
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 12f, Font.NORMAL));
                                    iparagraph.Alignment = Element.ALIGN_JUSTIFIED;
                                    iparagraph.SpacingAfter = 0;
                                    iparagraph.SpacingBefore = 0;
                                    iparagraph.FirstLineIndent = 0;
                                }
                                break;
                            case 5:
                                {
                                    textParagraph = textParagraph.Replace(arrayControlWords[i], _sectionNumber.ToString() + "." + (_tableNumber + 1).ToString());
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 12f, Font.NORMAL));
                                    iparagraph.Alignment = Element.ALIGN_JUSTIFIED;
                                    iparagraph.SpacingAfter = 0;
                                    iparagraph.SpacingBefore = 0;
                                    iparagraph.FirstLineIndent = 0;
                                }
                                break;
                            case 6:
                                {
                                    iparagraph = null;
                                    string csvPath = textParagraph.Replace(arrayControlWords[i], "") .Replace("*", "").Replace("\r", "").Replace("]", "");
                                    csvPath = new System.IO.FileInfo(path).DirectoryName + "\\" + csvPath;
                                    string[] listRows = System.IO.File.ReadAllLines(csvPath);
                                    string[] listTitle = listRows[0].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                                    PdfPTable table = new PdfPTable(listTitle.Length);
                                    for (var k = 0; k < listTitle.Length; k++)
                                    {
                                        PdfPCell cell = new PdfPCell(new Phrase(listTitle[k].ToString(), new Font(BfArial, 12f, Font.BOLD, iTextSharp.text.Color.RED)));
                                        cell.Column.Alignment = Element.ALIGN_CENTER;
                                        cell.FixedHeight = 30f;
                                        table.AddCell(cell);
                                    }
                                    for (var j = 1; j < listRows.Length; j++)
                                    {
                                        string[] listValues = listRows[j].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                                        for (var k = 0; k < listValues.Length; k++)
                                        {
                                            PdfPCell cell = new PdfPCell(new Phrase(listValues[k].ToString(),  new Font(BfArial, 12f, Font.BOLD, iTextSharp.text.Color.GREEN)));
                                            cell.BackgroundColor = iTextSharp.text.Color.BLACK;
                                            cell.BorderColor = iTextSharp.text.Color.WHITE;
                                            cell.FixedHeight = 30f;
                                            cell.Column.Alignment = Element.ALIGN_CENTER;
                                            table.AddCell(cell);
                                        }
                                    }
                                    table.HorizontalAlignment = Element.ALIGN_LEFT;
                                    // Ширина в процентах
                                    table.WidthPercentage = 100;
                                    //table.
                                    //добавляем таблицу в документ
                                    document.Add(table);

                                }
                                break;
                            case 7:
                                {
                                    string replaceString = "";
                                    for (int j = 0; j < sourceList.Count; j++)
                                    {
                                        replaceString = (j + 1).ToString() + ". " + sourceList[j].TrimStart('[').TrimEnd(']') + "\r\n";
                                        iparagraph = new iTextSharp.text.Paragraph(replaceString, new Font(BfArial, 12f, Font.NORMAL));
                                        iparagraph.SpacingAfter = 0;
                                        iparagraph.SpacingBefore = 0;
                                        iparagraph.FirstLineIndent = 20f;
                                        iparagraph.ExtraParagraphSpace = 10;
                                        iparagraph.Alignment = Element.ALIGN_JUSTIFIED;
                                        document.Add(iparagraph);
                                    }
                                }
                                break;
                            case 8:
                                {
                                    textParagraph = "тут будет ваш код";
                                    iparagraph = new iTextSharp.text.Paragraph(textParagraph, new Font(BfArial, 12f, Font.NORMAL));
                                }
                                break;
                            case 9:
                                {
                                    iparagraph = null;
                                    string jpgPath = textParagraph.Replace(arrayControlWords[i],"").Replace("*", "").Replace("\r", "").Replace("]", "");
                                    jpgPath = new System.IO.FileInfo(path).DirectoryName+ "\\" + jpgPath;
                                    Image jpg = Image.GetInstance(jpgPath);
                                    jpg.Alignment = Element.ALIGN_CENTER;
                                    jpg.SpacingBefore = 12f;

                                    float procent = 90;
                                    while (jpg.ScaledWidth > PageSize.A4.Width / 2.0f)
                                    {
                                        jpg.ScalePercent(procent);
                                        procent -= 10;
                                    }

                                    document.Add(jpg);

                                }
                                break;                                    
                        }
                    }
                }              

                if (iparagraph !=null)
                    document.Add(iparagraph);
            }

            document.Close();
        }

        private string SearchSource(string text)
        {
            if (text.Contains("["))
            {
                for (int j=0; j < text.Length; j++)
                {
                    if (text[j] == '[' && text[j + 1] != '*')
                    {
                        int startIndex = j;
                        int endIndex = startIndex + 1;
                        while (endIndex < text.Length && text[endIndex] != ']')
                        {
                            endIndex++;
                        }
                        string sourceName = "";

                        if (text[endIndex] == ']')
                        {
                            for (int k = startIndex; k <= endIndex; k++)
                            {
                                sourceName += text[k];
                            }
                            int index = 0;
                            if (!sourceList.Contains(sourceName))
                            {
                                //добавляем в список, увеличиваем номер текущей ссылки
                                sourceList.Add(sourceName);
                                _sourceNumber++;
                                index = _sourceNumber;
                            }
                            else
                            {
                                for (int k = 0; k < sourceList.Count; k++)
                                {
                                    //то находим его номер
                                    if (sourceList[k].Contains(sourceName))
                                    {
                                        index = k + 1;
                                    }
                                }
                            }
                            string replaceString = "[" + index.ToString() + "]";
                            text = text.Replace(sourceName, replaceString);
                            j = endIndex;
                        }
                    }
                }
            }

            return text;
        }

        private void ProcessStart(string path)
        {
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(path)
            {
                UseShellExecute = true
            };
            p.Start();
        }
    }
}
