using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Newtonsoft.Json;
using static Template4338.MainWindow;
using Microsoft.Office.Interop.Excel;

namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void b2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (sadEntities entities = new sadEntities())
            {
                for (int i = 1; i < _rows - 1; i++)
                {
                    entities.isrpo.Add(new isrpo(list[i, 0], list[i, 1], list[i, 2], list[i, 3], list[i, 4], list[i, 5], list[i, 6], list[i, 7], list[i, 8]));
                }
                entities.SaveChanges();
            }
        }
        person[] p;

        private void b3_Click(object sender, RoutedEventArgs e)
        {
            List<isrpo> allStudents;
            using (sadEntities entities = new sadEntities())
            {
                allStudents =
                entities.isrpo.ToList().OrderBy(s =>
                s.id).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
           
           
            foreach(isrpo s in allStudents)
            {
                if (s.DAtta != "")
                {
                    //MessageBox.Show(s.DAtta);
                    DateTime birthDate = DateTime.ParseExact(s.DAtta, "dd.MM.yyyy", null);
                    TimeSpan age = DateTime.Today - birthDate;
                    int ageInYears = (int)age.TotalDays / 365;
                    s.vozrast = ageInYears;
                }
            }

            var groupedPeople = allStudents.GroupBy(p => GetCategory(p.vozrast)).ToList();
            var cat1 = allStudents.Where(p => GetCategory(p.vozrast) == 1);
            var cat2 = allStudents.Where(p => GetCategory(p.vozrast) == 2);
            var cat3 = allStudents.Where(p => GetCategory(p.vozrast) == 3);
            for (int i = 0; i < 3; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(i + 1);
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "EMAIL";
                startRowIndex++;
                foreach (var s in allStudents)
                {
                    if(GetCategory(s.vozrast) == i+1)
                    {
                        worksheet.Cells[1][startRowIndex] = s.KOD;
                        worksheet.Cells[2][startRowIndex] = s.FIO;
                        worksheet.Cells[3][startRowIndex] = s.email;
                        startRowIndex++;
                    }
                }
               

            }
            app.Visible = true;


        }
        public static int GetCategory(int age)
        {
            if (age >= 20 && age <= 29)
                return 1;
            else if (age >= 30 && age <= 39)
                return 2;
            else if (age >= 40)
                return 3;
            else
                return 5;
        }
        public class person
        {
            public int Id { get; set; }
            public string FullName { get; set; }
            public string CodeClient { get; set; }
            public string BirthDate { get; set; }
            public string Index { get; set; }
            public string City { get; set; }
            public string Street { get; set; }
            public string Home { get; set; }
            public string Kvartira { get; set; }
            public string E_mail { get; set; }
        }

        private void b5_Click(object sender, RoutedEventArgs e)
        {
            List<isrpo> isrpos;
            using (sadEntities entities = new sadEntities())
            {
                isrpos =
                entities.isrpo.ToList().OrderBy(s =>
                s.id).ToList();
            }
            foreach (isrpo s in isrpos)
            {
                if (s.DAtta != "")
                {
                    //MessageBox.Show(s.DAtta);
                    DateTime birthDate = DateTime.ParseExact(s.DAtta, "dd.MM.yyyy", null);
                    TimeSpan age = DateTime.Today - birthDate;
                    int ageInYears = (int)age.TotalDays / 365;
                    s.vozrast = ageInYears;
                }
            }
            
            var groupedPeople = isrpos.GroupBy(p => GetCategory(p.vozrast)).ToList();
            var cat1 = isrpos.Where(p => GetCategory(p.vozrast) == 1);
            var cat2 = isrpos.Where(p => GetCategory(p.vozrast) == 2);
            var cat3 = isrpos.Where(p => GetCategory(p.vozrast) == 3);
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();
            using (sadEntities usersEntities = new sadEntities())
            {
                int i = 1;
                foreach (var group in groupedPeople)
                {
                    int j = 1;
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(i);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph =
                    document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table studentsTable =
                    document.Tables.Add(tableRange, group.Count() + 1, 3);
                    studentsTable.Borders.InsideLineStyle =
                    studentsTable.Borders.OutsideLineStyle =
                    Word.WdLineStyle.wdLineStyleSingle;
                    studentsTable.Range.Cells.VerticalAlignment =
                    Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    Word.Range cellRange;
                    cellRange = studentsTable.Cell(1, 1).Range;
                    cellRange.Text = "ФИО";
                    cellRange = studentsTable.Cell(1, 2).Range;
                    cellRange.Text = "КОД";
                    cellRange = studentsTable.Cell(1, 3).Range;
                    cellRange.Text = "email";
                    studentsTable.Rows[1].Range.Bold = 1;
                    studentsTable.Rows[1].Range.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    foreach (var s in isrpos)
                    {
                        if (GetCategory(s.vozrast) == i)
                        {
                            cellRange = studentsTable.Cell(j + 1, 1).Range;
                            cellRange.Text = s.FIO.ToString();
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = studentsTable.Cell(j + 1, 2).Range;
                            cellRange.Text = s.KOD;
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = studentsTable.Cell(j + 1, 3).Range;
                            cellRange.Text = s.email;
                            cellRange.ParagraphFormat.Alignment =
                            Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            j++;
                        }
                    }
                    i++;
                }
            }
            
            app.Visible = true;
        }

        private void b4_Click(object sender, RoutedEventArgs e)
        {
            string json = File.ReadAllText("D:\\лабы\\Импорт\\3.json");
            p = JsonConvert.DeserializeObject<person[]>(json);
            using (sadEntities entities = new sadEntities())
            {
                entities.SaveChanges();
                foreach (person p in p)
                {
                    if (p.FullName != null)
                        entities.isrpo.Add(new isrpo(p.FullName, p.CodeClient, p.BirthDate, p.Index, p.City, p.Street, p.Home, p.Kvartira, p.E_mail));
                }
                entities.SaveChanges();

            }

        }
    }
}
