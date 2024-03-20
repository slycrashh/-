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

namespace Template4338
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
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
    }
}
