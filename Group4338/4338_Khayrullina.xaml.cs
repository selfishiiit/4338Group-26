using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Data.SQLite;

namespace Group4338
{
    public partial class _4338_Khayrullina : Window
    {
        private DataTable importedData;

        public _4338_Khayrullina()
        {
            InitializeComponent();
            // Для старых версий EPPlus
            ExcelPackage.License.SetNonCommercialPersonal("Ландыш Хайруллина");
        }

        // КНОПКА ИМПОРТА
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.Title = "Выберите файл 3.xlsx";

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    importedData = LoadFromExcel(filePath);
                    SaveToDatabase(importedData);
                    MessageBox.Show($"Импортировано {importedData.Rows.Count} записей!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        // КНОПКА ЭКСПОРТА
        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (importedData == null || importedData.Rows.Count == 0)
                {
                    MessageBox.Show("Сначала импортируйте данные!", "Предупреждение");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = "export_po_ulicam.xlsx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    GroupAndExportToExcel(importedData, filePath);
                    MessageBox.Show("Экспорт завершен!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        // ЗАГРУЗКА ИЗ EXCEL
        private DataTable LoadFromExcel(string filePath)
        {
            DataTable dt = new DataTable();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    dt.Columns.Add(worksheet.Cells[1, col].Text);

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    dt.Rows.Add(dataRow);
                }
            }
            return dt;
        }

        // СОХРАНЕНИЕ В БД
        private void SaveToDatabase(DataTable data)
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "clients.db");
            using (var connection = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                connection.Open();
                using (var command = new SQLiteCommand(@"
                    CREATE TABLE IF NOT EXISTS Clients (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Код_клиента TEXT, ФИО TEXT, Email TEXT, Улица_проживания TEXT)", connection))
                {
                    command.ExecuteNonQuery();
                }

                using (var command = new SQLiteCommand("DELETE FROM Clients", connection))
                    command.ExecuteNonQuery();

                foreach (DataRow row in data.Rows)
                {
                    using (var command = new SQLiteCommand(@"
                        INSERT INTO Clients (Код_клиента, ФИО, Email, Улица_проживания) 
                        VALUES (@code, @fio, @email, @street)", connection))
                    {
                        command.Parameters.AddWithValue("@code", row[0].ToString());
                        command.Parameters.AddWithValue("@fio", row[1].ToString());
                        command.Parameters.AddWithValue("@email", row[2].ToString());
                        command.Parameters.AddWithValue("@street", row[3].ToString());
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private void GroupAndExportToExcel(DataTable data, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var groups = data.AsEnumerable().GroupBy(row => row[3].ToString()).OrderBy(g => g.Key);

                foreach (var group in groups)
                {
                    string streetName = string.IsNullOrEmpty(group.Key) ? "Без улицы" : group.Key;
                    var worksheet = package.Workbook.Worksheets.Add(streetName);

                    worksheet.Cells[1, 1].Value = "Код клиента";
                    worksheet.Cells[1, 2].Value = "ФИО";
                    worksheet.Cells[1, 3].Value = "E-mail";

                    int row = 2;
                    foreach (var item in group)
                    {
                        worksheet.Cells[row, 1].Value = item[0].ToString();
                        worksheet.Cells[row, 2].Value = item[1].ToString();
                        worksheet.Cells[row, 3].Value = item[2].ToString();
                        row++;
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
                package.SaveAs(new FileInfo(filePath));
            }
        }

        private int GetUniqueStreetsCount(DataTable data)
        {
            return data.AsEnumerable().Select(row => row[3].ToString()).Distinct().Count();
        }
    }
}