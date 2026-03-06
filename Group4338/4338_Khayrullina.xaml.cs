using System.Windows;
using System.Data;

namespace Group4338
{
    public partial class _4338_Khayrullina : Window
    {
        private DataTable importedData;

        public _4338_Khayrullina()
        {
            InitializeComponent();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Тут будет загрузка из Excel", "Импорт");

            importedData = new DataTable();
            importedData.Columns.Add("Код клиента");
            importedData.Columns.Add("ФИО");
            importedData.Columns.Add("E-mail");
            importedData.Columns.Add("Улица проживания");

            importedData.Rows.Add("1", "Хайруллина Ландыш", "landysh@mail.ru", "Красной позиции");
            importedData.Rows.Add("2", "Фахриева Марьям", "maryam@mail.ru", "Белой позиции");
            importedData.Rows.Add("3", "Тагирова Гайша", "gaisha@mail.ru", "Черной позиции");

            MessageBox.Show($"Загружено {importedData.Rows.Count} записей", "Импорт завершен");
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (importedData == null)
            {
                MessageBox.Show("Сначала импортируйте данные!", "Ошибка");
                return;
            }

            MessageBox.Show("Тут будет экспорт в Excel", "Экспорт");

            string streets = "";
            foreach (DataRow row in importedData.Rows)
            {
                streets += row["Улица проживания"] + "\n";
            }
            MessageBox.Show($"Улицы для группировки:\n{streets}", "Экспорт");
        }
    }
}