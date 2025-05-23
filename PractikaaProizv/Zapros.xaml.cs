using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
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
namespace PractikaaProizv {
    /// <summary>
    /// Логика взаимодействия для Zapros.xaml
    /// </summary>
    public partial class Zapros : Page
    {
        public SqlConnection Connection = new SqlConnection(@"Data source=(LocalDB)\MSSQLLocalDB;attachdbfilename=|DataDirectory|\Database1.mdf;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework");
        public Zapros()
        {
            InitializeComponent();
        }
        private void LoadComboBox()
        {
            string sql = "SELECT o.Familia FROM [dbo].[RavocheeVremya] vo JOIN [dbo].[Sotrudniki] o ON vo.IdSotr = o.IdSotr";
            using (SqlCommand cmd = new SqlCommand(sql, Connection))
            {
                cmd.CommandType = System.Data.CommandType.Text;
                DataTable table = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(table);
                List<string> namesList = new List<string>();
                foreach (DataRow row in table.Rows)
                {
                    namesList.Add(row["Familia"].ToString());
                }
                ZaprBox.ItemsSource = namesList;
            }
        }
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            LoadComboBox();
        }
        private void Zapr_Click(object sender, RoutedEventArgs e)
        {
            string a = Convert.ToString(ZaprBox.SelectionBoxItem);
            var select = Connect.context.RavocheeVremya.Select(x =>
            new
            {
                RavocheeVremya = x,
                Sotrudniki = x.Sotrudniki,
            }).Where(x => x.Sotrudniki.Familia == a).ToList();
            ZaprosDG.ItemsSource = select;
        }
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            ZaprosDG.ItemsSource = Connect.context.RavocheeVremya.Select(x =>
            new
            {
                RavocheeVremya = x,
                Sotrudniki = x.Sotrudniki,
            }).ToList();
        }
    }
}
