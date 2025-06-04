using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для SchtatRaspPage.xaml
    /// </summary>
    public partial class SchtatRaspPage : Page
    {
        public SchtatRaspPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddSchtatRasp(null));
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddSchtatRasp((sender as System.Windows.Controls.Button).DataContext as SchtatnoeRaspisanie));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = RaspDG.SelectedItems.Cast<SchtatnoeRaspisanie>().ToList();
            if (System.Windows.MessageBox.Show($"Удалить {delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Connect.context.SchtatnoeRaspisanie.RemoveRange(delClients);
            }
            try
            {
                Connect.context.SaveChanges();
                RaspDG.ItemsSource = Connect.context.SchtatnoeRaspisanie.ToList();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            RaspDG.ItemsSource = Connect.context.SchtatnoeRaspisanie.ToList();
        }
    }
}
