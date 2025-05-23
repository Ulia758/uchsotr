using System;
using System.Collections.Generic;
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

namespace PractikaaProizv
{
    /// <summary>
    /// Логика взаимодействия для DolgnostiPage.xaml
    /// </summary>
    public partial class DolgnostiPage : Page
    {
        public DolgnostiPage()
        {
            InitializeComponent();
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddDolgnosti(null));
        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var delClients = DolgnostiDG.SelectedItems.Cast<Dolgnosti>().ToList();
            foreach (var delClient in delClients)
                if (Connect.context.Dolgnosti.Any(x => x.OtdelId == delClient.OtdelId))
                {
                    MessageBox.Show("Данные используются в другой таблице", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            if (MessageBox.Show($"Удалить {delClients.Count} записей", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Connect.context.Dolgnosti.RemoveRange(delClients);
            }
            try
            {
                Connect.context.SaveChanges();
                DolgnostiDG.ItemsSource = Connect.context.Dolgnosti.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddDolgnosti((sender as Button).DataContext as Dolgnosti));
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            DolgnostiDG.ItemsSource = Connect.context.Dolgnosti.ToList();
        }
        private void Poisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Connect.context != null && DolgnostiDG != null)
            {
                DolgnostiDG.ItemsSource = Connect.context.Dolgnosti.Where(x => x.Nazvanie.StartsWith(Poisk.Text)).ToList();
            }
        }
    }
}
