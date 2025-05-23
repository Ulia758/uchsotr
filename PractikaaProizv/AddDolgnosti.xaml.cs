using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
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
    /// Логика взаимодействия для AddDolgnosti.xaml
    /// </summary>
    public partial class AddDolgnosti : Page
    {
        Dolgnosti k;
        public AddDolgnosti(Dolgnosti c)
        {
            InitializeComponent();
            if (c == null)
                c = new Dolgnosti();
            DataContext = k = c;
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (k.IdDolgnost == 0)
            {
                Connect.context.Dolgnosti.Add(k);
            }
            try
            {
                Connect.context.SaveChanges();
            }
            catch (DbEntityValidationException ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Nav.MainFrame.GoBack();
        }
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
    }
}
