using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Contexts;
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
    /// Логика взаимодействия для UvolnPage.xaml
    /// </summary>
    /// 
    public partial class UvolnPage : Page
    {
        public ObservableCollection<Uvolnenie> UvolnenieItems { get; set; } = new ObservableCollection<Uvolnenie>();
        public UvolnPage()
        {
            InitializeComponent();
            this.DataContext = this; // Устанавливаем DataContext на страницу
            RefreshUvolnenieData();
        }
        private void RefreshUvolnenieData()
        {
            UvolnenieItems.Clear();
            foreach (var item in Connect.context.Uvolnenie.ToList())
            {
                UvolnenieItems.Add(item);
            }
        }
        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new AddUvoln((sender as System.Windows.Controls.Button).DataContext as Uvolnenie));
        }
        private void Nazad_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.GoBack();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshUvolnenieData();
        }

    }
}
