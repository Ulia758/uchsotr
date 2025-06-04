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
    /// Логика взаимодействия для MainPage.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void SotrudnikiBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new SotrudnikiPage());
        }
        private void OtdeliBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new OtdeliPage());
        }
        private void PosechaemostBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new PosechaemostPage());
        }
        private void DolgnostiBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new DolgnostiPage());
        }
        private void OtpuskaBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new OtpuskaPage());
        }
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void OtchetiBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new OtchetiPage());
        }

        private void KvalificaciyaBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new PovKvalPage());
        }

        private void RaspisanieBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new SchtatRaspPage());
        }

        private void UvolnenieBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new UvolnPage());
        }

        private void ZachislenieBtn_Click(object sender, RoutedEventArgs e)
        {
            Nav.MainFrame.Navigate(new ZachislPage());
        }
    }
}
