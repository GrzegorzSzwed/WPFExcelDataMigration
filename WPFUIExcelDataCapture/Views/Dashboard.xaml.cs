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
using WPFUIExcelDataCapture.Views;
using WPFUIExcelDataCapture.Models;

namespace WPFUIExcelDataCapture
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MigrationContainer migrationContainer;
        public MainWindow()
        {
            InitializeComponent();
            migrationContainer = new MigrationContainer();
            DashboardContent.Content = new HomeView();
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            if (migrationContainer.Destination != null)
                migrationContainer.Destination.Save();

            if (migrationContainer.Source != null)
                migrationContainer.Source.Save();

            App.Current.Shutdown();
        }

        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void BtnHome_Click(object sender, RoutedEventArgs e)
        {
            DashboardContent.Content = new HomeView();
        }

        private void BtnData_Click(object sender, RoutedEventArgs e)
        {
            DashboardContent.Content = new ComparisonView(migrationContainer);
        }

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            DashboardContent.Content = new SettingsView(migrationContainer);
        }
    }
}
