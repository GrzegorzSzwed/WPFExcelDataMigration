using Microsoft.Win32;
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
using WPFUIExcelDataCapture.Models;
using System.IO;
using ExcelDataCapture;
using TextMatchCalculation;


namespace WPFUIExcelDataCapture.Views
{
    /// <summary>
    /// Interaction logic for SettingsView.xaml
    /// </summary>
    public partial class SettingsView : UserControl
    {
        private MigrationContainer _migrationContainer;
        public SettingsView(MigrationContainer migration)
        {
            InitializeComponent();

            _migrationContainer = migration;

            try
            {
                if (_migrationContainer.Destination != null)
                {
                    cmbExcelDestination.ItemsSource = _migrationContainer.Destination.ListOfWorksheets();
                    txtDestinationFileName.Text = _migrationContainer.Destination.FileName;
                    cmbExcelDestination.Text = _migrationContainer.Destination.CurrentWorksheet;
                    txtColumnDestinationCount.Text = _migrationContainer.Destination.Columns.ToString();
                    txtRowDestinationCount.Text =  (_migrationContainer.Destination.Rows - _migrationContainer.Destination.ColumnCaptionIndex).ToString();
                    txtSimililarityPercent.Text = Similarity();
                }

                if (_migrationContainer.Source != null)
                {
                    cmbExcelSource.ItemsSource = _migrationContainer.Source.ListOfWorksheets();
                    txtSourceFileName.Text = _migrationContainer.Source.FileName;
                    cmbExcelSource.Text = _migrationContainer.Source.CurrentWorksheet;
                    txtColumnSourceCount.Text = _migrationContainer.Source.Columns.ToString();
                    txtRowSourceCount.Text = (_migrationContainer.Source.Rows - _migrationContainer.Source.ColumnCaptionIndex).ToString();
                    txtSimililarityPercent.Text = Similarity();
                }
            }
            catch
            {
                var msg = new MessageView("Could not load worksheets");
                msg.Show();
            }
        }

        private void BtnLoadExcelSource_Click(object sender, RoutedEventArgs e)
        {
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsm|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true)
            {
                filePath = openFileDialog.FileName;
                _migrationContainer.Source = new NavisionExcel(filePath);
                txtSourceFileName.Text = System.IO.Path.GetFileNameWithoutExtension(filePath);
                cmbExcelSource.ItemsSource = _migrationContainer.Source.ListOfWorksheets();
                cmbExcelSource.Text = _migrationContainer.Source.CurrentWorksheet;
                txtColumnSourceCount.Text = _migrationContainer.Source.Columns.ToString();
                txtRowSourceCount.Text = (_migrationContainer.Source.Rows - _migrationContainer.Source.ColumnCaptionIndex).ToString();
                txtSimililarityPercent.Text = Similarity();
            }
        }

        private void BtnExcelDestination_Click(object sender, RoutedEventArgs e)
        {
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsm|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true)
            {
                filePath = openFileDialog.FileName;
                _migrationContainer.Destination = new NavisionExcel(filePath);
                txtDestinationFileName.Text = System.IO.Path.GetFileNameWithoutExtension(filePath);
                cmbExcelDestination.ItemsSource = _migrationContainer.Destination.ListOfWorksheets();
                cmbExcelDestination.Text = _migrationContainer.Destination.CurrentWorksheet;
                txtColumnDestinationCount.Text = _migrationContainer.Destination.Columns.ToString();
                txtRowDestinationCount.Text = (_migrationContainer.Destination.Rows - _migrationContainer.Destination.ColumnCaptionIndex).ToString();
                txtSimililarityPercent.Text = Similarity();
            }
        }

        private void CmbExcelSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_migrationContainer.Source.ChangeWorksheet(cmbExcelSource.SelectedItem.ToString()))
            {
                txtColumnSourceCount.Text = _migrationContainer.Source.Columns.ToString();
                txtRowSourceCount.Text = (_migrationContainer.Source.Rows - _migrationContainer.Source.ColumnCaptionIndex).ToString();
                txtSimililarityPercent.Text = Similarity();
            }
            else
            {
                var msg = new MessageView("Source worksheet cannot be loaded");
            }
            
        }

        private void CmbExcelDestination_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_migrationContainer.Destination.ChangeWorksheet(cmbExcelDestination.SelectedItem.ToString()))
            {
                txtColumnDestinationCount.Text = _migrationContainer.Destination.Columns.ToString();
                txtRowDestinationCount.Text = (_migrationContainer.Destination.Rows - _migrationContainer.Destination.ColumnCaptionIndex).ToString();
                txtSimililarityPercent.Text = Similarity();
            }
            else
            {
                var msg = new MessageView("Destination worksheet cannot be loaded");
            }

        }

        private string Similarity()
        {
            if (_migrationContainer.Destination != null && _migrationContainer.Source != null)
            {
                return Levenstein.Percent(_migrationContainer.Source.CurrentWorksheet,
                    _migrationContainer.Destination.CurrentWorksheet).ToString() + "%";
            }
            else
                return "0%";
        }
    }
}
