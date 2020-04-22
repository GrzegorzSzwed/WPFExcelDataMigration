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
using TextMatchCalculation;
using AsposeExcelDataCapture;

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
                    InitSourceExcel();
                }

                if (_migrationContainer.Source != null)
                {
                    cmbExcelSource.ItemsSource = _migrationContainer.Source.ListOfWorksheets();
                    InitDestinationExcel();
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
                _migrationContainer.Source = new AsposeExcel(filePath);
                InitSourceExcel();
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
                _migrationContainer.Destination = new AsposeExcel(filePath);
                InitDestinationExcel();
            }
        }

        private void CmbExcelSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_migrationContainer.Source.ChangeWorksheet(cmbExcelSource.SelectedItem.ToString()))
            {
                txtColumnSourceCount.Text = _migrationContainer.Source.ColumnsCount.ToString();
                txtRowSourceCount.Text = (_migrationContainer.Source.RowsCount - _migrationContainer.Source.ColumnCaptionRow).ToString();
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
                txtColumnDestinationCount.Text = _migrationContainer.Destination.ColumnsCount.ToString();
                txtRowDestinationCount.Text = (_migrationContainer.Destination.RowsCount - _migrationContainer.Destination.ColumnCaptionRow).ToString();
                txtSimililarityPercent.Text = Similarity();
            }
            else
            {
                var msg = new MessageView("Destination worksheet cannot be loaded");
            }

        }
        private void InitDestinationExcel()
        {
            txtDestinationFileName.Text = _migrationContainer.Destination.FileName;
            cmbExcelDestination.ItemsSource = _migrationContainer.Destination.ListOfWorksheets();
            cmbExcelDestination.Text = _migrationContainer.Destination.CurrentWorksheetName;
            txtColumnDestinationCount.Text = _migrationContainer.Destination.ColumnsCount.ToString();
            txtRowDestinationCount.Text = (_migrationContainer.Destination.RowsCount - _migrationContainer.Destination.ColumnCaptionRow).ToString();
            txtSimililarityPercent.Text = Similarity();
        }

        private void InitSourceExcel()
        {
            txtSourceFileName.Text = _migrationContainer.Source.FileName;
            cmbExcelSource.ItemsSource = _migrationContainer.Source.ListOfWorksheets();
            cmbExcelSource.Text = _migrationContainer.Source.CurrentWorksheetName;
            txtColumnSourceCount.Text = _migrationContainer.Source.ColumnsCount.ToString();
            txtRowSourceCount.Text = (_migrationContainer.Source.RowsCount - _migrationContainer.Source.ColumnCaptionRow).ToString();
            txtSimililarityPercent.Text = Similarity();
        }

        private string Similarity()
        {
            if (_migrationContainer.Destination != null && _migrationContainer.Source != null)
            {
                return Levenstein.Percent(_migrationContainer.Source.CurrentWorksheetName,
                    _migrationContainer.Destination.CurrentWorksheetName).ToString() + "%";
            }
            else
                return "0%";
        }

        private void LoadTemplate_Click(object sender, RoutedEventArgs e)
        {
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (_migrationContainer.Destination != null)
                openFileDialog.InitialDirectory = _migrationContainer.Destination.FileDirectory;
            else
                openFileDialog.InitialDirectory = "C:\\";

            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsm|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true)
            {
                filePath = openFileDialog.FileName;
                AsposeExcel Template = new AsposeExcel(filePath);
                LoadFromTemplate(Template);
            }
            else
            {
                var msg = new MessageView("Please choose some template file");
            }
        }

        private void LoadFromTemplate(AsposeExcel template)
        {
            string sourcefullpath = template.ReadCell(0, 1);
            string destinationfullpath = template.ReadCell(1, 1);
            string sourceworksheet = template.ReadCell(0, 2);
            string destinationworksheet = template.ReadCell(1, 2);

            if (File.Exists(sourcefullpath) && File.Exists(destinationfullpath))
            {
                _migrationContainer.Source = new AsposeExcel(sourcefullpath);
                InitSourceExcel();
                _migrationContainer.Destination = new AsposeExcel(destinationfullpath);
                InitDestinationExcel();

                if(_migrationContainer.Source.ChangeWorksheet(sourceworksheet) && _migrationContainer.Destination.ChangeWorksheet(destinationworksheet))
                {
                    _migrationContainer.TemplateAttached = true;

                    for(int crow = 3; crow < template.RowsCount; crow++)
                    {
                        ColumnParser col = new ColumnParser();
                        col.SourceColumnName = template.ReadCell(crow, 0);
                        col.SourceColumnIndex = Convert.ToInt16(template.ReadCell(crow, 1));
                        col.DestinationColumnName = template.ReadCell(crow, 2);
                        col.DestinationColumnIndex = Convert.ToInt16(template.ReadCell(crow, 3));
                        col.IsKey = Convert.ToBoolean(template.ReadCell(crow, 4));
                        col.LookupMatch = Convert.ToBoolean(template.ReadCell(crow, 5));

                        _migrationContainer.ColumnParsers.Add(col);
                    }
                    if (_migrationContainer.ColumnParsers.Count() > 0)
                    {
                        var msg = new MessageView("Template has been loaded");
                        msg.Show();
                    }
                    else
                    {
                        var msg = new MessageView("Cannot find any column relations");
                        msg.Show();
                    }
                    
                }
                else
                {
                    var msg = new MessageView("Cannot find worksheets");
                    msg.Show();
                }
            }
            else
            {
                var msg = new MessageView("Cannot find files. Put them in the same directory as written in excel file");
                msg.Show();
            }
            
        }
    }
}
