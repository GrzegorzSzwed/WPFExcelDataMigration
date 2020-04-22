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
using ExcelDataCapture;
using WPFUIExcelDataCapture.Models;
using TextMatchCalculation;

namespace WPFUIExcelDataCapture.Views
{
    /// <summary>
    /// Interaction logic for ComparisonView.xaml
    /// </summary>
    public partial class ComparisonView : UserControl
    {
        private MigrationContainer _migration;
        public List<ColumnParser> columnParsers = new List<ColumnParser>();
        public string DestinationCurrentWorksheet;
        public string SourceCurrentWorksheet;

        public ComparisonView(MigrationContainer migration)
        {
            InitializeComponent();
            _migration = migration;
            MergeByMatch.IsChecked = false;
            Merge.IsChecked = false;
            Overwrite.IsChecked = true;

            if(_migration.Source != null && _migration.Destination != null)
            {
                LoadComparisonView();
            }
            else
            {
                var msg = new MessageView("Navision Excel files should be loaded!");
                msg.Show();
            }
        }

        private void LoadComparisonView()
        {
            if (_migration.Destination != null)
            {
                var captionsWithIndex = _migration.Destination.ColumnCaptionsWithIndex;
                foreach (var col in captionsWithIndex)
                {
                    ColumnParser parser = new ColumnParser();
                    parser.DestinationColumnName = col.Key;
                    parser.DestinationColumnIndex = col.Value;
                    parser.DestinationColumnCaptionList = _migration.Destination.ColumnCaptions;
                    parser.SourceColumnCaptionList = _migration.Source.ColumnCaptions;
                    FillSourceIfFoundDestination(parser, col.Key);

                    columnParsers.Add(parser);
                }
                itemsListColumnParsers.ItemsSource = columnParsers;
            }
        }

        public void FillSourceIfFoundDestination(ColumnParser parser, string caption)
        {
            if(_migration.Source.ColumnCaptionsWithIndex.ContainsKey(caption))
            {
                parser.SourceColumnName = caption;
                parser.SourceColumnIndex = _migration.Source.ColumnCaptionsWithIndex[caption];
            }
            else
            {
                if (_migration.Source.ColumnCaptionsWithIndex.ContainsKey(caption))
                {
                    parser.SourceColumnName = caption;
                    parser.SourceColumnIndex = _migration.Source.ColumnCaptionsWithIndex[caption.ToLower()];
                }
                else
                {
                    string highest = string.Empty;
                    int memory = 0;
                    foreach (var par in _migration.Source.ColumnCaptions)
                    {
                        int distance = Levenstein.NettoDistance(par, caption);
                        if (distance < memory)
                        {
                            highest = par;
                            memory = distance;
                        }
                    }

                    if (memory <= 3)
                    {
                        parser.SourceColumnName = highest;
                        parser.SourceColumnIndex = _migration.Source.ColumnCaptionsWithIndex[highest];
                    }
                }
            }
        }

        private void CmbSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = sender as ComboBox;

        }

        private void BtnMerge_Click(object sender, RoutedEventArgs e)
        {
            if (columnParsers != null)
            {
                if(_migration.Destination != null && _migration.Source != null)
                {
                    if (Overwrite.IsChecked == true)
                    {
                        OverwriteDestinationBySource();
                    }

                    if (Merge.IsChecked == true)
                    {
                        MergeDestinationWithSource();
                    }

                    if(MergeWithRelation.IsChecked == true)
                    {
                        MergeAndCheckRelations();
                    }
                }
                else
                {
                    var msg = new MessageView("Please load some excels to merge them");
                    msg.Show();
                }
            }
        }

        private void MergeAndCheckRelations()
        {
            int columnparserscount = columnParsers.Count;
            if (IsSecondKeyChecked() && IsKeyChecked())
            {
                var keycolumn = from el in columnParsers
                                where el.IsKey == true
                                select el;

                if (keycolumn.Count() != 1)
                {
                    var msg = new MessageView("You cannot specify more than one first key!");
                    msg.Show();
                }
                else
                {
                    var col = keycolumn.FirstOrDefault();
                    Dictionary<string, int> sourceKeys = new Dictionary<string, int>();
                    Dictionary<string, int> destinationKeys = new Dictionary<string, int>();

                    sourceKeys = LoadColumn(_migration.Source, col.SourceColumnIndex);
                    destinationKeys = LoadColumn(_migration.Destination, col.DestinationColumnIndex);

                    SortedDictionary<string, int> sortedSourceKeys = new SortedDictionary<string, int>(sourceKeys);

                    int ifNotFoundKeyRow = LastEmptyRow(_migration.Destination);
                    bool permission = true;
                    int row = 0;

                    int columncaptionsindex = _migration.Destination.ColumnCaptionIndex;
                    foreach (var cl in sortedSourceKeys)
                    {
                        if (cl.Key != string.Empty)
                        {
                            if (destinationKeys.ContainsKey(cl.Key))
                            {
                                row = destinationKeys[cl.Key] + columncaptionsindex;
                            }
                            else
                            {
                                row = ifNotFoundKeyRow++;
                            }

                            FromSource2DestinationWithRelationByRow(row,
                                                   cl,
                                                   sourceKeys,
                                                   permission);
                        }
                    }

                    _migration.Destination.Save();
                    _migration.Source.Save();
                }
            }
        }
        private void FromSource2DestinationWithRelationByRow(int destinationRow, KeyValuePair<string, int> cl, Dictionary<string,int> sourceKeys, bool permission)
        {
            int columncaptionsindex = _migration.Destination.ColumnCaptionIndex;
            foreach (var columnparser in columnParsers)
            {
                if ((columnparser.SourceColumnIndex == 0 && permission == true) || columnparser.SourceColumnIndex > 0)
                {
                    if (columnparser.LookupMatch)
                    {
                        if (sourceKeys.Keys.Contains(_migration.Source.ReadCell(cl.Value, columnparser.SourceColumnIndex + 1)))
                        {
                            _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex + 1,
                                _migration.Source.ReadCell(cl.Value + columncaptionsindex, columnparser.SourceColumnIndex + 1));
                        }
                        else
                        {
                            _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex + 1, "");
                        }
                    }
                    else
                    {
                        _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex + 1,
                            _migration.Source.ReadCell(cl.Value + columncaptionsindex, columnparser.SourceColumnIndex + 1));
                    }
                }

                if (columnparser.SourceColumnIndex == 0)
                    permission = false; ;
            }
        }

        private void FromSource2DestinationByRow(int destinationRow, KeyValuePair<string, int> cl, Dictionary<string, int> sourceKeys, bool permission)
        {
            int columncaptionsindex = _migration.Destination.ColumnCaptionIndex;
            foreach (var columnparser in columnParsers)
            {
                if ((columnparser.SourceColumnIndex == 0 && permission == true) || columnparser.SourceColumnIndex > 0)
                {
                    _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex + 1,
                         _migration.Source.ReadCell(cl.Value + columncaptionsindex, columnparser.SourceColumnIndex + 1));
                }

                if (columnparser.SourceColumnIndex == 0)
                    permission = false; ;
            }
        }

        private void MergeDestinationWithSource()
        {
            int columnparserscount = columnParsers.Count;
            if (IsKeyChecked())
            {
                var keycolumn = from el in columnParsers
                                where el.IsKey == true
                                select el;

                if (keycolumn.Count() != 1)
                {
                    var msg = new MessageView("You cannot specify more than one key!");
                    msg.Show();
                }
                else
                {
                    var col = keycolumn.FirstOrDefault();
                    Dictionary<string, int> sourceKeys = new Dictionary<string, int>();
                    Dictionary<string, int> destinationKeys = new Dictionary<string, int>();

                    sourceKeys = LoadColumn(_migration.Source, col.SourceColumnIndex);
                    destinationKeys = LoadColumn(_migration.Destination, col.DestinationColumnIndex);

                    SortedDictionary<string, int> sortedSourceKeys = new SortedDictionary<string, int>(sourceKeys);
                    //SortedDictionary<string, int> sortedDestinationKeys = new SortedDictionary<string, int>(destinationKeys);

                    int ifNotFoundKeyRow = LastEmptyRow(_migration.Destination);
                    bool permission = true;
                    int row = 0;

                    int columncaptionsindex = _migration.Destination.ColumnCaptionIndex;
                    foreach (var cl in sortedSourceKeys)
                    {
                        if (cl.Key != string.Empty)
                        {
                            if (destinationKeys.ContainsKey(cl.Key))
                            {
                                row = destinationKeys[cl.Key] + columncaptionsindex;
                            }
                            else
                            {
                                row = ifNotFoundKeyRow++;
                            }

                            FromSource2DestinationByRow(row,
                                                   cl,
                                                   sourceKeys,
                                                   permission);
                        }
                    }

                    _migration.Destination.Save();
                    _migration.Source.Save();
                }
            }
        }

        private int LastEmptyRow(NavisionExcel destination)
        {
            return destination.Rows + 1;
        }

        private Dictionary<string, int> LoadColumn(NavisionExcel source, int sourceColumnIndex)
        {
            Dictionary<string, int> column = new Dictionary<string, int>();
            int emptykeyscounter = 0;
            for(int i = 1;i < source.Rows; i++)
            {
                var key = source.ReadCell(source.ColumnCaptionIndex + i, sourceColumnIndex + 1).ToLower();
                if (!column.ContainsKey(key))
                    column.Add(source.ReadCell(source.ColumnCaptionIndex + i, sourceColumnIndex + 1).ToLower(),i);

                if (key == string.Empty)
                    emptykeyscounter++;

                if (emptykeyscounter > 10)
                    break;
            }
            return column;
        }

        private bool IsKeyChecked()
        {
            var keys = from el in columnParsers
                       where el.IsKey == true
                       select el;

            if (keys.Count() > 0)
                return true;
            else
                return false;
        }

        private bool IsSecondKeyChecked()
        {
            var keys = from el in columnParsers
                       where el.LookupMatch == true
                       select el;

            if (keys.Count() > 0)
                return true;
            else
                return false;
        }

        private void OverwriteDestinationBySource()
        {
            int max = columnParsers.Count;
            int zerocounter = 0;
            int rowstartdest = _migration.Destination.ColumnCaptionIndex + 1;
            int rowstartsource = _migration.Source.ColumnCaptionIndex + 1;
            foreach (var col in columnParsers)
            {
                /*
                if ((col.SourceColumnIndex == 0 && zerocounter == 0) || col.SourceColumnIndex > 0)
                {
                    _migration.Destination.WriteRange(rowstartdest + 1, col.DestinationColumnIndex + 1,
                        _migration.Source.ReadColumn(rowstartsource, _migration.Source.Rows, col.SourceColumnIndex + 1));
                }*/

                if ((col.SourceColumnIndex == 0 && zerocounter == 0) || col.SourceColumnIndex > 0)
                {
                    _migration.Source.CopyColumn2AnotherWorkbook(rowstartsource, col.SourceColumnIndex + 1,
                        rowstartdest, col.DestinationColumnIndex + 1,
                        _migration.Destination.ActiveWorksheet);
                }

                if (col.SourceColumnIndex == 0)
                    zerocounter++;
            }
            _migration.Destination.Save();
            _migration.Source.Save();
        }
    }
}
