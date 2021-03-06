﻿using System;
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
using TextMatchCalculation;
using AsposeExcelDataCapture;
using Microsoft.Win32;

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
        private AsposeExcel _templateExcel;
        private List<string> MergeActions = new List<string>();

        public ComparisonView(MigrationContainer migration)
        {
            InitializeComponent();
            _migration = migration;
            MergeActions.Add("Overwrite");
            MergeActions.Add("Merge By Key");
            MergeActions.Add("Merge With Text Match");
            MergeActions.Add("Merge And Check Relation");
            MergeActions.Add("Add 2 Worksheet");

            cmbMergeOption.ItemsSource = MergeActions;

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
                    if (cmbMergeOption.Text == "Overwrite")
                    {
                        OverwriteDestinationBySource();
                    }

                    if (cmbMergeOption.Text == "Merge By Key")
                    {
                        MergeDestinationWithSource();
                    }

                    if(cmbMergeOption.Text == "Merge And Check Relation")
                    {
                        MergeAndCheckRelations();
                    }

                    if(cmbMergeOption.Text == "Merge With Text Match")
                    {
                        MergeCheckTextMatch();
                    }

                    if(cmbMergeOption.Text == "Add 2 Worksheet")
                    {
                        Add2Worksheet();
                    }
                }
                else
                {
                    var msg = new MessageView("Please load some excels to merge them");
                    msg.Show();
                }
            }
        }

        #region MergeCheckTextMatch
        private void MergeCheckTextMatch()
        {
            if (IsSecondKeyChecked() && IsKeyChecked())
            {
                var keycolumn = from el in columnParsers
                                where el.IsKey == true
                                select el;

                var matchcolumn = from el in columnParsers
                                  where el.LookupMatch == true
                                  select el;

                if (keycolumn.Count() != 1 || matchcolumn.Count() != 1)
                {
                    var msg = new MessageView("You cannot specify more than one first key!");
                    msg.Show();
                }
                else
                {
                    var keycol = keycolumn.FirstOrDefault();
                    var matchcol = matchcolumn.FirstOrDefault();

                    Dictionary<string, int> sourceKeys = new Dictionary<string, int>();
                    Dictionary<string, int> destinationKeys = new Dictionary<string, int>();

                    sourceKeys = LoadColumn(_migration.Source, keycol.SourceColumnIndex);
                    destinationKeys = LoadColumn(_migration.Destination, keycol.DestinationColumnIndex);
                    SortedDictionary<string, int> sortedSourceKeys = new SortedDictionary<string, int>(sourceKeys);

                    Dictionary<string, int> sourceNames = new Dictionary<string, int>();
                    Dictionary<string, int> destinationNames = new Dictionary<string, int>();

                    sourceNames = LoadColumn(_migration.Source, matchcol.SourceColumnIndex);
                    destinationNames = LoadColumn(_migration.Destination, matchcol.DestinationColumnIndex);

                    int ifNotFoundKeyRow = LastEmptyRow(_migration.Destination);
                    bool permission = true;
                    int row = 0;

                    int columncaptionsindex = _migration.Destination.ColumnCaptionRow;
                    foreach (var cl in sortedSourceKeys)
                    {
                        if (cl.Key != string.Empty)
                        {
                            if (destinationKeys.ContainsKey(cl.Key))
                            {
                                row = destinationKeys[cl.Key] + columncaptionsindex;
                                int rowsource = sourceKeys[cl.Key] + columncaptionsindex;
                                decimal percent = Levenstein.Percent(
                                    _migration.Source.ReadCell(rowsource, matchcol.SourceColumnIndex),
                                    _migration.Destination.ReadCell(row, matchcol.DestinationColumnIndex)
                                    );

                                if(percent < 80)
                                {
                                    var elements = from el in destinationNames
                                                   where Levenstein.NettoDistance(cl.Key, el.Key) <= 3 && el.Key != string.Empty
                                                   select el;

                                    if (elements.Count() > 0)
                                    {
                                        KeyValuePair<string, int> result = new KeyValuePair<string, int>();
                                        var comp = new SimilarNames(elements, result);
                                        comp.ShowDialog();
                                    }
                                }
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
        #endregion

        #region MergeByKeyWithRelation

        private void MergeAndCheckRelations()
        {
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

                    int columncaptionsindex = _migration.Destination.ColumnCaptionRow;
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
            int columncaptionsindex = _migration.Destination.ColumnCaptionRow;
            foreach (var columnparser in columnParsers)
            {
                if ((columnparser.SourceColumnIndex == 0 && permission == true) || columnparser.SourceColumnIndex > 0)
                {
                    if (columnparser.LookupMatch)
                    {
                        if (sourceKeys.Keys.Contains(_migration.Source.ReadCell(cl.Value, columnparser.SourceColumnIndex)))
                        {
                            _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex + 1,
                                _migration.Source.ReadCell(cl.Value + columncaptionsindex, columnparser.SourceColumnIndex));
                        }
                        else
                        {
                            _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex, "");
                        }
                    }
                    else
                    {
                        _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex,
                            _migration.Source.ReadCell(cl.Value + columncaptionsindex, columnparser.SourceColumnIndex));
                    }
                }

                if (columnparser.SourceColumnIndex == 0)
                    permission = false; ;
            }
        }

        #endregion

        #region MergeByKey
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

                    int columncaptionsindex = _migration.Destination.ColumnCaptionRow;
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

        private void FromSource2DestinationByRow(int destinationRow, KeyValuePair<string, int> cl, Dictionary<string, int> sourceKeys, bool permission)
        {
            int columncaptionsindex = _migration.Destination.ColumnCaptionRow;
            foreach (var columnparser in columnParsers)
            {
                if ((columnparser.SourceColumnIndex == 0 && permission == true) || columnparser.SourceColumnIndex > 0)
                {
                    _migration.Destination.WriteCell(destinationRow, columnparser.DestinationColumnIndex,
                         _migration.Source.ReadCell(cl.Value + columncaptionsindex, columnparser.SourceColumnIndex));
                }

                if (columnparser.SourceColumnIndex == 0)
                    permission = false; ;
            }
        }

        #endregion

        #region Overwrite

        private void OverwriteDestinationBySource()
        {
            int max = columnParsers.Count;
            int zerocounter = 0;
            int rowstartdest = _migration.Destination.ColumnCaptionRow + 1;
            int rowstartsource = _migration.Source.ColumnCaptionRow + 1;
            foreach (var col in columnParsers)
            {
                if ((col.SourceColumnIndex == 0 && zerocounter == 0) || col.SourceColumnIndex > 0)
                {
                    _migration.Source.CopyRange2Worksheet(col.SourceColumnIndex,
                        col.DestinationColumnIndex,
                        _migration.Destination);
                }

                if (col.SourceColumnIndex == 0)
                    zerocounter++;
            }
            _migration.Destination.Save();
            _migration.Source.Save();
        }

        #endregion

        #region Add2Worksheet

        private void Add2Worksheet()
        {
            int max = columnParsers.Count;
            int zerocounter = 0;
            int ifNotFoundKeyRow = LastEmptyRow(_migration.Destination);
            int row = _migration.Source.ColumnCaptionRow + 1;

            foreach (var col in columnParsers)
            {
                if ((col.SourceColumnIndex == 0 && zerocounter == 0) || col.SourceColumnIndex > 0)
                {
                    var SouceValue = _migration.Source.ReadCell(row++, col.SourceColumnIndex);
                    _migration.Destination.WriteCell(ifNotFoundKeyRow++, col.DestinationColumnIndex,
                            SouceValue);
                }

                if (col.SourceColumnIndex == 0)
                    zerocounter++;
            }
            _migration.Destination.Save();
            _migration.Source.Save();
        }

        #endregion

        #region Methods
        private void LoadComparisonView()
        {
            if (!_migration.TemplateAttached)
            {
                if (_migration.Destination != null)
                {
                    var captionsWithIndex = _migration.Destination.ColumnCaptionsWithIndexDictionary;
                    foreach (var col in captionsWithIndex)
                    {
                        ColumnParser parser = new ColumnParser();
                        parser.DestinationColumnName = col.Key;
                        parser.DestinationColumnIndex = col.Value;
                        parser.DestinationColumnCaptionList = _migration.Destination.ColumnCaptionList;
                        parser.SourceColumnCaptionList = _migration.Source.ColumnCaptionList;
                        FillSourceIfFoundDestination(parser, col.Key);

                        columnParsers.Add(parser);
                    }
                    itemsListColumnParsers.ItemsSource = columnParsers;
                }
            }
            else
            {
                if (_migration.Destination != null)
                {
                    foreach (var parser in _migration.ColumnParsers)
                    {
                        parser.DestinationColumnCaptionList = _migration.Destination.ColumnCaptionList;
                        parser.SourceColumnCaptionList = _migration.Source.ColumnCaptionList;

                        columnParsers.Add(parser);
                    }
                    itemsListColumnParsers.ItemsSource = columnParsers;
                }
            }

        }

        public void FillSourceIfFoundDestination(ColumnParser parser, string caption)
        {
            if (_migration.Source.ColumnCaptionsWithIndexDictionary.ContainsKey(caption))
            {
                parser.SourceColumnName = caption;
                parser.SourceColumnIndex = _migration.Source.ColumnCaptionsWithIndexDictionary[caption];
            }
            else
            {
                if (_migration.Source.ColumnCaptionsWithIndexDictionary.ContainsKey(caption))
                {
                    parser.SourceColumnName = caption;
                    parser.SourceColumnIndex = _migration.Source.ColumnCaptionsWithIndexDictionary[caption.ToLower()];
                }
                else
                {
                    string highest = string.Empty;
                    int memory = 100;
                    foreach (var par in _migration.Source.ColumnCaptionList)
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
                        parser.SourceColumnIndex = _migration.Source.ColumnCaptionsWithIndexDictionary[highest];
                    }
                    else
                    {
                        var s = _migration.Source.ColumnCaptionsWithIndexDictionary.First();
                        parser.SourceColumnName = s.Key;
                        parser.SourceColumnIndex = s.Value;
                    }
                }
            }
        }
        private int LastEmptyRow(AsposeExcel destination)
        {
            return destination.RowsCount + 1;
        }

        private Dictionary<string, int> LoadColumn(AsposeExcel source, int sourceColumnIndex)
        {
            Dictionary<string, int> column = new Dictionary<string, int>();
            int emptykeyscounter = 0;
            for (int i = 0; i < source.RowsCount; i++)
            {
                var key = source.ReadCell(source.ColumnCaptionRow + 1 + i, sourceColumnIndex).ToLower();
                if (!column.ContainsKey(key))
                    column.Add(source.ReadCell(source.ColumnCaptionRow + 1 + i, sourceColumnIndex).ToLower(), i);

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
        #endregion

        #region SaveTemplate
        private void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (_migration.Destination != null)
                openFileDialog.InitialDirectory = _migration.Destination.FileDirectory;
            else
                openFileDialog.InitialDirectory = "C:\\";

            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsm|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true)
            {
                filePath = openFileDialog.FileName;
                Template = new AsposeExcel(filePath);
                Write2Template();
                Template.Save();
                var msg = new MessageView("Template has been saved");
                msg.Show();
            }
            else
            {
                var msg = new MessageView("Please choose some template file");
                msg.Show();
            }
        }

        private void Write2Template()
        {
            if (columnParsers.Count() > 0)
            {
                //initial cells
                Template.WriteCell(0, 0, "Source:");
                Template.WriteCell(0, 1, _migration.Source.FullPath);
                Template.WriteCell(0, 2, _migration.Source.CurrentWorksheetName);
                Template.WriteCell(1, 0, "Destination:");
                Template.WriteCell(1, 1, _migration.Destination.FullPath);
                Template.WriteCell(1, 2, _migration.Destination.CurrentWorksheetName);

                //columns
                Template.WriteCell(2, 0, "SourceColumn");
                Template.WriteCell(2, 1, "SourceColumnIndex");
                Template.WriteCell(2, 2, "DestinationColumn");
                Template.WriteCell(2, 3, "DestinationColumnIndex");
                Template.WriteCell(2, 4, "isKey");
                Template.WriteCell(2, 5, "LookupMatch");

                int row = 3;
                foreach(var col in columnParsers)
                {
                    Template.WriteCell(row, 0, col.SourceColumnName);
                    Template.WriteCell(row, 1, col.SourceColumnIndex);
                    Template.WriteCell(row, 2, col.DestinationColumnName);
                    Template.WriteCell(row, 3, col.DestinationColumnIndex);
                    Template.WriteCell(row, 4, col.IsKey);
                    Template.WriteCell(row, 5, col.LookupMatch);
                    row++;
                }
            }
        }
#endregion

        #region Properties
        public AsposeExcel Template
        {
            get { return _templateExcel; }
            set { _templateExcel = value; }
        }
        #endregion

        private void ListViewItem_Selected(object sender, RoutedEventArgs e)
        {

        }
    }
}
