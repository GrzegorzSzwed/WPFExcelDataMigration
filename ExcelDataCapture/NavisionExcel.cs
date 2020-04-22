using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDataCapture
{
    public class NavisionExcel
    {
        private readonly _Application _excel = new _Excel.Application();
        private Workbook _workbook;
        private Worksheet _currentWorksheet;
        private string _filename;
        private string _extension;
        private string _directory;
        private string _fullPath;
        private int _columns;
        private int _rows;
        private static readonly int defaultSheet = 1;
        private int _columnCaptionIndex = 3;
        private List<string> _columnCaptions;
        private Dictionary<string,int> _columnCaptionsWithIndex;

        public Dictionary<string,int> ColumnCaptionsWithIndex
        {
            get { return _columnCaptionsWithIndex; }
            set { _columnCaptionsWithIndex = value; }
        }


        public List<string> ColumnCaptions
        {
            get { return _columnCaptions; }
            set { _columnCaptions = value; }
        }


        #region ctor
        public NavisionExcel(string path)
        {
            if (File.Exists(path))
            {
                FullPath = path;
                _workbook = _excel.Workbooks.Open(FullPath);
                _currentWorksheet = _workbook.Worksheets[defaultSheet];

                Columns = _currentWorksheet.UsedRange.Columns.Count;
                Rows = _currentWorksheet.UsedRange.Rows.Count;
                LoadColumnCaptions();
                LoadColumnCaptionsWithIndex();
            }
        }
        public NavisionExcel(string path, int worksheet)
        {
            if (File.Exists(path))
            {
                FullPath = path;
                _workbook = _excel.Workbooks.Open(FullPath);
                _currentWorksheet = _workbook.Worksheets[worksheet];

                Columns = _currentWorksheet.UsedRange.Columns.Count;
                Rows = _currentWorksheet.UsedRange.Rows.Count;
                LoadColumnCaptions();
                LoadColumnCaptionsWithIndex();
            }
        }
        public NavisionExcel(string path, string worksheet)
        {
            if (File.Exists(path))
            {
                FullPath = path;
                _workbook = _excel.Workbooks.Open(FullPath);

                if (GetNoOfWorksheet(worksheet) > 0)
                    _currentWorksheet = _workbook.Worksheets[GetNoOfWorksheet(worksheet)];
                else
                    _currentWorksheet = _workbook.Worksheets[defaultSheet];

                Columns = _currentWorksheet.UsedRange.Columns.Count;
                Rows = _currentWorksheet.UsedRange.Rows.Count;
                LoadColumnCaptions();
                LoadColumnCaptionsWithIndex();
            }
        }
        #endregion

        public bool ChangeWorksheet(string name)
        {
            if (name != string.Empty)
            {
                int i = 1;
                foreach(Worksheet ws in _workbook.Worksheets)
                {
                    if (ws.Name == name || ws.Name.ToLower() == name.ToLower())
                    {
                        _currentWorksheet = _workbook.Worksheets[i];
                        Columns = _currentWorksheet.UsedRange.Columns.Count;
                        Rows = _currentWorksheet.UsedRange.Rows.Count;
                        LoadColumnCaptions();
                        LoadColumnCaptionsWithIndex();
                        return true;
                    }
                    i++;
                }
                return false;
            }
            else
                return false;
        }

        private void LoadColumnCaptions()
        {
            _columnCaptions = new List<string>();
            for (int i = 0; i < Columns; i++)
            {
                _columnCaptions.Add(ReadCell(ColumnCaptionIndex,i+1));
            }
        }

        private void LoadColumnCaptionsWithIndex()
        {
            _columnCaptionsWithIndex = new Dictionary<string, int>();
            for (int i = 0; i < Columns; i++)
            {
                _columnCaptionsWithIndex.Add(ReadCell(ColumnCaptionIndex, i+1), i);
            }
        }

        private int GetNoOfWorksheet(string worksheet)
        {
            if (worksheet != string.Empty)
            {
                int i = 1;
                foreach (Worksheet ws in _workbook.Worksheets)
                {
                    if (ws.Name == worksheet|| ws.Name.ToLower() == worksheet.ToLower())
                    {
                        return i;
                    }

                    i++;
                }
                return 0;
            }
            return -1;
        }

        public Worksheet ActiveWorksheet
        {
            get
            {
                return _currentWorksheet;
            }
        }

        public void CopyColumn2AnotherWorkbook(int sourcerow,int sourcecolumn,int destinationrow,
            int destinationcolumn,Worksheet destinationworksheet)
        {

            _currentWorksheet.Activate();


            destinationworksheet.Activate();
            destinationworksheet.Cells.PasteSpecial(destinationworksheet.Range[destinationworksheet.Cells[destinationrow, destinationcolumn]].EntireColumn.Select());
        }

        public Range ReadColumn(int rowstart, int rowstop, int column)
        {
            _workbook.Activate();
            _currentWorksheet.Range[Cell(rowstart, column)].EntireColumn.Select();
            var range = _currentWorksheet.Range[Cell(rowstart, column), Cell(rowstop, column)];
            return range;
        }

        public void WriteColumn(int rowstart, int column, Range range)
        {
            Range pasterange = _currentWorksheet.Range[Cell(rowstart, column), Cell(Rows, column)];
            range.Copy();
            range.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationAdd, pasterange);
        }

        public string ReadCell(int row, int column)
        {
            if (_currentWorksheet.Cells[row, column].Value2 != null)
            {
                try
                {
                    return _currentWorksheet.Cells[row, column].Value2;
                }
                catch
                {
                    return Convert.ToString(_currentWorksheet.Cells[row, column].Value2);
                }
            }
            else
                return string.Empty;
        }

        public void WriteCell(int row, int column, string str)
        {
            if (str != string.Empty)
            {
                _currentWorksheet.Cells[row, column].Value2 = str;
            }
        }

        public void AddWorksheet(string name)
        {
            if(GetNoOfWorksheet(name) == 0)
            {
                _workbook.Worksheets.Add();
                Worksheet ws = _workbook.Worksheets[1];
                ws.Name = name;
            }
        }

        public void DeleteWorksheet(string name)
        {
            if (GetNoOfWorksheet(name) > 0)
            {
                _workbook.Worksheets.Add();
                Worksheet ws = _workbook.Worksheets[1];
                ws.Name = name;
            }
        }

        public List<string> ListOfWorksheets()
        {
            List<string> ws = new List<string>();
            foreach (Worksheet works in _workbook.Worksheets)
            {
                ws.Add(works.Name.ToString());
            }
            return ws;
        }

        public Range Cell(int row, int column)
        {
            return _currentWorksheet.Cells[row, column];
        }

        public void CloseAndSave()
        {
            if (this != null)
            {
                _excel.Workbooks.Close();
                _excel.Quit();
                Marshal.ReleaseComObject(_excel);
            }
        }

        public void Save()
        {
            _excel.Workbooks.Close();
        }

        #region props
        public int ColumnCaptionIndex
        {
            get { return _columnCaptionIndex; }
            set { _columnCaptionIndex = value; }
        }

        public int Rows
        {
            get { return _rows; }
            set { _rows = value; }
        }


        public int Columns
        {
            get { return _columns; }
            set { _columns = value; }
        }


        public string FullPath
        {
            get { return _fullPath; }
            set {
                _fullPath = value;
                FileName = Path.GetFileNameWithoutExtension(value);
                Directory = Path.GetDirectoryName(value);
                Extension = Path.GetExtension(value);
            }
        }


        public string Directory
        {
            get { return _directory; }
            set { _directory = value; }
        }


        public string Extension
        {
            get { return _extension; }
            set { _extension = value; }
        }


        public string FileName
        {
            get { return _filename; }
            set { _filename = value; }
        }

        public string CurrentWorksheet
        {
            get
            {
                return _currentWorksheet.Name;
            }
        }
        #endregion
    }
}
