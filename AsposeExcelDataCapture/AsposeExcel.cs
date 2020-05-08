using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;

namespace AsposeExcelDataCapture
{

    public class AsposeExcel
    {
        private Workbook _workbook;
        private Worksheet _worksheet;
        private string _fullPath;
        private string _fileName;
        private string _extension;
        private string _fileDirectory;
        private Dictionary<string,int> _columnCaptionsWithIndex;
        private int _columnsCount;
        private int _rowsCount;
        public int ColumnCaptionRow = 2;

        public int RowsCount
        {
            get { return _rowsCount; }
            set { _rowsCount = value; }
        }

        public int ColumnsCount
        {
            get { return _columnsCount; }
            set { _columnsCount = value; }
        }

        public Dictionary<string,int> ColumnCaptionsWithIndexDictionary
        {
            get { return _columnCaptionsWithIndex; }
            set { _columnCaptionsWithIndex = value; }
        }


        public List<string> ColumnCaptionList
        {
            get { return ColumnCaptionsWithIndexDictionary.Keys.ToList(); }
        }


        public string FileDirectory
        {
            get { return _fileDirectory; }
            set { _fileDirectory = value; }
        }


        public string Extension
        {
            get { return _extension; }
            set { _extension = value; }
        }

        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        public string FullPath
        {
            get { return _fullPath; }
            set {
                _fullPath = value;
                FileName = Path.GetFileNameWithoutExtension(_fullPath);
                Extension = Path.GetExtension(_fullPath);
                FileDirectory = Path.GetDirectoryName(_fullPath);
            }
        }


        public Worksheet CurrentWorksheet
        {
            get { return _worksheet; }
            set {
                _worksheet = value;
                CountRowAndColumns();
                LoadColumnsCaptionWithIndexesDictionary();
            }
        }

        public Workbook CurrentWorkbook
        {
            get { return _workbook; }
            set { _workbook = value; }
        }

        public string CurrentWorksheetName
        {
            get
            {
                return CurrentWorksheet.Name.ToString();
            }
        }

        public bool ChangeWorksheet(string name)
        {
            var ws = from el in CurrentWorkbook.Worksheets
                     where el.Name == name
                     select el;

            if(ws.Count()==1)
            {
                CurrentWorksheet = CurrentWorkbook.Worksheets[name];
                return true;
            }
            else
            {
                return false;
            }
        }

        public AsposeExcel(string fullpath)
        {
            if (File.Exists(fullpath))
            {
                FullPath = fullpath;
                CurrentWorkbook = new Workbook(fullpath);
                CurrentWorksheet = CurrentWorkbook.Worksheets[0];
            }
            else
            {
                throw new FileNotFoundException();
            }
        }

        private void LoadColumnsCaptionWithIndexesDictionary()
        {
            _columnCaptionsWithIndex = new Dictionary<string, int>();
            for (int i = 0; i < ColumnsCount; i++)
            {
                string str = CurrentWorksheet.Cells.GetCell(ColumnCaptionRow, i).Value.ToString();
                if (!_columnCaptionsWithIndex.ContainsKey(str))
                    _columnCaptionsWithIndex.Add(str, i);
                else
                {
                    _columnCaptionsWithIndex.Add(str + "E", i);
                }
            }
        }

        private void CountRowAndColumns()
        {
            ColumnsCount = CurrentWorksheet.Cells.MaxDataColumn + 1;
            RowsCount = CurrentWorksheet.Cells.MaxDataRow + 1;
        }

        public string ReadCell(int row, int column)
        {
            if(row >= 0 && column >= 0)
            {
                try
                {
                    return CurrentWorksheet.Cells.GetCell(row,column).Value.ToString();
                }
                catch (NullReferenceException)
                {
                    return string.Empty;
                }
                catch
                {
                    return Convert.ToString(CurrentWorksheet.Cells.GetCell(row, column).Value);
                }
            }
            else
            {
                throw new IndexOutOfRangeException();
            }
        }

        public void WriteCell(int row, int column, string str)
        {
            if (row >= 0 && column >= 0)
            {
                CurrentWorksheet.Cells[row, column].PutValue(str);
                CurrentWorksheet.AutoFitColumn(column);
            }
            else
            {
                throw new IndexOutOfRangeException();
            }
        }
        public void WriteCell(int row, int column, int i)
        {
            if (row >= 0 && column >= 0)
            {
                CurrentWorksheet.Cells[row, column].PutValue(i);
                CurrentWorksheet.AutoFitColumn(column);
            }
            else
            {
                throw new IndexOutOfRangeException();
            }
        }

        public void WriteCell(int row, int column, bool i)
        {
            if (row >= 0 && column >= 0)
            {
                CurrentWorksheet.Cells[row, column].PutValue(i);
                CurrentWorksheet.AutoFitColumn(column);
            }
            else
            {
                throw new IndexOutOfRangeException();
            }
        }

        public void CopyRange2Worksheet(int sourceColumnIndex, int destinationColumnIndex, AsposeExcel destinationExcel)
        {
            destinationExcel.CurrentWorksheet.Cells.CopyColumn(CurrentWorksheet.Cells,
                sourceColumnIndex, destinationColumnIndex);

            destinationExcel.CurrentWorksheet.AutoFitColumn(destinationColumnIndex);
        }

        public void CopyRange2Worksheet(int row,int sourceColumnIndex, int destinationColumnIndex, AsposeExcel destinationExcel)
        {
            List<Cell> list = new List<Cell>();

            for (int i = 0; i < row; i++)
                list.Add(destinationExcel.CurrentWorksheet.Cells.GetCell(i, destinationColumnIndex));

            destinationExcel.CurrentWorksheet.Cells.CopyColumn(CurrentWorksheet.Cells,
                sourceColumnIndex, destinationColumnIndex);

            for (int i = 0; i < row; i++)
            {
                try
                {
                    destinationExcel.CurrentWorksheet.Cells.GetCell(i, destinationColumnIndex).PutValue(list[i].Value);
                }
                catch(NullReferenceException)
                {
                    //do nothing
                }
            }
                

            destinationExcel.CurrentWorksheet.AutoFitColumn(destinationColumnIndex);
        }

        public void Save()
        {
            CurrentWorkbook.Save(FullPath);
        }

        public List<string> ListOfWorksheets()
        {
            List<string> list = new List<string>();
            foreach(Worksheet worksheet in CurrentWorkbook.Worksheets)
            {
                list.Add(worksheet.Name.ToString());
            }
            return list;
        }
    }


}
