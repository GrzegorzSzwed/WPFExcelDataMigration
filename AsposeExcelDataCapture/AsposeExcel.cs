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
        private List<string> _columnCaptions;
        private Dictionary<string,int> _columnCaptionsWithIndex;
        private int _columnsCount;
        private int _rowsCount;
        public int ColumnCaptionRow = 3;

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
            get { return _columnCaptions; }
            set { _columnCaptions = value; }
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
            set { _worksheet = value; }
        }

        public Workbook CurrentWorkbook
        {
            get { return _workbook; }
            set { _workbook = value; }
        }

        public AsposeExcel(string fullpath)
        {
            if (File.Exists(fullpath))
            {
                FullPath = fullpath;
                CurrentWorkbook = new Workbook(fullpath);
                CurrentWorksheet = CurrentWorkbook.Worksheets[0];
                CountRowAndColumns();
                LoadColumnsCaptionWithIndexesDictionary();
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
                _columnCaptionsWithIndex.Add(CurrentWorksheet.Cells.GetCell(ColumnCaptionRow, i).Value.ToString(),i);
            }
        }

        private void CountRowAndColumns()
        {
            ColumnsCount = CurrentWorksheet.Cells.MaxDataColumn;
            RowsCount = CurrentWorksheet.Cells.MaxDataRow;
        }

        public string ReadCell(int row, int column)
        {
            if(row > 0 && column > 0)
            {
                try
                {
                    return CurrentWorksheet.Cells.GetCell(row,column).Value.ToString();
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
            if (row > 0 && column > 0 && str != string.Empty)
            {
                CurrentWorksheet.Cells.GetCell(row, column).PutValue(str);
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
    }


}
