using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataCapture;

namespace WPFUIExcelDataCapture.Models
{
    public class MigrationContainer
    {
        private NavisionExcel _source;
        private NavisionExcel _destination;
        private List<string> _sourceColumns;
        private List<string> _destinationColumns;

        public MigrationContainer()
        {

        }


        public List<string> DestinationColumns
        {
            get { return _destinationColumns; }
            set { _destinationColumns = value; }
        }


        public List<string> SourceColumns
        {
            get { return _sourceColumns; }
            set { _sourceColumns = value; }
        }



        public NavisionExcel Destination
        {
            get { return _destination; }
            set {
                _destination = value;
                DestinationColumns = _destination.ColumnCaptions;
            }
        }

        public NavisionExcel Source
        {
            get { return _source; }
            set {
                _source = value;
                SourceColumns = _source.ColumnCaptions;
            }
        }



    }
}
