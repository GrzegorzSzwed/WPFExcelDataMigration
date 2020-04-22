using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AsposeExcelDataCapture;

namespace WPFUIExcelDataCapture.Models
{
    public class MigrationContainer
    {
        private AsposeExcel _source;
        private AsposeExcel _destination;
        private List<string> _sourceColumns;
        private List<string> _destinationColumns;
        private bool _sourceFromTemplate = false;
        private List<ColumnParser> _columnParsers;

        public List<ColumnParser> ColumnParsers
        {
            get { return _columnParsers; }
            set { _columnParsers = value; }
        }


        public bool TemplateAttached
        {
            get { return _sourceFromTemplate; }
            set { _sourceFromTemplate = value; }
        }


        public MigrationContainer()
        {
            _sourceColumns = new List<string>();
            _destinationColumns = new List<string>();
            _columnParsers = new List<ColumnParser>();
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



        public AsposeExcel Destination
        {
            get { return _destination; }
            set {
                _destination = value;
                DestinationColumns = _destination.ColumnCaptionList;
            }
        }

        public AsposeExcel Source
        {
            get { return _source; }
            set {
                _source = value;
                SourceColumns = _source.ColumnCaptionList;
            }
        }
    }
}
