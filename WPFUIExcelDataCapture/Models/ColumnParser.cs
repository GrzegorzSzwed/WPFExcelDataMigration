using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFUIExcelDataCapture.Models
{
    public class ColumnParser
    {
        private string _sourceColumnName;
        private string _destinationColumnName;
        private int _sourceColumnIndex;
        private int _destinationColumnIndex;
        private bool _key;
        private bool _textMatch;
        private bool _avoidZero;

        public bool AvoidZero
        {
            get { return _avoidZero; }
            set { _avoidZero = value; }
        }


        private List<string> _sourceColumnList;
        private List<string> _destinationColumnList;

        public List<string> DestinationColumnCaptionList
        {
            get { return _destinationColumnList; }
            set { _destinationColumnList = value; }
        }


        public List<string> SourceColumnCaptionList
        {
            get { return _sourceColumnList; }
            set { _sourceColumnList = value; }
        }


        public bool IsKey
        {
            get { return _key; }
            set { _key = value; }
        }


        public int DestinationColumnIndex
        {
            get { return _destinationColumnIndex; }
            set { _destinationColumnIndex = value; }
        }


        public int SourceColumnIndex
        {
            get { return _sourceColumnIndex; }
            set { _sourceColumnIndex = value; }
        }


        public string DestinationColumnName
        {
            get { return _destinationColumnName; }
            set { _destinationColumnName = value; }
        }


        public string SourceColumnName
        {
            get { return _sourceColumnName; }
            set { _sourceColumnName = value; }
        }

        public bool LookupMatch
        {
            get { return _textMatch; }
            set { _textMatch = value; }
        }
    }
}
