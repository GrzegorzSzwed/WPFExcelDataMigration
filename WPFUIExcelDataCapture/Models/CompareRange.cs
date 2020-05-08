namespace WPFUIExcelDataCapture.Models
{
    public class CompareRange
    {
        private string _keyValue;
        private int _row;
        private int _keyColumn;
        private string _matchValue;
        private int _matchColumn;

        public int MatchColumn
        {
            get { return _matchColumn; }
            set { _matchColumn = value; }
        }

        public string MatchValue
        {
            get { return _matchValue; }
            set { _matchValue = value; }
        }

        public int KeyColumn
        {
            get { return _keyColumn; }
            set { _keyColumn = value; }
        }

        public int Row
        {
            get { return _row; }
            set { _row = value; }
        }

        public string KeyValue
        {
            get { return _keyValue; }
            set { _keyValue = value; }
        }

        public CompareRange(string key, int keyrow, int column, string match, int matchcolumn)
        {
            KeyValue = key;
            Row = keyrow;
            KeyColumn = column;
            MatchValue = match;
            MatchColumn = matchcolumn;
        }
    }
}
