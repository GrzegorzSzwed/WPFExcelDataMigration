using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFUIExcelDataCapture.Models
{
    public class CompareRangeList : CompareRange
    {
        private Dictionary<string, int> _keyValues;
        private Dictionary<string, int> _matchValues;
        private decimal _acceptablePercentSimilarity;
        private List<CompareRange> _similarRows;

        public List<CompareRange> SimilarRows
        {
            get { return _similarRows; }
            set { _similarRows = value; }
        }


        public CompareRangeList(string key, int keyrow, int column, string match, int matchcolumn)
            : base(key, keyrow, column, match, matchcolumn)
        {
            _keyValues = new Dictionary<string, int>();
            _matchValues = new Dictionary<string, int>();
            _similarRows = new List<CompareRange>();
        }

        public decimal AcceptablePercentSimilarity
        {
            get { return _acceptablePercentSimilarity; }
            set {
                if (value >= 0 && value <= 100)
                    _acceptablePercentSimilarity = value;
                else
                    throw new IndexOutOfRangeException();
            }
        }

        public Dictionary<string, int> MatchValues
        {
            get { return _matchValues; }
            set { _matchValues = value; }
        }

        public Dictionary<string, int> KeyValues
        {
            get { return _keyValues; }
            set { _keyValues = value; }
        }


    }
}
