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
using System.Windows.Shapes;

namespace WPFUIExcelDataCapture.Views
{
    /// <summary>
    /// Interaction logic for SimilarNames.xaml
    /// </summary>
    public partial class SimilarNames : Window
    {
        List<Element2Compare> elements = new List<Element2Compare>();
        KeyValuePair<string, int> _result = new KeyValuePair<string, int>();

        public SimilarNames(IEnumerable<KeyValuePair<string,int>> list,KeyValuePair<string,int> result)
        {
            InitializeComponent();

            if (list.Count() > 0)
            {
                foreach(var el in list)
                {
                    elements.Add(new Element2Compare { ColumnName = el.Key, Index = el.Value , Chosen = false });
                }
                _result = result;
            }
        }
    }

    public class Element2Compare
    {
        private string _columnName;
        private int _index;
        private bool _chosen;

        public bool Chosen
        {
            get { return _chosen; }
            set { _chosen = value; }
        }


        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }


        public string ColumnName
        {
            get { return _columnName; }
            set { _columnName = value; }
        }

    }
}
