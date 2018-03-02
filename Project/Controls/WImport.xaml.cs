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
using System.Windows.Navigation;
using System.Windows.Shapes;
using SystemWeaver.WordImport.ViewModel;

namespace SystemWeaver.WordImport.Controls
{
    /// <summary>
    /// Interaction logic for WImport.xaml
    /// </summary>
    public partial class WImport : UserControl
    {
        private WordImportModel _dataContext;
        public WImport(Window parent)
        {
            InitializeComponent();

            _dataContext = (WordImportModel)DataContext;
            _dataContext.CurrentWindow = this;
            _dataContext.Parent = parent;
        }
    }
}