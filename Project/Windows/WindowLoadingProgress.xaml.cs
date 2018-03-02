using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using SystemWeaver.WordImport.Common;

namespace SystemWeaver.WordImport.Windows
{
    /// <summary>
    /// Interaction logic for WindowLoadingProgress.xaml
    /// </summary>
    public partial class WindowLoadingProgress : Window, ThreadedWindowWrapper.IThreadedWindow, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public Window Window { get { return this; } }
        public string Status { get { return _status; } set { _status = value; OnPropertyChanged("Status"); } }

        private string _status;
        private bool _canClose;

        public WindowLoadingProgress()
        {
            InitializeComponent();
            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            DataContext = this;
            _canClose = false;
        }

        public void SetStatus(string status)
        {
            Status = status;
        }

        public void SetTitle(string title)
        {
            Title = title;
        }
        public void SetProgress(string value)
        {
            progress.Value = Int32.Parse(value);
        }

        public void SetCanClose(bool canClose)
        {
            _canClose = canClose;
        }

        private void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(name));
        }
    }
}
