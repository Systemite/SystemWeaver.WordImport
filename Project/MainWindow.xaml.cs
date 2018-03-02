using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Systemite.SystemWeaver.Controls.UserControls.Data;
using SystemWeaver.WordImport.Controls;
using SystemWeaver.WordImport.ViewModel;

namespace SystemWeaver.WordImport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, ILoginValidator
    {
        //Wpf private IswItemViewHost _host;
        //private SWEventManager _eventManager;
        //Wpf private IswDialogs _dialogs;
        //private WordImportModel _dataContext;

        public MainWindow()
        {
            InitializeComponent();

            appControl.LoginInfoManager = new FileLoginInfoManager("user.config");
            appControl.LoginValidator = this;
            appControl.AddTab("WordImport Import", () => new WImport(this), true);
            //_dataContext = (ItemViewModel)DataContext;
            //_dataContext.CurrentWindow = this;
        }
        public bool ValidateBeforeLogin(string server, string port, string user, SecureString password)
        {
            //if (user != "system")
            //{
            //    MessageBox.Show("Must login as a system user.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //    return false;
            //}

            return true;
        }

    }
}
