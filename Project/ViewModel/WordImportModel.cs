using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Windows.Input;
using System.IO;
using System.Windows;
using System.Collections.ObjectModel;
using SystemWeaver.Common;
using SystemWeaver.Images;
using SystemWeaverAPI;
using SystemWeaver.WordImport.Common;
using SystemWeaver.WordImport.Windows;
using SystemWeaver.WordImport.Controls;
using Systemite.SystemWeaver.Controls.WTreeViews;
using Systemite.SystemWeaver.Controls.ViewModels;
using System.Windows.Controls;

namespace SystemWeaver.WordImport.ViewModel
{
    public class WordImportModel : ViewModelBase
    {
        private static readonly SWImages SwImages = new SWImages();

        public WordImportModel()
        {
            Paragraphs = new List<SwParagraph>();
            XidText = "";
        }

        public Window Parent { get; set; }
        public WImport CurrentWindow { get; set; }

        private RelayCommand _selectWordFileCommand;
        public RelayCommand SelectWordFileCommand
        {
            get
            {
                return GetCommand(ref _selectWordFileCommand, o => SelectWordFile());
            }
        }

        private RelayCommand _importWordDocumentCommand;
        public RelayCommand ImportWordDocumentCommand
        {
            get
            {
                return GetCommand(ref _importWordDocumentCommand, o => ImportWordDocument());
            }
        }

        private RelayCommand _moveDownCommand;
        public RelayCommand MoveDownCommand
        {
            get
            {
                return GetCommand(ref _moveDownCommand, o => MoveDown());
            }
        }

        private string _fileNameText;
        public string FileNameText
        {
            get { return _fileNameText; }
            set
            {
                _fileNameText = value;
                OnPropertyChanged("FileNameText");
            }
        }

        private string _xidText;
        public string XidText
        {
            get { return _xidText; }
            set
            {
                _xidText = value;
                OnPropertyChanged("XidText");
            }
        }

        private string _errorText = "";
        public string ErrorText
        {
            get { return _errorText; }
            set
            {
                _errorText = value;
                if (InformationText.Length > 0)
                    InformationText = "";
                OnPropertyChanged("ErrorText");
            }
        }

        private bool _warningsVisible = false;
        public bool WarningsVisible
        {
            get { return _warningsVisible; }
            set
            {
                _warningsVisible = value;
                OnPropertyChanged("WarningsVisible");
            }
        }

        private string _importWarnings = "";
        public string ImportWarnings
        {
            get { return _importWarnings; }
            set
            {
                _importWarnings = value;
                OnPropertyChanged("ImportWarnings");
            }
        }

        private string _informationText = "";
        public string InformationText
        {
            get { return _informationText; }
            set
            {
                _informationText = value;
                if (ErrorText.Length > 0)
                    ErrorText = "";
                OnPropertyChanged("InformationText");
            }
        }

        private ObservableCollection<SwStyle> _swStyles;
        public ObservableCollection<SwStyle> SwStyles
        {
            get { return _swStyles; }
            set
            {
                _swStyles = value;
                OnPropertyChanged("SwStyles");
            }
        }

        private ObservableCollection<SwStyle> _descriptionStyles;
        public ObservableCollection<SwStyle> DescriptionStyles
        {
            get { return _descriptionStyles; }
            set
            {
                _descriptionStyles = value;
                OnPropertyChanged("DescriptionStyles");
            }
        }

        public List<SwParagraph> Paragraphs { get; set; }

        private IswItem _currentItem = null;
        public IswItem CurrentItem
        {
            set
            {
                _currentItem = value;
            }
            get
            {
                return _currentItem;
            }
        }

        private void MoveDown()
        {
            if (SwStyles.Count == 0)
                return;
            var last = SwStyles.Last();
            SwStyles.Remove(last);
            DescriptionStyles.Insert(0, last);
        }

        private void ImportWordDocument()
        {
            if (CurrentItem == null)
                return;

            if (CurrentItem.GetAllParts().Count > 0)
                throw new Exception("Item not empty");

            ErrorText = "";

            if (Paragraphs == null || Paragraphs.Count == 0)
            {
                ErrorText = "No word-file selected!";
                return;
            }

            try
            {
                InformationText = "Word import started!";
                ValidateParagraphs(Paragraphs, DescriptionStyles.ToList());

                using (var wrapper = new ThreadedWindowWrapper())
                {
                    wrapper.LaunchThreadedWindow<WindowLoadingProgress>(true);
                    wrapper.SetTitle("Importing word document");

                    Import imp = new Import(_currentItem, Paragraphs, SwStyles.ToList(), DescriptionStyles.ToList(), SWConnection.Instance.Broker.ServerId, wrapper);
                }
                InformationText = "Word import completed!";
            }
            catch (Exception)
            {

                throw;
            }
        }
   
        private void ValidateParagraphs(List<SwParagraph> paragraphs, List<SwStyle> descriptionList)
        {
            long totalLength = 0;
            long nrParagraphs = 0;
            StringBuilder warnings = new StringBuilder();
            WarningsVisible = false;
            foreach (var p in paragraphs)
            {
                if (!descriptionList.Select(x => x.Name).ToList().Contains(p.Style.Name))
                {
                    if(p.RtfData != null)
                        totalLength += p.RtfData.Length;
                    nrParagraphs++;
                    long avgLength = totalLength / nrParagraphs;
                    //Table, image etc. could be included in the header. Table or image will be missing in that case.
                    if (p.RtfData != null && p.RtfData.Length > avgLength * 5)
                    {
                        warnings.AppendLine("Paragraph with potential error: " + p.Text);
                        ImportWarnings = warnings.ToString();
                        WarningsVisible = true;
                    }
                }
            }
        }

        private void SelectWordFile()
        {
            OpenItem();
            if (_currentItem == null)
            {
                MessageBox.Show("Paste item id to the \"Item id\" textbox first");
                return;
            }

            Nullable<bool> result;
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Word files (*.docx)|*.docx|Word files (*.doc)|*.doc";
            result = openFileDialog.ShowDialog();
            if (result == true)
            {
                FileNameText = openFileDialog.FileName;
                SystemWeaver.WordImport.Common.ReadWord rw;
                using (var wrapper = new ThreadedWindowWrapper())
                {
                    wrapper.LaunchThreadedWindow<WindowLoadingProgress>(true);
                    wrapper.SetTitle("Loading word document");
                    wrapper.SetStatus("Loading word document, don't use clipboard now");

                    bool failedWithLoadDLL;
                    rw = new SystemWeaver.WordImport.Common.ReadWord(FileNameText, _currentItem, wrapper, out failedWithLoadDLL);
                    if (failedWithLoadDLL)
                    {
                        FileNameText = "";
                        return;
                    }
                }

                Paragraphs = rw.Paragraphs;
                DescriptionStyles = new ObservableCollection<SwStyle>();
                var stylesUsed = rw.StylesInUse;
                if ((from s in stylesUsed where s.Level == 10 select s).Count() > 0)
                {
                    var temp = (from s in stylesUsed where s.Level == 10 select s).ToList();
                    DescriptionStyles = new ObservableCollection<SwStyle>(temp);
                    stylesUsed = (from s in stylesUsed where s.Level != 10 select s).ToList();
                }
                SwStyles = new ObservableCollection<SwStyle>(stylesUsed);
                InformationText = "Word document loaded!";
            }
        }

        private void OpenItem()
        {
            if (XidText == null)
                return;
            long handle;
            string handleStr = XidText;
            int pos = XidText.LastIndexOf('/');
            if (pos >= 0)
                handleStr = XidText.Substring(pos + 1);
            if(SWHandleUtility.TryParseHandle(handleStr, out handle))
                CurrentItem = SystemWeaverAPI.SWConnection.Instance.Broker.GetItem(handle);
        }

        private RelayCommand GetCommand(
            ref RelayCommand relayCommand,
            Action<object> action,
            Predicate<object> canExecute = null)
        {
            if (relayCommand == null)
                relayCommand = new RelayCommand(action, canExecute);

            return relayCommand;
        }

    }
}
