using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Deployment.Application;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using GalaSoft.MvvmLight.CommandWpf;
using MichaelBrandonMorris.Extensions.CollectionExtensions;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;
using Microsoft.Win32;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.FieldPages;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.Field;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.Script;
using ProgressReport = System.Tuple<int, int, string>;
using Progress = System.IProgress<System.Tuple<int, int, string>>;
using static System.IO.File;
using static System.Windows.MessageBox;
using static System.Windows.MessageBoxResult;
using static System.Windows.MessageBoxButton;
using static System.IO.Path;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    internal class ViewModel : INotifyPropertyChanged
    {
        private static readonly string HelpFile =
            Combine(Combine("Resources", "Help"), "Help.chm");

        private const string ReplaceFilesMessageBoxText =
            "Would you like to replace the already selected files?";

        private const string SelectFilesFilter =
            "Office Files & PDFs|*.doc;*.docx;*.pdf;*.ppt;*.pptx;*.xls;*.xlsx";

        private AboutWindow _aboutWindow;
        private string _customPageNumbers;
        private int _fieldBottomY;
        private int _fieldLeftX;
        private int _fieldRightX;
        private string _fieldTitle;
        private int _fieldTopY;
        private bool _isBusy;
        private string _message;
        private bool _messageIsVisible;
        private int _messageZIndex;
        private int _progressPercent;
        private string _progressText;
        private FieldPages _selectedFieldPages;
        private string _selectedScript;
        private ScriptTiming _selectedTiming;
        private bool _shouldShowCustomPageNumbers;
        private Process _userGuide;

        public ICommand Cancel
        {
            get
            {
                return new RelayCommand(ExecuteCancel, CanCancel);
            }
        }

        public ICommand Convert
        {
            get
            {
                return new RelayCommand(ExecuteConvert, CanExecuteAction);
            }
        }

        public ICommand CustomAction
        {
            get
            {
                return new RelayCommand(ExecuteCustomAction, CanExecuteAction);
            }
        }

        public IList<FieldPages> FieldPages
        {
            get;
        } = GetFieldPages();

        public ICommand OpenAboutWindow
        {
            get
            {
                return new RelayCommand(ExecuteOpenAboutWindow);
            }
        }

        public ICommand OpenUserGuide
        {
            get
            {
                return new RelayCommand(ExecuteOpenUserGuide);
            }
        }

        public IList<ScriptTiming> ScriptTimings
        {
            get;
        } = GetScriptTimings();

        public ICommand SelectFiles
        {
            get
            {
                return new RelayCommand(ExecuteSelectFiles);
            }
        }

        public ICommand SelectScript
        {
            get
            {
                return new RelayCommand(ExecuteSelectScript);
            }
        }

        public ICommand TimeStampDay
        {
            get
            {
                return new RelayCommand(ExecuteTimeStampDay, CanExecuteAction);
            }
        }

        public ICommand TimeStampMonth
        {
            get
            {
                return new RelayCommand(
                    ExecuteTimeStampMonth,
                    CanExecuteAction);
            }
        }

        public string Version
        {
            get
            {
                string version;

                try
                {
                    version =
                        ApplicationDeployment.CurrentDeployment.CurrentVersion
                            .ToString();
                }
                catch (InvalidDeploymentException)
                {
                    version = "Dev";
                }

                return version;
            }
        }

        public string CustomPageNumbers
        {
            get
            {
                return _customPageNumbers;
            }
            set
            {
                if (_customPageNumbers == value)
                {
                    return;
                }

                _customPageNumbers = value;
                NotifyProeprtyChanged();
            }
        }

        public int FieldBottomY
        {
            get
            {
                return _fieldBottomY;
            }
            set
            {
                if (_fieldBottomY == value)
                {
                    return;
                }

                _fieldBottomY = value;
                NotifyProeprtyChanged();
            }
        }

        public int FieldLeftX
        {
            get
            {
                return _fieldLeftX;
            }
            set
            {
                if (_fieldLeftX == value)
                {
                    return;
                }

                _fieldLeftX = value;
                NotifyProeprtyChanged();
            }
        }

        public int FieldRightX
        {
            get
            {
                return _fieldRightX;
            }
            set
            {
                if (_fieldRightX == value)
                {
                    return;
                }

                _fieldRightX = value;
                NotifyProeprtyChanged();
            }
        }

        public string FieldTitle
        {
            get
            {
                return _fieldTitle;
            }
            set
            {
                if (_fieldTitle == value)
                {
                    return;
                }

                _fieldTitle = value;
                NotifyProeprtyChanged();
            }
        }

        public int FieldTopY
        {
            get
            {
                return _fieldTopY;
            }
            set
            {
                if (_fieldTopY == value)
                {
                    return;
                }

                _fieldTopY = value;
                NotifyProeprtyChanged();
            }
        }

        public bool IsBusy
        {
            get
            {
                return _isBusy;
            }
            set
            {
                if (_isBusy == value)
                {
                    return;
                }

                _isBusy = value;
                NotifyProeprtyChanged();
            }
        }

        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                if (_message == value)
                {
                    return;
                }

                _message = value;
                NotifyProeprtyChanged();
            }
        }

        public bool MessageIsVisible
        {
            get
            {
                return _messageIsVisible;
            }
            set
            {
                if (_messageIsVisible == value)
                {
                    return;
                }

                _messageIsVisible = value;
                NotifyProeprtyChanged();
            }
        }

        public int MessageZIndex
        {
            get
            {
                return _messageZIndex;
            }
            set
            {
                if (_messageZIndex == value)
                {
                    return;
                }

                _messageZIndex = value;
                NotifyProeprtyChanged();
            }
        }

        public int ProgressPercent
        {
            get
            {
                return _progressPercent;
            }
            set
            {
                if (_progressPercent == value)
                {
                    return;
                }

                _progressPercent = value;
                NotifyProeprtyChanged();
            }
        }

        public string ProgressText
        {
            get
            {
                return _progressText;
            }
            set
            {
                if (_progressText == value)
                {
                    return;
                }

                _progressText = value;
                NotifyProeprtyChanged();
            }
        }

        public FieldPages SelectedFieldPages
        {
            get
            {
                return _selectedFieldPages;
            }
            set
            {
                if (_selectedFieldPages == value)
                {
                    return;
                }

                _selectedFieldPages = value;
                NotifyProeprtyChanged();
                ShouldShowCustomPageNumbers = _selectedFieldPages == Custom;
            }
        }

        public ObservableCollection<CheckedListItem<string>> SelectedFileNames
        {
            get;
            set;
        } = new ObservableCollection<CheckedListItem<string>>();

        public string SelectedScript
        {
            get
            {
                return _selectedScript;
            }
            set
            {
                if (_selectedScript == value)
                {
                    return;
                }

                _selectedScript = value;
                NotifyProeprtyChanged();
            }
        }

        public ScriptTiming SelectedTiming
        {
            get
            {
                return _selectedTiming;
            }
            set
            {
                if (_selectedTiming == value)
                {
                    return;
                }

                _selectedTiming = value;
                NotifyProeprtyChanged();
            }
        }

        public bool ShouldShowCustomPageNumbers
        {
            get
            {
                return _shouldShowCustomPageNumbers;
            }
            set
            {
                if (_shouldShowCustomPageNumbers == value)
                {
                    return;
                }

                _shouldShowCustomPageNumbers = value;
                NotifyProeprtyChanged();
            }
        }

        private AboutWindow AboutWindow
        {
            get
            {
                if (_aboutWindow == null || !_aboutWindow.IsVisible)
                {
                    _aboutWindow = new AboutWindow();
                }

                return _aboutWindow;
            }
        }

        private Progress Progress
        {
            get
            {
                return new Progress<ProgressReport>(HandleProgressReport);
            }
        }

        private Process UserGuide
        {
            get
            {
                if (_userGuide != null && !_userGuide.HasExited)
                {
                    _userGuide.Kill();
                }
                _userGuide = new Process
                {
                    StartInfo =
                    {
                        FileName = HelpFile
                    }
                };

                return _userGuide;
            }
        }

        private FileProcessor FileProcessor
        {
            get;
            set;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private bool CanCancel()
        {
            return IsBusy;
        }

        private bool CanExecuteAction()
        {
            return !SelectedFileNames.IsEmpty();
        }

        private void ExecuteCancel()
        {
            FileProcessor.Cancel();
        }

        private void ExecuteConvert()
        {
            ExecuteTask();
        }

        private void ExecuteCustomAction()
        {
            var field = new Field(
                FieldTitle,
                FieldLeftX,
                FieldTopY,
                FieldRightX,
                FieldBottomY,
                SelectedFieldPages,
                GetCustomPageNumbers());

            var scriptText = ReadAllText(SelectedScript);
            var script = new Script(scriptText, SelectedTiming);

            ExecuteTask(field, script);
        }

        private void ExecuteOpenAboutWindow()
        {
            AboutWindow.Show();
            AboutWindow.Activate();
        }

        private void ExecuteOpenUserGuide()
        {
            try
            {
                UserGuide.Start();
            }
            catch (Exception)
            {
                ShowMessage("User guide could not be opened.");
            }
        }

        private void ExecuteSelectFiles()
        {
            if (!SelectedFileNames.IsEmpty())
            {
                var messageBoxResult = Show(
                    ReplaceFilesMessageBoxText,
                    "Replace Files?",
                    YesNo);

                if (messageBoxResult == Yes)
                {
                    SelectedFileNames.Clear();
                }
            }

            var openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = SelectFilesFilter
            };

            openFileDialog.ShowDialog();

            foreach (var item in openFileDialog.FileNames)
            {
                SelectedFileNames.Add(new CheckedListItem<string>(item, true));
            }
        }

        private void ExecuteSelectScript()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "JavaScript File (*.js)|*.js"
            };

            openFileDialog.ShowDialog();
            SelectedScript = openFileDialog.FileName;
        }

        private async void ExecuteTask(
            Field field = null,
            Script script = null)
        {
            IsBusy = true;
            HideMessage();
            var fileNames = from x in SelectedFileNames select x.Item;

            FileProcessor = new FileProcessor(
                fileNames.ToList(),
                Progress,
                field,
                script);

            try
            {
                await FileProcessor.Execute();
            }
            catch (OperationCanceledException)
            {
                ShowMessage("The operation was cancelled.");
            }

            IsBusy = false;
        }

        private void ExecuteTimeStampDay()
        {
            ExecuteTask(TimeStampField, TimeStampOnPrintDay);
        }

        private void ExecuteTimeStampMonth()
        {
            ExecuteTask(TimeStampField, TimeStampOnPrintMonth);
        }

        private IEnumerable<int> GetCustomPageNumbers()
        {
            if (SelectedFieldPages != Custom
                || CustomPageNumbers.IsNullOrWhiteSpace())
            {
                return null;
            }

            var customPageNumbers = new List<int>();
            var customPageNumberStrings = CustomPageNumbers.Split(',');

            foreach (var customPageNumber in customPageNumberStrings)
            {
                int result;

                if (customPageNumber.TryParse(out result))
                {
                    customPageNumbers.Add(result);
                }
            }

            return customPageNumbers;
        }

        private void HandleProgressReport(
            ProgressReport progressReport)
        {
            var current = progressReport.Item1;
            var total = progressReport.Item2;
            ProgressPercent = current * 100 / total;
            ProgressText = $"{current} / {total}";
        }

        private void HideMessage()
        {
            Message = string.Empty;
            MessageIsVisible = false;
            MessageZIndex = -1;
        }

        private void NotifyProeprtyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }

        private void ShowMessage(string message)
        {
            Message = message + "\n\nDouble-click to dismiss.";
            MessageIsVisible = true;
            MessageZIndex = 1;
        }
    }
}