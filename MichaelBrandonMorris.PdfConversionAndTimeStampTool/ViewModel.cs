using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using GalaSoft.MvvmLight.CommandWpf;
using MichaelBrandonMorris.Extensions.CollectionExtensions;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;
using Microsoft.Win32;
using static MichaelBrandonMorris.PdfConversionAndTimeStampTool.FieldPages;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    internal class ViewModel : INotifyPropertyChanged
    {
        private const string ReplaceFilesMessageBoxText =
            "Would you like to replace the already selected files?";

        private string _customPageNumbers;
        private int _fieldBottomY;
        private int _fieldLeftX;
        private int _fieldRightX;
        private string _fieldTitle;
        private int _fieldTopY;

        private int _progressPercent;
        private FieldPages _selectedFieldPages;
        private string _selectedScript;
        private ScriptTiming _selectedTiming;
        private bool _shouldShowCustomPageNumbers;

        public ICommand Convert => new RelayCommand(
            ExecuteConvert, CanExecuteAction);

        public ICommand CustomAction => new RelayCommand(
            ExecuteCustomAction, CanExecuteAction);

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

        public IList<FieldPages> FieldPages
        {
            get;
        } = Field.GetFieldPages();

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

        public IList<ScriptTiming> ScriptTimings
        {
            get;
        } = Script.GetScriptTimings();

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

        public ICommand SelectFiles => new RelayCommand(ExecuteSelectFiles);

        public ICommand SelectScript => new RelayCommand(ExecuteSelectScript);

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

        public ICommand TimeStampDay => new RelayCommand(
            ExecuteTimeStampDay, CanExecuteAction);

        public ICommand TimeStampMonth => new RelayCommand(
            ExecuteTimeStampMonth, CanExecuteAction);

        private IProgress<Tuple<int, int, string>> Progress =>
            new Progress<Tuple<int, int, string>>(HandleProgressReport);

        public event PropertyChangedEventHandler PropertyChanged;

        private bool CanExecuteAction()
        {
            return !SelectedFileNames.IsEmpty();
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

            var scriptText = File.ReadAllText(SelectedScript);
            var script = new Script(scriptText, SelectedTiming);

            ExecuteTask(field, script);
        }

        private void ExecuteSelectFiles()
        {
            if (!SelectedFileNames.IsEmpty())
            {
                var messageBoxResult = MessageBox.Show(
                    ReplaceFilesMessageBoxText,
                    "Replace Files?",
                    MessageBoxButton.YesNo);

                if (messageBoxResult == MessageBoxResult.Yes)
                {
                    SelectedFileNames.Clear();
                }
            }

            var openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Documents (*.doc;*.docx;*.pdf)|*.doc;*.docx;*.pdf"
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
            Field field = null, Script script = null)
        {
            var fileNames = from x in SelectedFileNames select x.Item;
            var fileProcessor = new FileProcessor(
                fileNames.ToList(), Progress, field, script);
            await fileProcessor.Execute();
        }

        private void ExecuteTimeStampDay()
        {
            ExecuteTask(Field.TimeStampField, Script.TimeStampOnPrintDay);
        }

        private void ExecuteTimeStampMonth()
        {
            ExecuteTask(Field.TimeStampField, Script.TimeStampOnPrintMonth);
        }

        private IEnumerable<int> GetCustomPageNumbers()
        {
            if (SelectedFieldPages != Custom ||
                CustomPageNumbers.IsNullOrWhiteSpace())
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
            Tuple<int, int, string> progressReport)
        {
            ProgressPercent = progressReport.Item1 * 100 / progressReport.Item2;
        }

        private void NotifyProeprtyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this, new PropertyChangedEventArgs(propertyName));
        }
    }
}