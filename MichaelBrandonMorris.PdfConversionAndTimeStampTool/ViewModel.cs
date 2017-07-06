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
using Progress =
    System.IProgress<(int Current, int Total, string CurrentFileName)>;
using static System.IO.File;
using static System.Windows.MessageBox;
using static System.Windows.MessageBoxResult;
using static System.Windows.MessageBoxButton;
using static System.IO.Path;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    /// <summary>
    ///     Class ViewModel.
    /// </summary>
    /// <seealso cref="INotifyPropertyChanged" />
    /// TODO Edit XML Comment Template for ViewModel
    internal class ViewModel : INotifyPropertyChanged
    {
        /// <summary>
        ///     The help file
        /// </summary>
        /// TODO Edit XML Comment Template for HelpFile
        private static readonly string HelpFile =
            Combine(Combine("Resources", "Help"), "Help.chm");

        /// <summary>
        ///     The replace files message box text
        /// </summary>
        /// TODO Edit XML Comment Template for ReplaceFilesMessageBoxText
        private const string ReplaceFilesMessageBoxText =
            "Would you like to replace the already selected files?";

        /// <summary>
        ///     The select files filter
        /// </summary>
        /// TODO Edit XML Comment Template for SelectFilesFilter
        private const string SelectFilesFilter =
            "Office Files & PDFs|*.doc;*.docx;*.pdf;*.ppt;*.pptx;*.xls;*.xlsx";

        /// <summary>
        ///     The about window
        /// </summary>
        /// TODO Edit XML Comment Template for _aboutWindow
        private AboutWindow _aboutWindow;

        /// <summary>
        ///     The custom page numbers
        /// </summary>
        /// TODO Edit XML Comment Template for _customPageNumbers
        private string _customPageNumbers;

        /// <summary>
        ///     The field bottom y
        /// </summary>
        /// TODO Edit XML Comment Template for _fieldBottomY
        private int _fieldBottomY;

        /// <summary>
        ///     The field left x
        /// </summary>
        /// TODO Edit XML Comment Template for _fieldLeftX
        private int _fieldLeftX;

        /// <summary>
        ///     The field right x
        /// </summary>
        /// TODO Edit XML Comment Template for _fieldRightX
        private int _fieldRightX;

        /// <summary>
        ///     The field title
        /// </summary>
        /// TODO Edit XML Comment Template for _fieldTitle
        private string _fieldTitle;

        /// <summary>
        ///     The field top y
        /// </summary>
        /// TODO Edit XML Comment Template for _fieldTopY
        private int _fieldTopY;

        /// <summary>
        ///     The is busy
        /// </summary>
        /// TODO Edit XML Comment Template for _isBusy
        private bool _isBusy;

        /// <summary>
        ///     The message
        /// </summary>
        /// TODO Edit XML Comment Template for _message
        private string _message;

        /// <summary>
        ///     The message is visible
        /// </summary>
        /// TODO Edit XML Comment Template for _messageIsVisible
        private bool _messageIsVisible;

        /// <summary>
        ///     The message z index
        /// </summary>
        /// TODO Edit XML Comment Template for _messageZIndex
        private int _messageZIndex;

        /// <summary>
        ///     The progress percent
        /// </summary>
        /// TODO Edit XML Comment Template for _progressPercent
        private int _progressPercent;

        /// <summary>
        ///     The progress text
        /// </summary>
        /// TODO Edit XML Comment Template for _progressText
        private string _progressText;

        /// <summary>
        ///     The selected field pages
        /// </summary>
        /// TODO Edit XML Comment Template for _selectedFieldPages
        private FieldPages _selectedFieldPages;

        /// <summary>
        ///     The selected script
        /// </summary>
        /// TODO Edit XML Comment Template for _selectedScript
        private string _selectedScript;

        /// <summary>
        ///     The selected timing
        /// </summary>
        /// TODO Edit XML Comment Template for _selectedTiming
        private ScriptTiming _selectedTiming;

        /// <summary>
        ///     The should show custom page numbers
        /// </summary>
        /// TODO Edit XML Comment Template for _shouldShowCustomPageNumbers
        private bool _shouldShowCustomPageNumbers;

        /// <summary>
        ///     The user guide
        /// </summary>
        /// TODO Edit XML Comment Template for _userGuide
        private Process _userGuide;

        /// <summary>
        ///     Gets the cancel.
        /// </summary>
        /// <value>The cancel.</value>
        /// TODO Edit XML Comment Template for Cancel
        public ICommand Cancel => new RelayCommand(ExecuteCancel, CanCancel);

        /// <summary>
        ///     Gets the convert.
        /// </summary>
        /// <value>The convert.</value>
        /// TODO Edit XML Comment Template for Convert
        public ICommand Convert => new RelayCommand(
            ExecuteConvert,
            CanExecuteAction);

        /// <summary>
        ///     Gets the custom action.
        /// </summary>
        /// <value>The custom action.</value>
        /// TODO Edit XML Comment Template for CustomAction
        public ICommand CustomAction => new RelayCommand(
            ExecuteCustomAction,
            CanExecuteAction);

        /// <summary>
        ///     Gets the field pages.
        /// </summary>
        /// <value>The field pages.</value>
        /// TODO Edit XML Comment Template for FieldPages
        public IList<FieldPages> FieldPages
        {
            get;
        } = GetFieldPages();

        /// <summary>
        ///     Gets the open about window.
        /// </summary>
        /// <value>The open about window.</value>
        /// TODO Edit XML Comment Template for OpenAboutWindow
        public ICommand OpenAboutWindow => new RelayCommand(
            ExecuteOpenAboutWindow);

        /// <summary>
        ///     Gets the open user guide.
        /// </summary>
        /// <value>The open user guide.</value>
        /// TODO Edit XML Comment Template for OpenUserGuide
        public ICommand OpenUserGuide => new RelayCommand(ExecuteOpenUserGuide);

        /// <summary>
        ///     Gets the script timings.
        /// </summary>
        /// <value>The script timings.</value>
        /// TODO Edit XML Comment Template for ScriptTimings
        public IList<ScriptTiming> ScriptTimings
        {
            get;
        } = GetScriptTimings();

        /// <summary>
        ///     Gets the select files.
        /// </summary>
        /// <value>The select files.</value>
        /// TODO Edit XML Comment Template for SelectFiles
        public ICommand SelectFiles => new RelayCommand(ExecuteSelectFiles);

        /// <summary>
        ///     Gets the select script.
        /// </summary>
        /// <value>The select script.</value>
        /// TODO Edit XML Comment Template for SelectScript
        public ICommand SelectScript => new RelayCommand(ExecuteSelectScript);

        /// <summary>
        ///     Gets the time stamp day.
        /// </summary>
        /// <value>The time stamp day.</value>
        /// TODO Edit XML Comment Template for TimeStampDay
        public ICommand TimeStampDay => new RelayCommand(
            ExecuteTimeStampDay,
            CanExecuteAction);

        /// <summary>
        ///     Gets the time stamp month.
        /// </summary>
        /// <value>The time stamp month.</value>
        /// TODO Edit XML Comment Template for TimeStampMonth
        public ICommand TimeStampMonth => new RelayCommand(
            ExecuteTimeStampMonth,
            CanExecuteAction);

        /// <summary>
        ///     Gets the version.
        /// </summary>
        /// <value>The version.</value>
        /// TODO Edit XML Comment Template for Version
        public string Version
        {
            get
            {
                string version;

                try
                {
                    version = ApplicationDeployment.CurrentDeployment
                        .CurrentVersion.ToString();
                }
                catch (InvalidDeploymentException)
                {
                    version = "Dev";
                }

                return version;
            }
        }

        /// <summary>
        ///     Gets or sets the custom page numbers.
        /// </summary>
        /// <value>The custom page numbers.</value>
        /// TODO Edit XML Comment Template for CustomPageNumbers
        public string CustomPageNumbers
        {
            get => _customPageNumbers;
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

        /// <summary>
        ///     Gets or sets the field bottom y.
        /// </summary>
        /// <value>The field bottom y.</value>
        /// TODO Edit XML Comment Template for FieldBottomY
        public int FieldBottomY
        {
            get => _fieldBottomY;
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

        /// <summary>
        ///     Gets or sets the field left x.
        /// </summary>
        /// <value>The field left x.</value>
        /// TODO Edit XML Comment Template for FieldLeftX
        public int FieldLeftX
        {
            get => _fieldLeftX;
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

        /// <summary>
        ///     Gets or sets the field right x.
        /// </summary>
        /// <value>The field right x.</value>
        /// TODO Edit XML Comment Template for FieldRightX
        public int FieldRightX
        {
            get => _fieldRightX;
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

        /// <summary>
        ///     Gets or sets the field title.
        /// </summary>
        /// <value>The field title.</value>
        /// TODO Edit XML Comment Template for FieldTitle
        public string FieldTitle
        {
            get => _fieldTitle;
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

        /// <summary>
        ///     Gets or sets the field top y.
        /// </summary>
        /// <value>The field top y.</value>
        /// TODO Edit XML Comment Template for FieldTopY
        public int FieldTopY
        {
            get => _fieldTopY;
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

        /// <summary>
        ///     Gets or sets a value indicating whether this instance
        ///     is busy.
        /// </summary>
        /// <value>
        ///     <c>true</c> if this instance is busy; otherwise,
        ///     <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for IsBusy
        public bool IsBusy
        {
            get => _isBusy;
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

        /// <summary>
        ///     Gets or sets the message.
        /// </summary>
        /// <value>The message.</value>
        /// TODO Edit XML Comment Template for Message
        public string Message
        {
            get => _message;
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

        /// <summary>
        ///     Gets or sets a value indicating whether [message is
        ///     visible].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [message is visible]; otherwise,
        ///     <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for MessageIsVisible
        public bool MessageIsVisible
        {
            get => _messageIsVisible;
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

        /// <summary>
        ///     Gets or sets the index of the message z.
        /// </summary>
        /// <value>The index of the message z.</value>
        /// TODO Edit XML Comment Template for MessageZIndex
        public int MessageZIndex
        {
            get => _messageZIndex;
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

        /// <summary>
        ///     Gets or sets the progress percent.
        /// </summary>
        /// <value>The progress percent.</value>
        /// TODO Edit XML Comment Template for ProgressPercent
        public int ProgressPercent
        {
            get => _progressPercent;
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

        /// <summary>
        ///     Gets or sets the progress text.
        /// </summary>
        /// <value>The progress text.</value>
        /// TODO Edit XML Comment Template for ProgressText
        public string ProgressText
        {
            get => _progressText;
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

        /// <summary>
        ///     Gets or sets the selected field pages.
        /// </summary>
        /// <value>The selected field pages.</value>
        /// TODO Edit XML Comment Template for SelectedFieldPages
        public FieldPages SelectedFieldPages
        {
            get => _selectedFieldPages;
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

        /// <summary>
        ///     Gets or sets the selected file names.
        /// </summary>
        /// <value>The selected file names.</value>
        /// TODO Edit XML Comment Template for SelectedFileNames
        public ObservableCollection<CheckedListItem<string>> SelectedFileNames
        {
            get;
            set;
        } = new ObservableCollection<CheckedListItem<string>>();

        /// <summary>
        ///     Gets or sets the selected script.
        /// </summary>
        /// <value>The selected script.</value>
        /// TODO Edit XML Comment Template for SelectedScript
        public string SelectedScript
        {
            get => _selectedScript;
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

        /// <summary>
        ///     Gets or sets the selected timing.
        /// </summary>
        /// <value>The selected timing.</value>
        /// TODO Edit XML Comment Template for SelectedTiming
        public ScriptTiming SelectedTiming
        {
            get => _selectedTiming;
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

        /// <summary>
        ///     Gets or sets a value indicating whether [should show
        ///     custom page numbers].
        /// </summary>
        /// <value>
        ///     <c>true</c> if [should show custom page numbers];
        ///     otherwise, <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for ShouldShowCustomPageNumbers
        public bool ShouldShowCustomPageNumbers
        {
            get => _shouldShowCustomPageNumbers;
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

        /// <summary>
        ///     Gets the about window.
        /// </summary>
        /// <value>The about window.</value>
        /// TODO Edit XML Comment Template for AboutWindow
        private AboutWindow AboutWindow
        {
            get
            {
                if (_aboutWindow == null
                    || !_aboutWindow.IsVisible)
                {
                    _aboutWindow = new AboutWindow();
                }

                return _aboutWindow;
            }
        }

        /// <summary>
        ///     Gets the progress.
        /// </summary>
        /// <value>The progress.</value>
        /// TODO Edit XML Comment Template for Progress
        private Progress Progress => new
            Progress<(int Current, int Total, string CurrentFileName)>(
                HandleProgressReport);

        /// <summary>
        ///     Gets the user guide.
        /// </summary>
        /// <value>The user guide.</value>
        /// TODO Edit XML Comment Template for UserGuide
        private Process UserGuide
        {
            get
            {
                if (_userGuide != null
                    && !_userGuide.HasExited)
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

        /// <summary>
        ///     Gets or sets the file processor.
        /// </summary>
        /// <value>The file processor.</value>
        /// TODO Edit XML Comment Template for FileProcessor
        private FileProcessor FileProcessor
        {
            get;
            set;
        }

        /// <summary>
        ///     Occurs when a property value changes.
        /// </summary>
        /// TODO Edit XML Comment Template for PropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        ///     Determines whether this instance can cancel.
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance can cancel;
        ///     otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanCancel
        private bool CanCancel()
        {
            return IsBusy;
        }

        /// <summary>
        ///     Determines whether this instance [can execute action].
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this instance [can execute action];
        ///     otherwise, <c>false</c>.
        /// </returns>
        /// TODO Edit XML Comment Template for CanExecuteAction
        private bool CanExecuteAction()
        {
            return !SelectedFileNames.IsEmpty();
        }

        /// <summary>
        ///     Executes the cancel.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteCancel
        private void ExecuteCancel()
        {
            FileProcessor.Cancel();
        }

        /// <summary>
        ///     Executes the convert.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteConvert
        private void ExecuteConvert()
        {
            ExecuteTask();
        }

        /// <summary>
        ///     Executes the custom action.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteCustomAction
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

        /// <summary>
        ///     Executes the open about window.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteOpenAboutWindow
        private void ExecuteOpenAboutWindow()
        {
            AboutWindow.Show();
            AboutWindow.Activate();
        }

        /// <summary>
        ///     Executes the open user guide.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteOpenUserGuide
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

        /// <summary>
        ///     Executes the select files.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteSelectFiles
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

        /// <summary>
        ///     Executes the select script.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteSelectScript
        private void ExecuteSelectScript()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "JavaScript File (*.js)|*.js"
            };

            openFileDialog.ShowDialog();
            SelectedScript = openFileDialog.FileName;
        }

        /// <summary>
        ///     Executes the task.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <param name="script">The script.</param>
        /// TODO Edit XML Comment Template for ExecuteTask
        private async void ExecuteTask(Field field = null, Script script = null)
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

        /// <summary>
        ///     Executes the time stamp day.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteTimeStampDay
        private void ExecuteTimeStampDay()
        {
            ExecuteTask(TimeStampField, TimeStampOnPrintDay);
        }

        /// <summary>
        ///     Executes the time stamp month.
        /// </summary>
        /// TODO Edit XML Comment Template for ExecuteTimeStampMonth
        private void ExecuteTimeStampMonth()
        {
            ExecuteTask(TimeStampField, TimeStampOnPrintMonth);
        }

        /// <summary>
        ///     Gets the custom page numbers.
        /// </summary>
        /// <returns>IEnumerable&lt;System.Int32&gt;.</returns>
        /// TODO Edit XML Comment Template for GetCustomPageNumbers
        private IEnumerable<int> GetCustomPageNumbers()
        {
            if (SelectedFieldPages != Custom
                || CustomPageNumbers.IsNullOrWhiteSpace())
            {
                return null;
            }

            var customPageNumbers = new List<int>();
            var customPageNumberStrings = CustomPageNumbers.Split(',');

            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (var customPageNumber in customPageNumberStrings)
            {
                if (customPageNumber.TryParse(out int result))
                {
                    customPageNumbers.Add(result);
                }
            }

            return customPageNumbers;
        }

        /// <summary>
        ///     Handles the progress report.
        /// </summary>
        /// <param name="progressReport">The progress report.</param>
        /// TODO Edit XML Comment Template for HandleProgressReport
        private void HandleProgressReport(
            (int Curremt, int Total, string CurrentFileName) progressReport)
        {
            var current = progressReport.Curremt;
            var total = progressReport.Total;
            ProgressPercent = current * 100 / total;
            ProgressText = $"{current} / {total}";
        }

        /// <summary>
        ///     Hides the message.
        /// </summary>
        /// TODO Edit XML Comment Template for HideMessage
        private void HideMessage()
        {
            Message = string.Empty;
            MessageIsVisible = false;
            MessageZIndex = -1;
        }

        /// <summary>
        ///     Notifies the proeprty changed.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// TODO Edit XML Comment Template for NotifyProeprtyChanged
        private void NotifyProeprtyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        ///     Shows the message.
        /// </summary>
        /// <param name="message">The message.</param>
        /// TODO Edit XML Comment Template for ShowMessage
        private void ShowMessage(string message)
        {
            Message = message + "\n\nDouble-click to dismiss.";
            MessageIsVisible = true;
            MessageZIndex = 1;
        }
    }
}