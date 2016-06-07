//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfConversionAndTimeStampToolView.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using static PdfConversionAndTimeStampTool.Properties.Resources;

namespace PdfConversionAndTimeStampTool
{
    internal partial class PdfConversionAndTimeStampToolView : Form,
        IPdfConversionAndTimeStampToolView
    {
        internal PdfConversionAndTimeStampToolView()
        {
            InitializeComponent();
            BindComponent();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = OpenFileDialogFilter;
            openFileDialog.Title = OpenFileDialogTitle;
        }

        public event Action FilesSelected;

        public event Action TaskRequested;

        public List<string> CheckedFileNames
        {
            get
            {
                return fileView.CheckedItems.OfType<string>().ToList();
            }
        }

        public Field Field { get; set; }

        public List<string> FileNames
        {
            get
            {
                return openFileDialog.FileNames.ToList();
            }

            set
            {
                foreach (string fileName in value)
                {
                    fileView.Items.Add(fileName, isChecked: true);
                }
            }
        }

        public List<string> OpenFileNames
        {
            get
            {
                return fileView.Items.OfType<string>().ToList();
            }
        }

        public Script Script { get; set; }

        public void ClearFiles()
        {
            fileView.Items.Clear();
        }

        public void ClearProgress()
        {
            progressBar.Value = 0;
        }

        public void Report(ProgressReport progressReport)
        {
            progressBar.Value = progressReport.Percent;
        }

        public void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }

        public void ToggleEnabled()
        {
            Enabled = !Enabled;
        }

        private void BindComponent()
        {
            selectFilesButton.Click += OnSelectFilesButtonClick;
            convertOnlyButton.Click += OnTaskButtonClick;
            convertAndTimeStampDefaultDayButton.Click += OnTaskButtonClick;
            convertAndTimeStampDefaultMonthButton.Click += OnTaskButtonClick;
        }

        private void OnSelectFilesButtonClick(object sender, EventArgs e)
        {
            var DialogResult = openFileDialog.ShowDialog();
            if (DialogResult == DialogResult.OK)
            {
                FilesSelected?.Invoke();
            }
        }

        private void OnTaskButtonClick(object sender, EventArgs e)
        {
            if (CheckedFileNames.Any())
            {
                if (sender == convertOnlyButton)
                {
                    Field = null;
                    Script = null;
                }
                else if (sender == convertAndTimeStampDefaultDayButton)
                {
                    Field = Field.DefaultTimeStampField;
                    Script = Script.TimeStampOnPrintDefaultDayScript;
                }
                else if (sender == convertAndTimeStampDefaultMonthButton)
                {
                    Field = Field.DefaultTimeStampField;
                    Script = Script.TimeStampOnPrintDefaultMonthScript;
                }
                else if (sender == addCustomFieldButton)
                {
                    // TODO
                    // Field = new Field();
                }
                else if (sender == addCustomScriptButton)
                {
                    // TODO
                    // Script = new Script();
                }
                TaskRequested?.Invoke();
            }
            else
            {
                ShowMessage("Please select at least one file for processing.");
            }
        }
    }
}