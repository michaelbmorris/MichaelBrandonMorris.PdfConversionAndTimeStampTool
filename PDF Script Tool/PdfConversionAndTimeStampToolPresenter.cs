//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfConversionAndTimeStampToolPresenter.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace PdfConversionAndTimeStampTool
{
    internal class PdfConversionAndTimeStampToolPresenter
    {
        private readonly IPdfConversionAndTimeStampToolView view;

        internal PdfConversionAndTimeStampToolPresenter(
            IPdfConversionAndTimeStampToolView view)
        {
            this.view = view;
            this.view.FilesSelected += OnFilesSelected;
            this.view.TaskRequested += OnTaskRequested;
        }

        private void OnFilesSelected()
        {
            var fileNames = new List<string>();
            if (view.FileNames != null)
            {
                foreach (var fileName in view.FileNames)
                {
                    if (fileName.FileNameIsContainedIn(view.OpenFileNames))
                    {
                        view.ShowMessage("File \"" +
                            Path.GetFileNameWithoutExtension(fileName) +
                            "\" is already open.");
                    }
                    else
                    {
                        fileNames.Add(FileProcessor.PrepareFile(fileName));
                    }
                }
                view.FileNames = fileNames;
            }
        }

        private void OnTaskRequested()
        {
            view.ToggleEnabled();
            FileProcessor.ProcessFiles(
                view.OpenFileNames,
                view,
                view.Field,
                view.Script);
            view.ClearFiles();
            view.ToggleEnabled();
            view.ClearProgress();
            view.ShowMessage("Files saved to " + FileProcessor.OutputPath);
            Process.Start(FileProcessor.OutputPath);
        }
    }
}