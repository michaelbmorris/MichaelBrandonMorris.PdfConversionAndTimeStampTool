//-----------------------------------------------------------------------------------------------------------
// <copyright file="IPdfConversionAndTimeStampToolView.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;

namespace PdfConversionAndTimeStampTool
{
    internal interface IPdfConversionAndTimeStampToolView :
        IProgress<ProgressReport>
    {
        event Action FilesSelected;

        event Action TaskRequested;

        Field Field { get; }
        List<string> FileNames { get; set; }
        List<string> OpenFileNames { get; }
        Script Script { get; }

        void ClearFiles();

        void ClearProgress();

        void ShowMessage(string message);

        void ToggleEnabled();
    }
}