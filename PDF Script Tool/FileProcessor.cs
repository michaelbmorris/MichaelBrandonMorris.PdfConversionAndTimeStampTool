//-----------------------------------------------------------------------------------------------------------
// <copyright file="FileProcessor.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

using iTextSharp.text.pdf;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using static PdfConversionAndTimeStampTool.Properties.Resources;
using static System.Environment;

namespace PdfConversionAndTimeStampTool
{
    internal static class FileProcessor
    {
        private const int EveryOtherPage = 2;

        private const int EveryPage = 1;

        private const int FirstPageNumber = 1;

        private const int SecondPageNumber = 2;

        internal static string OutputPath => Path.Combine(
            GetFolderPath(SpecialFolder.MyDocuments),
            RootFolderName);

        internal static string ProcessingPath => Path.Combine(
            GetFolderPath(SpecialFolder.ApplicationData),
            RootFolderName);

        internal static void ClearProcessing()
        {
            var processingDirectory = new DirectoryInfo(ProcessingPath);
            foreach (var file in processingDirectory.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception e)
                { 
                    Debug.WriteLine(e.Message + e.StackTrace);
                }
                
            }
        }

        internal static string CopyFileToProcessing(string filename)
        {
            var processingPath = GetProcessingPath(filename);
            File.Copy(filename, processingPath);
            return processingPath;
        }

        internal static string PrepareFile(string fileName)
        {
            return CopyFileToProcessing(fileName);
        }

        internal static List<string> PrepareFiles(List<string> fileNames)
        {
            for (int i = 0; i < fileNames.Count; i++)
            {
                fileNames[i] = CopyFileToProcessing(fileNames[i]);
            }
            return fileNames;
        }

        internal static void ProcessFiles(
            List<string> fileNames,
            IProgress<ProgressReport> progressReport,
            Field field = null,
            Script script = null)
        {
            for (var i = 0; i < fileNames.Count; i++)
            {
                var currentFile = fileNames[i];
                if (!IsPdf(currentFile))
                {
                    currentFile = ConvertToPdf(currentFile);
                }

                if (field != null || script != null)
                {
                    ProcessPdf(currentFile, field, script);
                }
                else
                {
                    MovePdfToOutput(currentFile);
                }
                progressReport.Report(new ProgressReport
                {
                    Total = fileNames.Count,
                    CurrentCount = i + 1
                });
            }
            ClearProcessing();
        }

        private static void AddFieldToPage(
            Field field,
            int pageNumber,
            PdfStamper pdfStamper,
            PdfFormField parentField)
        {
            var textField = new TextField(
                pdfStamper.Writer,
                new iTextSharp.text.Rectangle(
                    field.TopLeftX,
                    field.TopLeftY,
                    field.BottomRightX,
                    field.BottomRightY),
                null);
            var childField = textField.GetTextField();
            parentField.AddKid(childField);
            childField.PlaceInPage = pageNumber;
        }

        private static void AddFieldToPdf(
            Field field, PdfStamper pdfStamper, int numberOfPages)
        {
            var parentField = PdfFormField.CreateTextField(
                pdfStamper.Writer, false, false, 0);
            parentField.FieldName = field.Title;
            var pageNumber = field.Pages == Pages.Last ?
                numberOfPages : FirstPageNumber;
            if (field.Pages == Pages.First || field.Pages == Pages.Last)
            {
                AddFieldToPage(
                    field,
                    pageNumber,
                    pdfStamper,
                    parentField);
            }
            else
            {
                var increment = field.Pages == Pages.All ?
                    EveryPage : EveryOtherPage;
                if (field.Pages == Pages.Even)
                {
                    pageNumber += 1;
                }

                for (; pageNumber <= numberOfPages; pageNumber += increment)
                {
                    AddFieldToPage(
                        field,
                        pageNumber,
                        pdfStamper,
                        parentField);
                }
            }

            pdfStamper.AddAnnotation(parentField, FirstPageNumber);
        }

        private static void AddScriptToPdf(
            Script script, PdfStamper pdfStamper)
        {
            var pdfAction = PdfAction.JavaScript(
                script.ScriptText, pdfStamper.Writer);
            PdfName actionType = null;
            switch (script.ScriptEvent)
            {
                case ScriptEvent.DidPrint:
                    actionType = PdfWriter.DID_PRINT;
                    break;

                case ScriptEvent.DidSave:
                    actionType = PdfWriter.DID_SAVE;
                    break;

                case ScriptEvent.WillPrint:
                    actionType = PdfWriter.WILL_PRINT;
                    break;

                case ScriptEvent.WillSave:
                    actionType = PdfWriter.WILL_SAVE;
                    break;

                default:
                    throw new ArgumentOutOfRangeException();
            }

            pdfStamper.Writer.SetAdditionalAction(actionType, pdfAction);
        }

        private static string ConvertToPdf(string filename)
        {
            var outputFilename = Path.GetFileNameWithoutExtension(filename) +
                PdfFileExtension;
            var outputPath = Path.Combine(ProcessingPath, outputFilename);
            var wordApplication = new Application();
            wordApplication.Application.AutomationSecurity =
                MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            var wordDocument = wordApplication.Documents.Open(filename);
            const WdExportFormat exportFormat =
                WdExportFormat.wdExportFormatPDF;
            wordDocument.ExportAsFixedFormat(outputPath, exportFormat);
            wordDocument.Close(false);
            wordApplication.Quit();
            return outputPath;
        }

        private static string GetOutputPath(string inputPath)
        {
            return Path.Combine(OutputPath, Path.GetFileName(inputPath));
        }

        private static string GetProcessingPath(string inputPath)
        {
            return Path.Combine(ProcessingPath, Path.GetFileName(inputPath));
        }

        private static bool IsPdf(string filename)
        {
            return string.Equals(
                Path.GetExtension(filename),
                PdfFileExtension,
                StringComparison.InvariantCultureIgnoreCase);
        }

        private static string MovePdfToOutput(string filename)
        {
            var outputPath = GetOutputPath(filename);
            File.Move(filename, outputPath);
            return outputPath;
        }

        private static void ProcessPdf(
            string filename, Field field, Script script)
        {
            using (var pdfReader = new PdfReader(filename))
            {
                using (var pdfStamper = new PdfStamper(
                    pdfReader,
                    new FileStream(GetOutputPath(filename), FileMode.Create)))
                {
                    if (field != null)
                    {
                        AddFieldToPdf(
                            field, pdfStamper, pdfReader.NumberOfPages);
                    }

                    if (script != null)
                    {
                        AddScriptToPdf(script, pdfStamper);
                    }
                }
            }
        }
    }
}