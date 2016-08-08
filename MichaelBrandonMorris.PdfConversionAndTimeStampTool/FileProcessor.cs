using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using iTextSharp.text.pdf;
using MichaelBrandonMorris.Extensions.CollectionExtensions;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;
using Microsoft.Office.Interop.PowerPoint;
using static System.Environment;
using static System.Environment.SpecialFolder;
using static System.IO.Path;
using static Microsoft.Office.Core.MsoAutomationSecurity;
using Action = System.Action;
using Task = System.Threading.Tasks.Task;
using static System.IO.Directory;
using Progress = System.IProgress<System.Tuple<int, int, string>>;
using ProgressReport = System.Tuple<int, int, string>;
using static Microsoft.Office.Interop.Word.WdExportFormat;
using static Microsoft.Office.Interop.Excel.XlFixedFormatType;
using static Microsoft.Office.Interop.PowerPoint.PpFixedFormatType;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using static System.IO.FileMode;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    internal class FileProcessor
    {
        private const string DateTimeFormat = "yyyyMMddTHHmmss";
        private const string FolderName = "PDF Conversion And Time Stamp Tool";

        private static readonly string OutputFolderPath = Combine(
            GetFolderPath(MyDocuments), FolderName);

        private static readonly string ProcessingFolderPath = Combine(
            GetFolderPath(ApplicationData), FolderName);

        private List<string> Log
        {
            get;
        } = new List<string>();

        internal FileProcessor(
            IList<string> fileNames,
            Progress progress,
            Field field = null,
            Script script = null)
        {
            FileNames = fileNames;
            Progress = progress;
            Field = field;
            Script = script;
            CreateDirectory(OutputFolderPath);
            CreateDirectory(ProcessingFolderPath);
        }

        private CancellationTokenSource CancellationTokenSource
        {
            get;
        } = new CancellationTokenSource();

        private Field Field
        {
            get;
        }

        private IList<string> FileNames
        {
            get;
        }

        private Progress Progress
        {
            get;
        }

        private Script Script
        {
            get;
        }

        internal void Cancel()
        {
            CancellationTokenSource.Cancel();
        }

        internal async Task Execute()
        {
            var task = Task.Run(
                () =>
                {
                    var count = 0;
                    foreach (var t in FileNames)
                    {
                        var currentFileName = t;

                        if (!IsPdf(currentFileName))
                        {
                            currentFileName = ConvertToPdf(currentFileName);
                        }

                        if (Field != null || Script != null)
                        {
                            ProcessPdf(currentFileName);
                        }

                        MoveToOutput(currentFileName);

                        Progress.Report(
                            new ProgressReport(
                                ++count, FileNames.Count, currentFileName));
                    }
                },
                CancellationTokenSource.Token);

            await task;

            if (!Log.IsEmpty())
            {
                var now = DateTime.Now;
                var logFileName = $"Log - {now.ToString(DateTimeFormat)}";
                var logFilePath = Combine(OutputFolderPath, logFileName);
                File.WriteAllLines(logFilePath, Log);
            }
        }

        private void ConvertExcelToPdf(
            string fileName, string processingPath)
        {
            var excelApplication = new ExcelApplication
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false,
                    AutomationSecurity = msoAutomationSecurityForceDisable
                };

            var excelWorkbook = excelApplication.Workbooks.Open(fileName);

            if (excelWorkbook == null)
            {
                excelApplication.Quit();
                return;
            }

            try
            {
                excelWorkbook.ExportAsFixedFormat(xlTypePDF, processingPath);
            }
            catch (Exception e)
            {
                Log.Add(e.Message);
            }
            finally
            {
                excelWorkbook.Close();
                excelApplication.Quit();
            }
        }

        private void ConvertPowerPointToPdf(
            string fileName, string processingPath)
        {
            var powerPointApplication = new PowerPointApplication
            {
                DisplayAlerts = PpAlertLevel.ppAlertsNone,
                AutomationSecurity = msoAutomationSecurityForceDisable
            };

            var powerPointPresentation = 
                powerPointApplication.Presentations.Open(fileName);

            if (powerPointPresentation == null)
            {
                powerPointApplication.Quit();
                return;
            }

            try
            {
                powerPointPresentation.ExportAsFixedFormat(
                    processingPath, ppFixedFormatTypePDF);
            }
            catch (Exception e)
            {
                Log.Add(e.Message);
            }
            finally
            {
                powerPointPresentation.Close();
                powerPointApplication.Quit();
            }
        }

        private void ConvertWordToPdf(string fileName, string processingPath)
        {
            var wordApplication = new WordApplication
                {
                    AutomationSecurity = msoAutomationSecurityForceDisable
                };


            wordApplication.Application.AutomationSecurity =
                msoAutomationSecurityForceDisable;

            var wordDocument = wordApplication.Documents.Open(fileName);

            if (wordDocument == null)
            {
                wordApplication.Quit();
                return;
            }

            try
            {
                wordDocument.ExportAsFixedFormat(processingPath, wdExportFormatPDF);
            }
            catch (Exception e)
            {
                Log.Add(e.Message);
            }
            finally
            {
                wordDocument.Close(false);
                wordApplication.Quit();
            }
            
            
        }

        private static bool IsPdf(string fileName)
        {
            return GetExtension(fileName).EqualsOrdinalIgnoreCase(".pdf");
        }

        private string ConvertToPdf(string fileName)
        {
            var processingFileName = GetFileNameWithoutExtension(fileName) +
                                     ".pdf";

            var processingPath = Combine(
                ProcessingFolderPath, processingFileName);

            var extension = GetExtension(fileName)?.ToLower();

            if (extension.IsNullOrWhiteSpace())
            {
                throw new FileNotFoundException(
                    $"The file '{fileName}'does not have an extension.");
            }

            var extensionMapping = new Dictionary<string, Action>
            {
                [".doc"] = () =>
                {
                    ConvertWordToPdf(fileName, processingPath);
                },
                [".docx"] = () =>
                {
                    ConvertWordToPdf(fileName, processingPath);
                },
                [".xls"] = () =>
                {
                    ConvertExcelToPdf(fileName, processingPath);
                },
                [".xlsx"] = () =>
                {
                    ConvertExcelToPdf(fileName, processingPath);
                },
                [".ppt"] = () =>
                {
                    ConvertPowerPointToPdf(fileName, processingPath);
                },
                [".pptx"] = () =>
                {
                    ConvertPowerPointToPdf(fileName, processingPath);
                }
            };

            Debug.Assert(extension != null, "extension != null");
            extensionMapping[extension]();
            return processingPath;
        }

        private void MoveToOutput(string fileName)
        {
        }

        private void ProcessPdf(string fileName)
        {
            using (var pdfReader = new PdfReader(fileName))
            using(var fileStream = new FileStream(
                GetOutputPath(fileName), Create))
            using (var pdfStamper = new PdfStamper(pdfReader, fileStream))
            {
                if (Field != null)
                {
                    AddFieldToPdf(pdfStamper, pdfReader.NumberOfPages);
                }

                if (Script != null)
                {
                    
                }
            }
        }

        private void AddFieldToPdf(PdfStamper pdfStamper, int numberOfPages)
        {
            
        }

        private void AddScriptToPdf(PdfStamper pdfStamper)
        {
            var pdfAction = PdfAction.JavaScript(
                Script.Text, pdfStamper.Writer);

            var actionTypeMapping = new Dictionary<ScriptTiming, PdfName>
            {
                [ScriptTiming.DidPrint] = PdfWriter.DID_PRINT,
                [ScriptTiming.DidSave] = PdfWriter.DID_SAVE,
                [ScriptTiming.WillPrint] = PdfWriter.WILL_PRINT,
                [ScriptTiming.WillSave] = PdfWriter.WILL_SAVE
            };
        }

        private string GetOutputPath(string fileName)
        {
            var outputFileName = GetFileName(fileName);

            if (outputFileName == null)
            {
                throw new Exception(
                    $"The name of the file '{fileName}' " +
                    "is incorrectly formatted.");
            }

            return  Combine(OutputFolderPath, outputFileName);
        }
    }
}