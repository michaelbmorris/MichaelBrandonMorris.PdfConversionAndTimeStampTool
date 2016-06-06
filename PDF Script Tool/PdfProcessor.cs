//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfProcessor.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfTool
{
    using Application = Microsoft.Office.Interop.Word.Application;
    using Directory = System.IO.Directory;
    using DirectoryInfo = System.IO.DirectoryInfo;
    using Environment = System.Environment;
    using File = System.IO.File;
    using FileInfo = System.IO.FileInfo;
    using FileMode = System.IO.FileMode;
    using FileStream = System.IO.FileStream;
    using IProgress = System.IProgress<ProgressReport>;
    using List = System.Collections.Generic.List<string>;
    using Path = System.IO.Path;
    using PdfAction = iTextSharp.text.pdf.PdfAction;
    using PdfFormField = iTextSharp.text.pdf.PdfFormField;
    using PdfName = iTextSharp.text.pdf.PdfName;
    using PdfReader = iTextSharp.text.pdf.PdfReader;
    using PdfStamper = iTextSharp.text.pdf.PdfStamper;
    using PdfWriter = iTextSharp.text.pdf.PdfWriter;
    using Rectangle = iTextSharp.text.Rectangle;
    using Resources = Properties.Resources;
    using SpecialFolder = System.Environment.SpecialFolder;
    using StringComparison = System.StringComparison;
    using Task = System.Threading.Tasks.Task;
    using TextField = iTextSharp.text.pdf.TextField;
    using WdExportFormat = Microsoft.Office.Interop.Word.WdExportFormat;

    internal class PdfProcessor
    {
        private const int EveryOtherPage = 2;

        private const int EveryPage = 1;

        private const int FirstPageNumber = 1;

        private const int SecondPageNumber = 2;

        internal PdfProcessor()
        {
            Directory.CreateDirectory(OutputPath);
            Directory.CreateDirectory(ProcessingPath);
            ClearProcessing();
        }

        internal static string OutputPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(SpecialFolder.MyDocuments),
                    Resources.RootFolderName);
            }
        }

        internal static string ProcessingPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(SpecialFolder.ApplicationData),
                    Resources.RootFolderName);
            }
        }

        internal List Files { get; set; }

        internal static void ClearProcessing()
        {
            var processingDirectory = new DirectoryInfo(ProcessingPath);
            foreach (FileInfo file in processingDirectory.GetFiles())
            {
                file.Delete();
            }
        }

        internal static string CopyFileToProcessing(string filename)
        {
            var processingPath = GetProcessingPath(filename);
            File.Copy(filename, processingPath);
            return processingPath;
        }

        internal async Task ProcessFiles(
            IProgress progress, Field field = null, Script script = null)
        {
            await Task.Run(() =>
            {
                for (int i = 0; i < Files.Count; i++)
                {
                    var currentFile = Files[i];
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

                    progress.Report(new ProgressReport
                    {
                        Total = Files.Count,
                        CurrentCount = i + 1
                    });
                }
            });
        }

        private static void AddFieldToPage(
            Field field,
            int pageNumber,
            PdfStamper pdfStamper,
            PdfFormField parentField)
        {
            var textField = new TextField(
                pdfStamper.Writer,
                new Rectangle(
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
            int pageNumber = field.Pages == Pages.Last ?
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
                int increment = field.Pages == Pages.All ?
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
            }

            pdfStamper.Writer.SetAdditionalAction(
                actionType, pdfAction);
        }

        private static string ConvertToPdf(string filename)
        {
            var outputFilename = Path.GetFileNameWithoutExtension(filename)
                + Resources.PdfFileExtension;
            var outputPath = Path.Combine(ProcessingPath, outputFilename);
            var wordApplication = new Application();
            var wordDocument = wordApplication.Documents.Open(filename);
            var exportFormat = WdExportFormat.wdExportFormatPDF;
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
                Resources.PdfFileExtension,
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
                            field,
                            pdfStamper,
                            pdfReader.NumberOfPages);
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