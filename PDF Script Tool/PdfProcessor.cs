//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfProcessor.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
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

    /// <summary>
    /// The back end of the PDF Script Tool.
    /// </summary>
    internal class PdfProcessor
    {
        /// <summary>
        /// Increment of two to skip a page.
        /// </summary>
        private const int EveryOtherPage = 2;

        /// <summary>
        /// Increment of one to get every page.
        /// </summary>
        private const int EveryPage = 1;

        /// <summary>
        /// Page one should be the first page.
        /// </summary>
        private const int FirstPageNumber = 1;

        /// <summary>
        /// Page two should be the first page.
        /// </summary>
        private const int SecondPageNumber = 2;

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfProcessor"/> class.
        /// </summary>
        internal PdfProcessor()
        {
            Directory.CreateDirectory(OutputPath);
            Directory.CreateDirectory(ProcessingPath);
            ClearProcessing();
        }

        /// <summary>
        /// Gets the output folder for the program.
        /// RootFolderName in the user's "My Documents" folder.
        /// </summary>
        internal static string OutputPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(SpecialFolder.MyDocuments),
                    Resources.RootFolderName);
            }
        }

        /// <summary>
        /// Gets the folder processing files are stored in.
        /// "Processing" in the root folder.
        /// </summary>
        internal static string ProcessingPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(SpecialFolder.ApplicationData),
                    Resources.RootFolderName);
            }
        }

        /// <summary>
        /// Gets or sets the list of files (paths) to be processed.
        /// </summary>
        internal List Files { get; set; }

        /// <summary>
        /// Deletes all files in the Processing folder.
        /// </summary>
        internal static void ClearProcessing()
        {
            var processingDirectory = new DirectoryInfo(ProcessingPath);
            foreach (FileInfo file in processingDirectory.GetFiles())
            {
                file.Delete();
            }
        }

        /// <summary>
        /// Copies a file to the processing folder.
        /// </summary>
        /// <param name="filename">The file to move.</param>
        /// <returns>The processing path for the copied file.</returns>
        internal static string CopyFileToProcessing(string filename)
        {
            var processingPath = GetProcessingPath(filename);
            File.Copy(filename, processingPath);
            return processingPath;
        }

        /// <summary>
        /// Adds a field and a script to the currently selected files.
        /// </summary>
        /// <param name="progress">
        /// The object to which progress is reported.
        /// </param>
        /// <param name="field"> The field to be added to the files.</param>
        /// <param name="script">The script to be added to the files.</param>
        /// <returns>The completed task.</returns>
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

        /// <summary>
        /// Adds a field to a page of a PDF document.
        /// </summary>
        /// <param name="field">The field to add.</param>
        /// <param name="pageNumber">
        /// The page number on which the field will be added.
        /// </param>
        /// <param name="pdfStamper">The PDF stamper for the document.</param>
        /// <param name="parentField">The parent field.</param>
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

        /// <summary>
        /// Adds a field to a PDF document.
        /// </summary>
        /// <param name="field">The field to add.</param>
        /// <param name="pdfStamper">The PDF stamper for the document.</param>
        /// <param name="numberOfPages">
        /// The number of pages in the document.
        /// </param>
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

        /// <summary>
        /// Adds a script to a PDF document.
        /// </summary>
        /// <param name="script">The script to add.</param>
        /// <param name="pdfStamper">The PDF stamper for the document.</param>
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

        /// <summary>
        /// Converts a file to a PDF document.
        /// </summary>
        /// <param name="filename">The path of the file to convert.</param>
        /// <returns>The path of the converted PDF document.</returns>
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

        /// <summary>
        /// Gets the output path for a specified input file.
        /// </summary>
        /// <param name="inputPath">The input file path.</param>
        /// <returns>The output path for the file.</returns>
        private static string GetOutputPath(string inputPath)
        {
            return Path.Combine(OutputPath, Path.GetFileName(inputPath));
        }

        /// <summary>
        /// Gets the processing path for a specified input file.
        /// </summary>
        /// <param name="inputPath">The input file path.</param>
        /// <returns>The processing path for the file.</returns>
        private static string GetProcessingPath(string inputPath)
        {
            return Path.Combine(ProcessingPath, Path.GetFileName(inputPath));
        }

        /// <summary>
        /// Checks if a file is a PDF document.
        /// </summary>
        /// <param name="filename">The path of the file to check.</param>
        /// <returns>Whether or not the file is a PDF document.</returns>
        private static bool IsPdf(string filename)
        {
            return string.Equals(
                Path.GetExtension(filename),
                Resources.PdfFileExtension,
                StringComparison.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Moves a PDF file to the output folder.
        /// </summary>
        /// <param name="filename">The PDF file to move.</param>
        /// <returns>The output path for the moved file.</returns>
        private static string MovePdfToOutput(string filename)
        {
            var outputPath = GetOutputPath(filename);
            File.Move(filename, outputPath);
            return outputPath;
        }

        /// <summary>
        /// Adds features to a PDF document.
        /// </summary>
        /// <param name="filename">The path of the PDF document.</param>
        /// <param name="field">The field to add.</param>
        /// <param name="script">The script to add.</param>
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