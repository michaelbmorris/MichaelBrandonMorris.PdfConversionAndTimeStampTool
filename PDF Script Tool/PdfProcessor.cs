//-----------------------------------------------------------------------------------------------------------
// <copyright file="PdfProcessor.cs" company="Michael Brandon Morris">
//     Copyright © Michael Brandon Morris 2016
// </copyright>
//-----------------------------------------------------------------------------------------------------------

namespace PdfScriptTool
{
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
            System.IO.Directory.CreateDirectory(RootPath);
            System.IO.Directory.CreateDirectory(OutputRootPath);
            System.IO.Directory.CreateDirectory(ProcessingPath);
        }

        /// <summary>
        /// Gets the folder output is stored in.
        /// "Output" in the root folder.
        /// </summary>
        internal static string OutputRootPath
        {
            get
            {
                return System.IO.Path.Combine(
                    RootPath, Properties.Resources.OutputFolderName);
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
                return System.IO.Path.Combine(
                    RootPath, Properties.Resources.ProcessingFolderName);
            }
        }

        /// <summary>
        /// Gets the root folder the program works in.
        /// "PDF Script Tool" in the user's "My Documents" folder.
        /// </summary>
        internal static string RootPath
        {
            get
            {
                return System.IO.Path.Combine(
                    System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.MyDocuments),
                    Properties.Resources.RootFolderName);
            }
        }

        /// <summary>
        /// Gets or sets the list of files (paths) to be processed.
        /// </summary>
        internal System.Collections.Generic.List<string> Files
        {
            get; set;
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
        internal async System.Threading.Tasks.Task ProcessPdfs(
            System.IProgress<ProgressReport> progress,
            Field field = null,
            Script script = null)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 0; i < this.Files.Count; i++)
                {
                    var currentFile = Files[i].ToString();
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
            iTextSharp.text.pdf.PdfStamper pdfStamper,
            iTextSharp.text.pdf.PdfFormField parentField)
        {
            var textField = new iTextSharp.text.pdf.TextField(
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

        /// <summary>
        /// Adds a field to a PDF document.
        /// </summary>
        /// <param name="field">The field to add.</param>
        /// <param name="pdfStamper">The PDF stamper for the document.</param>
        /// <param name="numberOfPages">
        /// The number of pages in the document.
        /// </param>
        private static void AddFieldToPdf(
            Field field,
            iTextSharp.text.pdf.PdfStamper pdfStamper,
            int numberOfPages)
        {
            var parentField = iTextSharp.text.pdf.PdfFormField.CreateTextField(
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
            Script script, iTextSharp.text.pdf.PdfStamper pdfStamper)
        {
            var pdfAction = iTextSharp.text.pdf.PdfAction.JavaScript(
                script.ScriptText, pdfStamper.Writer);

            iTextSharp.text.pdf.PdfName actionType = null;

            switch (script.ScriptEvent)
            {
                case ScriptEvent.DidPrint:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .DID_PRINT;
                    break;

                case ScriptEvent.DidSave:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .DID_SAVE;
                    break;

                case ScriptEvent.WillPrint:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .WILL_PRINT;
                    break;

                case ScriptEvent.WillSave:
                    actionType = iTextSharp.text.pdf.PdfWriter
                            .WILL_SAVE;
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
            var outputFilename =
                System.IO.Path.GetFileNameWithoutExtension(filename)
                + Properties.Resources.PdfFileExtension;
            var outputPath = System.IO.Path.Combine(
                ProcessingPath,
                outputFilename);
            var wordApplication
                = new Microsoft.Office.Interop.Word.Application();
            var wordDocument = wordApplication.Documents.Open(filename);
            var exportFormat =
                Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
            wordDocument.ExportAsFixedFormat(
                outputPath,
                exportFormat);
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
            return System.IO.Path.Combine(
                OutputRootPath,
                System.IO.Path.GetFileName(inputPath));
        }

        /// <summary>
        /// Checks if a file is a PDF document.
        /// </summary>
        /// <param name="filename">The path of the file to check.</param>
        /// <returns>Whether or not the file is a PDF document.</returns>
        private static bool IsPdf(string filename)
        {
            return string.Equals(
                System.IO.Path.GetExtension(filename),
                Properties.Resources.PdfFileExtension,
                System.StringComparison.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Moves a PDF file to the output folder.
        /// </summary>
        /// <param name="filename">The PDF file to move.</param>
        private static void MovePdfToOutput(string filename)
        {
            System.IO.File.Move(filename, GetOutputPath(filename));
        }

        /// <summary>
        /// Adds features to a PDF document.
        /// </summary>
        /// <param name="filename">The path of the PDF document.</param>
        /// <param name="field">The field to add.</param>
        /// <param name="script">The script to add.</param>
        private static void ProcessPdf(string filename, Field field, Script script)
        {
            using (var pdfReader = new iTextSharp.text.pdf.PdfReader(filename))
            {
                using (var pdfStamper = new iTextSharp.text.pdf.PdfStamper(
                    pdfReader,
                    new System.IO.FileStream(
                        GetOutputPath(filename),
                        System.IO.FileMode.Create)))
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