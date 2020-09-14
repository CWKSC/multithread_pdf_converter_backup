# Multiple Task Pdf Converter (doc, docx, ppt, pptx)

Run `PdfConverter.exe`  [Directly Download Link](https://github.com/CWKSC/MultipleThreadPdfConverter/raw/master/PdfConverter.exe)

A File Dialog will open and you can select multiple files of doc, docx, ppt, pptx.

Progress show on Console, like [1 / 4], [2 / 4], [3 / 4].

Also show the target path.

 Source code here:

```C#
using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Task = System.Threading.Tasks.Task;

namespace PdfConverter
{
    public class Program
    {

        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Please select doc, docx, ppt, pptx files that need to be converted to pdf",
                Filter = "MEOW? (*.doc, *.docx, *.ppt, *pptx)|*.doc;*.docx;*.ppt;*.pptx"
            };

            if (fileDialog.ShowDialog() != DialogResult.OK) { fileDialog.Dispose(); return; }

            Stopwatch stopwatch = Stopwatch.StartNew();

            string[] names = fileDialog.FileNames;

            totalWork = names.Length;

            Task[] tasks = new Task[totalWork];

            for (int i = 0; i < totalWork; i++)
            {
                string file = names[i];

                string extension = System.IO.Path.GetExtension(file);

                string sourcePath = file;
                string targetPath = file.Substring(0, file.Length - extension.Length) + ".pdf";

                if (IsWord(extension))
                {
                    tasks[i] = Task.Run(() => WordToPDF(sourcePath, targetPath));
                }
                else if (IsPowerPoint(extension))
                {
                    tasks[i] = Task.Run(() => PowerPointToPDF(sourcePath, targetPath));
                }
            }

            Thread.Sleep(2000);

            Task.WhenAll(tasks).Wait();

            fileDialog.Dispose();

            stopwatch.Stop();
            Console.WriteLine("\nAll work finsih! Spent " + stopwatch.Elapsed.TotalSeconds + " seconds");
            Console.Write("Press any Enter to exit ...");
            Console.ReadLine();
        }

        public static bool IsWord(string extension) => extension.Equals(".doc") || extension.Equals(".docx");
        public static bool IsPowerPoint(string extension) => extension.Equals(".ppt") || extension.Equals(".pptx");

        public static void WordToPDF(string sourcePath, string targetPath)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null;
            try
            {
                application.Visible = false;
                document = application.Documents.Open(sourcePath);
                document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                FinishedWorkAddOne_ShowProgress(targetPath);
                document.Close();
                application.Quit();
            }
        }

        public static void PowerPointToPDF(string sourcePath, string targetPath)
        {
            Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation presentation = application.Presentations.Open(sourcePath, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
            try
            {
                presentation.ExportAsFixedFormat(targetPath, PpFixedFormatType.ppFixedFormatTypePDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                FinishedWorkAddOne_ShowProgress(targetPath);
                presentation.Close();
                application.Quit();
            }
        }


        public static int totalWork = 0;

        public static int finishedWorkNumber = 0;
        public static readonly object Lock = new object();
        public static void FinishedWorkAddOne_ShowProgress(string targetPath)
        {
            lock (Lock) {
                finishedWorkNumber++;
                Console.WriteLine($"[{finishedWorkNumber} / {totalWork}] {targetPath}");
            }
        }

    }
}
```
