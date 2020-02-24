# Multiple Thread Pdf Converter (doc, docx, ppt, pptx)

Run `PdfConverter.exe`

A File Dialog will open and you can select multiple files of doc, docx, ppt, pptx.

Progress show on Console, like 1/4, 2/4, 3/4.

Console will auto close when all convert work is finish.

Here is source code:

```C#
using System;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;

namespace PdfConverter
{
    class Program
    {

        public static int totalWork = 0;

        public static int finishedWorkNumber = 0;
        public static object Lock = new object();
        public static void FinishedWorkAddOne_ShowProgress()
        {
            lock (Lock) {
                finishedWorkNumber++;
                Console.WriteLine(finishedWorkNumber + " / " + totalWork);
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "請選擇需要轉換為 pdf 的 doc, docx, ppt, pptx 文件",
                Filter = "MEOW? (*.doc, *.docx, *.ppt, *pptx)|*.doc;*.docx;*.ppt;*.pptx"
            };

            if (fileDialog.ShowDialog() != DialogResult.OK) { fileDialog.Dispose(); return; }

            string[] names = fileDialog.FileNames;

            totalWork = names.Length;

            Thread[] threads = new Thread[totalWork];

            for (int i = 0; i < totalWork; i++)
            {
                string file = names[i];

                string extension = System.IO.Path.GetExtension(file);

                string sourcePath = file;
                string targetPath = file.Substring(0, file.Length - extension.Length) + ".pdf";
                string[] passParameter = { sourcePath, targetPath };

                if (IsWord(extension))
                {
                    threads[i] = new Thread(WordToPDF);
                }
                else if (IsPowerPoint(extension))
                {
                    threads[i] = new Thread(PowerPointToPDF);
                }

                threads[i].Start(passParameter);
                threads[i].Join();
            }

            fileDialog.Dispose();
        }

        public static bool IsWord(string extension) => extension.Equals(".doc") || extension.Equals(".docx");
        public static bool IsPowerPoint(string extension) => extension.Equals(".ppt") || extension.Equals(".pptx");

        public static void WordToPDF(object parameter)
        {
            string sourcePath = ((string[])parameter)[0];
            string targetPath = ((string[])parameter)[1];

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
                document.Close();
                application.Quit();
                FinishedWorkAddOne_ShowProgress();
            }
        }

        public static void PowerPointToPDF(object parameter)
        {
            string sourcePath = ((string[])parameter)[0];
            string targetPath = ((string[])parameter)[1];

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
                presentation.Close();
                application.Quit();
                FinishedWorkAddOne_ShowProgress();
            }
        }


    }
}
```
