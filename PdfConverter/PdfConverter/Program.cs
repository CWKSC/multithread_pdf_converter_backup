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
        public const string patternDoc = @"(.*)(\.doc)$";
        public const string patternDocx = @"(.*)(\.docx)$";
        public const string patternPpt = @"(.*)(\.ppt)$";
        public const string patternPptx = @"(.*)(\.pptx)$";

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

            Thread[] threads = new Thread[names.Length];

            for (int i = 0; i < names.Length; i++)
            {
                string file = names[i];
                string extension = System.IO.Path.GetExtension(file);
                string[] path = { file, file.Substring(0, file.Length - extension.Length) + ".pdf" };
                Console.WriteLine(file.Substring(0, file.Length - extension.Length) + ".pdf");

                if(Regex.IsMatch(file, patternDoc) || Regex.IsMatch(file, patternDocx))
                {
                    threads[i] = new Thread(WordToPDF);
                }
                else if (Regex.IsMatch(file, patternPpt) || Regex.IsMatch(file, patternPptx))
                {
                    threads[i] = new Thread(PowerPointToPDF);
                }

                threads[i].Start(path);
                threads[i].Join();
            }

            fileDialog.Dispose();
        }


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
            }
        }

        public static void PowerPointToPDF(object parameter)
        {
            string sourcePath = ((string[])parameter)[0];
            string targetPath = ((string[])parameter)[1];

            Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation presentation = application.Presentations.Open(sourcePath);
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
            }
        }


    }
}
