using CsvHelper;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;

namespace MsWordAutomation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "docx";
            if (openFileDialog.ShowDialog() == true)
                WordFilePath.Text = openFileDialog.FileName;
        }

        private void SelectExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "csv";
            if (openFileDialog.ShowDialog() == true)
                ExcelFilePath.Text = openFileDialog.FileName;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application fileOpen = new Microsoft.Office.Interop.Word.Application(); ;
            try
            {
                var header = File.ReadAllLines(ExcelFilePath.Text).FirstOrDefault().Split(',');
                fileOpen.Visible = false;
                fileOpen.ScreenUpdating = false;
                using (var reader = new StreamReader(ExcelFilePath.Text))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<dynamic>();
                    foreach (var line in records)
                    {
                        //Open a already existing word file into the new document created
                        Microsoft.Office.Interop.Word.Document document = fileOpen.Documents.Open(WordFilePath.Text, ReadOnly: true);
                        //Make the file visible 
                        document.Activate();
                        //The FindAndReplace takes the text to find under any formatting and replaces it with the
                        //new text with the same exact formmating (e.g red bold text will be replaced with red bold text)
                        var path = SaveFormat.Text;
                        foreach (var heading in header)
                        {
                            var x = (((IDictionary<String, Object>)line)[heading]).ToString();
                            x = x.Replace("\n", "\n\r");
                            FindAndReplace(fileOpen, "$$" + heading + "$$", x);
                            path = path.Replace("$$" + heading + "$$", x);
                        }
                        var a = (object)Path.Combine(OutputFolder.Text, path + ".docx");
                        document.SaveAs2(ref a);
                        document.Close();
                    }
                    //Save the editted file in a specified location
                    //Can use SaveAs instead of SaveAs2 and just give it a name to have it saved by default
                    //to the documents folder
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                fileOpen?.Quit();
            }
        }

        static void FindAndReplace(Microsoft.Office.Interop.Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == true)
            {
                OutputFolder.Text = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
            }
        }

        private void FolderForDocx_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == true)
            {
                DocxFolder.Text = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
            }
        }

        private void FolderForPdf_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == true)
            {
                PdfFolder.Text = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
            }
        }

        private void ConvertToPdf_Click(object sender, RoutedEventArgs e)
        {
            // Create a new Microsoft Word application object
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            word.Visible = false;
            word.ScreenUpdating = false;

            // Get list of Word files in specified directory
            DirectoryInfo dirInfo = new DirectoryInfo(DocxFolder.Text);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.docx");
            foreach (FileInfo wordFile in wordFiles)
            {
                // Cast as Object for word Open method
                Object filename = (Object)wordFile.FullName;

                // Use the dummy value as a placeholder for optional arguments
                Document doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                object outputFileName = Path.Combine(PdfFolder.Text, wordFile.Name.Replace(".docx", ".pdf"));
                object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.                
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;
            }
            word.Quit();
        }
    }
}
