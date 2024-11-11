using System;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.IO;
using Aspose.Pdf;
using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using Document = Aspose.Pdf.Document;
using Application = Microsoft.Office.Interop.Word.Application;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using System.Text.RegularExpressions;
using Task = System.Threading.Tasks.Task;
using System.Threading;

namespace Pdf2Doc
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        private OpenFileDialog opendPdf;
        private bool Work_Flag = false;
        private int converType = 0;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectBtn_Click(object sender, RoutedEventArgs e)
        {
            if (Work_Flag)
            {
                MessageBox.Show(this, "转换正在运行中","提示",MessageBoxButton.OK,MessageBoxImage.Error);
                return;
            }

            opendPdf = new OpenFileDialog();
            opendPdf.Filter = "PDF files (*.pdf)|*.pdf";

            if (opendPdf.ShowDialog() == true && opendPdf.FileName != "")
            {
                SelectedFileName.Text = opendPdf.SafeFileName;
                String FileUri = opendPdf.FileName;
                FileUri = FileUri.Replace("\\", "/");
                preViewer.Source = new Uri(FileUri);
            }
        }

        private void convertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (Work_Flag)
            {
                MessageBox.Show(this, "转换正在运行中", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (opendPdf is null || opendPdf.FileName == null || opendPdf.FileName =="")
            {
                MessageBox.Show(this, "请先选择PDF文件", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if(converType == 0)
            {
                MessageBox.Show(this, "请先选择需要转换的文件类型", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 打开保存文件对话框以选择保存路径
            SaveFileDialog saveFileDialog = new SaveFileDialog();


            switch(converType){
                case 1:
                    saveFileDialog.Filter = "Word Document (*.doc)|*.doc";
                    saveFileDialog.DefaultExt = "doc";
                    break;
                case 2:
                    saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                    saveFileDialog.DefaultExt = "docx";
                    break;
            }

            if(saveFileDialog.ShowDialog() == true)
                if(TextChange_Box.IsChecked == true)
                    ConvertPdf2Word(opendPdf.FileName, saveFileDialog.FileName, converType);
                else
                    ConvertPdf2File(opendPdf.FileName, saveFileDialog.FileName, converType);


        }


        private async void ConvertPdf2File(String input_path,String output_path,int output_type)
        {
            if(!File.Exists(input_path))
            {
                MessageBox.Show(this, "路径错误", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Document doc = new Document(input_path);
            Work_Flag = true;
            try
            {
                if (converType == 1)
                    await System.Threading.Tasks.Task.Run(() => { doc.Save(output_path, SaveFormat.Doc); });

                else
                    await System.Threading.Tasks.Task.Run(() => { doc.Save(output_path, SaveFormat.DocX); });
            }
            catch
            {
                MessageBox.Show(this, "文件转换失败", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            finally
            {
                Work_Flag = false;
            }
            MessageBox.Show(this, "文件转换成功", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;

        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {       
             converType = 1;
        }

        private void RadioButton_Checked_1(object sender, RoutedEventArgs e)
        {
            converType = 2;
        }

        private async void ConvertPdf2Word(String input_path, String output_path, int output_type)
        {
            if (!File.Exists(input_path))
            {
                MessageBox.Show(this, "路径错误", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Work_Flag = true;

            if(output_type == 1)
            {
                output_path += "x";
            }
            try
            {

                ProgressWindow progressWindow = new ProgressWindow();
                progressWindow.Owner = this;
                progressWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                progressWindow.Show();

                await Task.Run(() =>
                {

                    CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
                    progressWindow.Dispatcher.Invoke(() => progressWindow.SetCancellationTokenSource(cancellationTokenSource));


                    using WordprocessingDocument wordDocument = WordprocessingDocument.Create(output_path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    Body body = new Body();
                    mainPart.Document.Append(body);


                    using PdfReader reader = new PdfReader(input_path);

                    progressWindow.Dispatcher.Invoke(() => progressWindow.UpdateMaximum(reader.NumberOfPages));

                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        cancellationTokenSource.Token.ThrowIfCancellationRequested();
                        string text = PdfTextExtractor.GetTextFromPage(reader, i);
                        text = Regex.Replace(text, @"[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]", "");

                        string[] texts = text.Split("\n");

                        foreach (string context in texts)
                        {
                            Paragraph paragraph = new Paragraph(new Run(new Text(context)));
                            body.Append(paragraph);
                        }
                        Paragraph emptyParagraph = new Paragraph(new Run(new Text("")));
                        body.Append(emptyParagraph);

                        progressWindow.Dispatcher.Invoke(() => progressWindow.UpdateProgress(i));

                    }
                    Thread.Sleep(1000);
                    progressWindow.Dispatcher.Invoke(() => progressWindow.Close());
                });
            }
            catch
            {
                MessageBox.Show(this, "文件转换失败", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            finally
            {
                Work_Flag = false;
            }
            MessageBox.Show(this, "文件转换成功", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;

        }

        private void TextChange_Box_Checked(object sender, RoutedEventArgs e)
        {
            if(TextChange_Box.IsChecked == true)
            {
                DocBtn.IsEnabled = false;
                DocBtn.IsChecked = false;
            }
            else { 
                DocBtn.IsEnabled = true;
            }
        }
    }
}
