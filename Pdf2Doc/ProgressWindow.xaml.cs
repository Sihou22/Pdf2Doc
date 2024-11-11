using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Pdf2Doc
{
    /// <summary>
    /// ProgressWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ProgressWindow : Window
    {

        private CancellationTokenSource cancellationTokenSource;
        public ProgressWindow()
        {
            InitializeComponent();
        }

        public void UpdateMaximum(int Max)
        {
            progressBar.Maximum = Max;
        }
        public void UpdateProgress(int progress)
        {
            progressBar.Value = progress;
            lblProgress.Content = $"已转换: {progress}/{progressBar.Maximum}页";
        }

        public void SetCancellationTokenSource(CancellationTokenSource cancellationTokenSource)
        {
            this.cancellationTokenSource = cancellationTokenSource;
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            cancellationTokenSource?.Cancel();
        }
    }
}
