using Info.Hnbc.InvigilatorExcel.ExcelHandler.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Info.Hnbc.InvigilatorExcel.WpfClient
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly TaskScheduler _syncContextTaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();

        public MainWindow()
        {
            InitializeComponent();
        }
         

        private List<string> Excels = new List<string>();
        private void Button_Click_excels(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Filter = "Excel文件.xlsx|*.xlsx";
            fileDialog.Multiselect = true;
            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Excels = fileDialog.FileNames.ToList();
                text_excel_files.Text = $"共选择了{Excels.Count()}个文件";
            }
        }

        string output = null;
        private void Button_Click_folder(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folder = new System.Windows.Forms.FolderBrowserDialog();
            folder.Description = "请选择输出文件夹";
            if (folder.ShowDialog() == System.Windows.Forms.DialogResult.OK || folder.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
                output = folder.SelectedPath;
                text_output.Text = folder.SelectedPath;
            }


        }

        private void Button_Click_start(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew(SchedulerWork);
           
            log.Text += $"{DateTime.Now}：全部完成\r\n";
        }
        private void SchedulerWork()
        {
            foreach (var p in Excels)
            {
                Task.Factory.StartNew(HandlerExcel, p).Wait();
                 
            } 
        }

        private void Log(string msg)
        {
            log.Text += $"{DateTime.Now}：{msg}\r\n";
        }
        private void L(string msg)
        {
            Task.Factory.StartNew(() => Log(msg),
                    new CancellationTokenSource().Token, TaskCreationOptions.None, _syncContextTaskScheduler).Wait();
        }

        private void HandlerExcel(object pathobj)
        {
            var path = pathobj.ToString();
            var outputdir = output;
            var name = System.IO.Path.GetFileName(path);
            L($"开始处理{ name}...");
            var jiankao = BianhaoService.Test(path);

            L($"{name} 解析完成\r\n");
            jiankao.SetJianKao();
            L($"{name} 分配监考\r\n");
            if (jiankao.KaoshiInfos.Count(c => c.Room == null) > 0)
            {
                L($" Error：教室不足，部分考试未分配教室\r\n");
            }

            if (jiankao.KaoshiInfos.Count(c => c.Teachers.Count < 2) > 0)
            {
                L($" Error：监考老师不足，部分考试未分配监考\r\n");
            }

            //var noroom = jiankao.KaoshiInfos.Where(w => w.Room == null).ToList();
            //var noteacher = jiankao.KaoshiInfos.Where(w => w.Teachers.Count < 2).ToList();

            L($"开始导出\r\n");
            var exportfile = System.IO.Path.Combine(outputdir,"处理完成"+name);
            BianhaoService.Export(jiankao,path, exportfile);
            L($"{name} 完成√\r\n");
        }
    }
}
