using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Linq;
using XlsxToXmlDll;

namespace XlsxToXml
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<string, CodeInfoControl> codeInfoControlMap = new Dictionary<string, CodeInfoControl>();

        public MainWindow()
        {
            InitializeComponent();

            XlsxManager.Init(Environment.CurrentDirectory, Log);

            //初始化xlsx文件列表的右键菜单
            MenuItem deleteMenuItem = new MenuItem();
            deleteMenuItem.Header = "删除";
            deleteMenuItem.Click += (sender, e) =>
            {
                RemoveCurrentSelectFile();
            };
            fileListBox.ContextMenu.Items.Add(deleteMenuItem);
            fileListBox.SelectionMode = SelectionMode.Extended;

            //初始化配置
            importXlsxRootPathTextBox.Text = XlsxManager.GetImportXlsxAbsolutePath();

            List<string> codeNameList = XlsxManager.GetCodeNameList();
            for (int i = 0; i < codeNameList.Count; i++)
            {
                CodeInfoControl codeInfoControl = new CodeInfoControl();
                codeInfoControlMap.Add(codeNameList[i], codeInfoControl);
                codeInfoControl.SetCodeName(codeNameList[i]);
                codeInfoControlGrid.Children.Add(codeInfoControl);
                if (codeNameList.Count == 1)
                {
                    codeInfoControl.VerticalAlignment = VerticalAlignment.Center;
                }
                else
                {
                    codeInfoControl.VerticalAlignment = VerticalAlignment.Top;
                    codeInfoControl.Margin = new Thickness(0, i * 120, 0, 0);
                }
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            XlsxManager.UnInit();
        }

        private void SelectImportXlsxRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            dialog.InitialDirectory = importXlsxRootPathTextBox.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                importXlsxRootPathTextBox.Text = dialog.FileName;
                XlsxManager.SetImportXlsxAbsolutePath(dialog.FileName);
                fileListBox.Items.Clear();
            }
        }

        private void SelectDifferentFileButton_Click(object sender, RoutedEventArgs e)
        {
            List<string> fileList = XlsxManager.GetDifferentFileRelativePathList();
            foreach (var item in fileList)
            {
                AddRelativeFileToFileList(item);
            }
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;//等于true表示可以选择多个文件
            fileDialog.DefaultExt = ".xlsx";
            fileDialog.Filter = "Excel文件|*.xlsx";
            fileDialog.InitialDirectory = importXlsxRootPathTextBox.Text;
            if (fileDialog.ShowDialog() == true)
            {
                foreach (string filePath in fileDialog.FileNames)
                {
                    AddFileToFileList(filePath);
                }
            }
        }

        private void SelectAllFileButton_Click(object sender, RoutedEventArgs e)
        {
            AddDirectoryToFileList(importXlsxRootPathTextBox.Text);
        }

        private void FileListBox_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.RightButton == MouseButtonState.Pressed)
            {
                e.Handled = true;
            }
        }

        private void ImportXlsxRootPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            fileListBox.Items.Clear();
        }

        private void GenFileButton_Click(object sender, RoutedEventArgs e)
        {
            logRichTextBox.Document.Blocks.Clear();
            if (fileListBox.Items.Count <= 0)
            {
                Log(true,$"没有需要生成的文件！");
                return;
            }
            GenFile();
        }

        /// <summary>
        /// 输出日志
        /// </summary>
        /// <param name="isNormal"></param>
        /// <param name="content"></param>
        void Log(bool isNormal,string content)
        {
            logRichTextBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                logRichTextBox.Document.Blocks.Add(new Paragraph(new Run(content) { Foreground = isNormal ? Brushes.Black:Brushes.Red }));
            }));
        }

        void ProgressCallback(float percent)
        {
            gemFileProgressBar.Dispatcher.Invoke(() =>
            {
                gemFileProgressBar.Value = 100 * percent;
            });
         }

        /// <summary>
        /// 添加文件夹到列表中
        /// </summary>
        /// <param name="directoryPath"></param>
        void AddDirectoryToFileList(string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(directoryPath);
                foreach (FileInfo fileInfo in directoryInfo.GetFiles())
                {
                    AddFileToFileList(fileInfo.FullName);
                }
                foreach (DirectoryInfo childDirectoryInfo in directoryInfo.GetDirectories())
                {
                    AddDirectoryToFileList(childDirectoryInfo.FullName);
                }
            }
        }

        /// <summary>
        /// 添加文件到列表中
        /// </summary>
        /// <param name="filePath"></param>
        void AddFileToFileList(string filePath)
        {
            if (filePath.StartsWith(importXlsxRootPathTextBox.Text))
            {
                if (XlsxManager.CheckIsXlsxFile(filePath))
                {
                    string fileRelativePath = System.IO.Path.GetRelativePath(importXlsxRootPathTextBox.Text, filePath);
                    AddRelativeFileToFileList(fileRelativePath);
                }
            }
            else
            {
                Log(false,$"选择的文件：{filePath}不在xlsx根路径下。");
            }
        }

        /// <summary>
        /// 添加文件到列表中
        /// </summary>
        /// <param name="filePath"></param>
        void AddRelativeFileToFileList(string fileRelativePath)
        {
            if (!fileListBox.Items.Contains(fileRelativePath))
            {
                fileListBox.Items.Add(fileRelativePath);
            }
        }

        void GenFile()
        {
            List<string> fileRelaticePathList = new List<string>();
            foreach (string item in fileListBox.Items)
            {
                fileRelaticePathList.Add(item);
            }
            XlsxManager.GenFile(fileRelaticePathList, (result)=>
            {
                if (result)
                {
                    MessageBox.Show("生成文件成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("生成文件失败！", "提示", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }, ProgressCallback);
        }

        private void FileListBox_Drop(object sender, DragEventArgs e)
        {
            if(e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] fileList = e.Data.GetData(DataFormats.FileDrop) as string[];
                foreach (string filePath in fileList)
                {
                    if (Directory.Exists(filePath))
                    {
                        AddDirectoryToFileList(filePath);
                    }
                    else
                    {
                        AddFileToFileList(filePath);
                    }
                }
            }
        }

        private void FileListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key==Key.Delete)
            {
                RemoveCurrentSelectFile();
            }
        }

        /// <summary>
        /// 移除当前选择文件
        /// </summary>
        void RemoveCurrentSelectFile()
        {
            for (int i = fileListBox.SelectedItems.Count - 1; i >= 0; i--)
            {
                fileListBox.Items.Remove(fileListBox.SelectedItems[i]);
            }
        }

        private void ExportAllRecorderOverviewButton_Click(object sender, RoutedEventArgs e)
        {
            string allRecorderOverviewFileName = "AllRecorderOverview.xml";
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string allRecorderOverviewRootPath = dialog.FileName;
                string allRecorderOverviewFilePath = $"{allRecorderOverviewRootPath}\\{allRecorderOverviewFileName}";
                bool needExport = false;
                if (File.Exists(allRecorderOverviewFilePath))
                {
                    if (MessageBox.Show($"当前路径已经存在文件：{allRecorderOverviewFileName}，是否覆盖", "提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        needExport = true;
                    }
                }
                else
                {
                    needExport = true;
                }
                if(needExport)
                {
                    XlsxManager.ExportAllRecorderOverview(allRecorderOverviewFilePath);
                    if (MessageBox.Show($"配置的总览文件已经导出，是否打开？", "提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        Process process = new Process();
                        ProcessStartInfo processStartInfo = new ProcessStartInfo(allRecorderOverviewFilePath);
                        process.StartInfo = processStartInfo;
                        process.StartInfo.UseShellExecute = true;
                        process.Start();
                    }
                }
            }
        }
    }
}
