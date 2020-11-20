using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
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

namespace XlsxToXml
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ConfigData configData = null;

        public MainWindow()
        {
            InitializeComponent();

            //设置XLSX相关
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            XLSXFile.SetLogCallback(Log);

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
            configData = ConfigData.GetSingle();
            importXlsxRootPathTextBox.Text = System.IO.Path.GetFullPath(Environment.CurrentDirectory + configData.ImportXlsxRelativePath);
            exportXmlRootPathTextBox.Text = System.IO.Path.GetFullPath(Environment.CurrentDirectory + configData.ExportXmlRelativePath);
            exportCSRootPathTextBox.Text = System.IO.Path.GetFullPath(Environment.CurrentDirectory + configData.ExportCSRelativePath);
            configData.ExportCSAbsolutePath = exportCSRootPathTextBox.Text;
            if (File.Exists(Environment.CurrentDirectory + configData.CSRecorderTemplateFileRelativePath))
            {
                using (StreamReader streamReader = new StreamReader(Environment.CurrentDirectory + configData.CSRecorderTemplateFileRelativePath))
                {
                    XLSXFile.SetCSRecorderTemplateContent(streamReader.ReadToEnd());
                }
            }
            else
            {
                Log("缺少CSRecorder模板！");
            }
            if (File.Exists(Environment.CurrentDirectory + configData.CSEnumTemplateFileRelativePath))
            {
                using (StreamReader streamReader = new StreamReader(Environment.CurrentDirectory + configData.CSEnumTemplateFileRelativePath))
                {
                    XLSXFile.SetCSEnumTemplateContent(streamReader.ReadToEnd());
                }
            }
            else
            {
                Log("缺少CSEnum模板！");
            }
        }

        private void SelectImportXlsxRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            dialog.InitialDirectory = importXlsxRootPathTextBox.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                importXlsxRootPathTextBox.Text = dialog.FileName;
                configData.ImportXlsxRelativePath = $"/{System.IO.Path.GetRelativePath(Environment.CurrentDirectory, importXlsxRootPathTextBox.Text)}/";
                fileListBox.Items.Clear();
            }
        }

        private void SelectDifferentFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string differentFileListString = "";
                if (ConfigData.GetSingle().ProjectVersionTool == "git")
                {
                    differentFileListString = ProcessHelper.Run("git.exe", importXlsxRootPathTextBox.Text, $"status {importXlsxRootPathTextBox.Text} -s");
                }
                else if (ConfigData.GetSingle().ProjectVersionTool == "svn")
                {
                    differentFileListString = ProcessHelper.Run("svn.exe", importXlsxRootPathTextBox.Text, $"status");
                }
                if (string.IsNullOrEmpty(differentFileListString))
                {
                    Log("没有差异文件！");
                }
                else
                {
                    string[] differentFileList = differentFileListString.Split('\n');
                    foreach (string differentFileString in differentFileList)
                    {
                        string differentFilePath = differentFileString.Trim();
                        if (differentFilePath.StartsWith('M') || differentFilePath.StartsWith("?"))
                        {
                            string[] differentFilePathParamList = differentFilePath.Split(' ');
                            AddFileToFileList(importXlsxRootPathTextBox.Text + "/" + differentFilePathParamList[differentFilePathParamList.Length - 1]);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("选择差异文件失败！可能是没有在配置文件Config.xml中的ProjectVersionTool属性设置svn或git，又或者是安装svn或git时没添加命令行工具。");
                Log(exception.Message);
                Log(exception.StackTrace);
                throw;
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

        private void SelectExportXmlRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            dialog.InitialDirectory = exportXmlRootPathTextBox.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportXmlRootPathTextBox.Text = dialog.FileName;
                configData.ExportXmlRelativePath = $"/{System.IO.Path.GetRelativePath(Environment.CurrentDirectory, exportXmlRootPathTextBox.Text)}/";
            }
        }

        private void OpenExportXmlRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            ProcessHelper.Run("explorer.exe","", exportXmlRootPathTextBox.Text);
        }

        private void SelectExportCSRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            dialog.InitialDirectory = exportCSRootPathTextBox.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportCSRootPathTextBox.Text = dialog.FileName;
                configData.ExportCSAbsolutePath = exportCSRootPathTextBox.Text;
                configData.ExportCSRelativePath = $"/{System.IO.Path.GetRelativePath(Environment.CurrentDirectory, exportCSRootPathTextBox.Text)}/";
            }
        }

        private void OpenExportCSRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            ProcessHelper.Run("explorer.exe", "", exportCSRootPathTextBox.Text);
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
            logTextBox.Text = "";
            if (fileListBox.Items.Count <= 0)
            {
                Log($"没有需要生成的文件！");
                return;
            }

            GenFile();
        }

        /// <summary>
        /// 输出日志
        /// </summary>
        /// <param name="content"></param>
        void Log(string content)
        {
            logTextBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                logTextBox.Text += content + "\n";
            }));
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
                if(!File.Exists(filePath))
                {
                    Log($"选择的文件：{filePath}不存在。");
                }
                else if(!filePath.EndsWith(".xlsx"))
                {
                    Log($"选择的文件：{filePath}不是xlsx文件。");
                }
                else if (filePath.Contains("~$"))
                {
                    //Log($"选择的文件：{filePath}是~$临时文件。");
                }
                else
                {
                    string fileRelativePath = System.IO.Path.GetRelativePath(importXlsxRootPathTextBox.Text, filePath);
                    if (!fileListBox.Items.Contains(fileRelativePath))
                    {
                        fileListBox.Items.Add(fileRelativePath);
                    }
                }
            }
            else
            {
                Log($"选择的文件：{filePath}不在xlsx根路径下。");
            }
        }

        async void GenFile()
        {
            List<string> fileRelaticePathList = new List<string>();
            foreach (string item in fileListBox.Items)
            {
                fileRelaticePathList.Add(item);
            }
            string importXlsxRootPathText = importXlsxRootPathTextBox.Text;
            string exportXmlRootPathText = exportXmlRootPathTextBox.Text;
            string exportCSRootPathText = exportCSRootPathTextBox.Text;
            
            bool needExportXml = needGenXmlFileCheckBox.IsChecked.HasValue && (bool)needGenXmlFileCheckBox.IsChecked;
            bool needExportCS = needGenCSFileCheckBox.IsChecked.HasValue && (bool)needGenCSFileCheckBox.IsChecked;

            if(!needExportXml && !needExportCS)
            {
                return;
            }
            if (needExportXml && !Directory.Exists(exportXmlRootPathText))
            {
                Log($"xml配置文件根路径:{exportXmlRootPathText}不存在！");
                return;
            }
            if (needExportCS && !Directory.Exists(exportCSRootPathText))
            {
                Log($"cs代码文件根路径:{exportCSRootPathText}不存在！");
                return;
            }

            string currentXlsxFilePath = "";
            try
            {
                await Task.Run(() =>
                {
                    Log($"开始生成文件！");
                    foreach (string xlsxFileRelativePath in fileRelaticePathList)
                    {
                        string xlsxFilePath = importXlsxRootPathText + "/" + xlsxFileRelativePath;
                        currentXlsxFilePath = xlsxFilePath;
                        FileInfo xlsxFileInfo = new FileInfo(xlsxFilePath);
                        string fileName = xlsxFileInfo.Name.Substring(0, xlsxFileInfo.Name.LastIndexOf('.'));

                        string xmlFilePath = exportXmlRootPathText + "/" + xlsxFileRelativePath;
                        xmlFilePath = xmlFilePath.Substring(0, xmlFilePath.LastIndexOf('.')) + ".xml";
                        string csFilePath = exportCSRootPathText + "/" + xlsxFileRelativePath;
                        csFilePath = csFilePath.Substring(0, csFilePath.LastIndexOf('.')) + ".cs";

                        XLSXFile xlsxFile = new XLSXFile(xlsxFilePath);
                        if (needExportXml)
                        {
                            xlsxFile.ExportXML(xmlFilePath);
                        }
                        if (needExportCS)
                        {
                            xlsxFile.ExportCS(csFilePath);
                        }
                    }
                    Log($"生成文件结束！");
                    MessageBox.Show("生成文件结束！");
                });
            }
            catch (Exception exception)
            {
                MessageBox.Show($"生成文件失败！{currentXlsxFilePath}");
                Log($"生成文件失败！{currentXlsxFilePath}");
                Log(exception.Message);
                Log(exception.StackTrace);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            configData.Save();
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
    }
}
