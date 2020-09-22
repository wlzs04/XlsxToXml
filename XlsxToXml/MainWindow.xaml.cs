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
                for (int i = fileListBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    fileListBox.Items.Remove(fileListBox.SelectedItems[i]);
                }
            };
            fileListBox.ContextMenu.Items.Add(deleteMenuItem);
            fileListBox.SelectionMode = SelectionMode.Multiple;

            //初始化配置
            configData = new ConfigData(Environment.CurrentDirectory + "/Config.xml");
            importXlsxRootPathTextBox.Text = System.IO.Path.GetFullPath(Environment.CurrentDirectory + configData.ImportXlsxRelativePath);
            exportXmlRootPathTextBox.Text = System.IO.Path.GetFullPath(Environment.CurrentDirectory + configData.ExportXmlRelativePath);
            exportCSRootPathTextBox.Text = System.IO.Path.GetFullPath(Environment.CurrentDirectory + configData.ExportCSRelativePath);
            if(File.Exists(Environment.CurrentDirectory + configData.CSClassTemplateFileRelativePath))
            {
                using (StreamReader streamReader = new StreamReader(Environment.CurrentDirectory + configData.CSClassTemplateFileRelativePath))
                {
                    XLSXFile.SetCSClassTemplateContent(streamReader.ReadToEnd());
                }
            }
            else
            {
                Log("缺少CSClass模板！");
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
                fileListBox.Items.Clear();
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
                fileListBox.Items.Clear();
                foreach (string filePath in fileDialog.FileNames)
                {
                    if (filePath.StartsWith(importXlsxRootPathTextBox.Text))
                    {
                        string fileRelativePath = filePath.Substring(importXlsxRootPathTextBox.Text.Length + 1);
                        if (!fileRelativePath.Contains("~$"))
                        {
                            fileListBox.Items.Add(fileRelativePath);
                        }
                    }
                    else
                    {
                        Log($"选择的文件：{filePath}不在xlsx根路径下。");
                    }
                }
            }
        }

        private void SelectExportXmlRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportXmlRootPathTextBox.Text = dialog.FileName;
            }
        }

        private void SelectExportCSRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportCSRootPathTextBox.Text = dialog.FileName;
            }
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
            Log($"开始生成文件！");
            if (fileListBox.Items.Count <= 0)
            {
                Log($"没有需要生成的文件！");
                return;
            }

            GenFile();

            Log($"生成文件结束！");
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

            await Task.Run(() =>
            {
                foreach (string xlsxFileRelativePath in fileRelaticePathList)
                {
                    string xlsxFilePath = importXlsxRootPathText + "/" + xlsxFileRelativePath;
                    FileInfo xlsxFileInfo = new FileInfo(xlsxFilePath);
                    string fileName = xlsxFileInfo.Name.Substring(0, xlsxFileInfo.Name.LastIndexOf('.'));
                    
                    string xmlFilePath = exportXmlRootPathText + "/" + xlsxFileRelativePath;
                    xmlFilePath = xmlFilePath.Substring(0, xmlFilePath.LastIndexOf('.')) + ".xml";
                    string csFilePath = exportCSRootPathText + "/" + xlsxFileRelativePath;
                    csFilePath = csFilePath.Substring(0, csFilePath.LastIndexOf('.')) + ".cs";

                    XLSXFile xlsxFile = new XLSXFile(xlsxFilePath);
                    if(needExportXml)
                    {
                        xlsxFile.ExportXML(xmlFilePath);
                    }
                    if(needExportCS)
                    {
                        xlsxFile.ExportCS(csFilePath);
                    }
                }
            });
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            configData.ImportXlsxRelativePath = $"/{System.IO.Path.GetRelativePath(Environment.CurrentDirectory, importXlsxRootPathTextBox.Text)}/";
            configData.ExportXmlRelativePath = $"/{System.IO.Path.GetRelativePath(Environment.CurrentDirectory, exportXmlRootPathTextBox.Text)}/";
            configData.ExportCSRelativePath = $"/{System.IO.Path.GetRelativePath(Environment.CurrentDirectory, exportCSRootPathTextBox.Text)}/";
            configData.Save();
        }
    }
}
