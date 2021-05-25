using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using XlsxToXmlDll;

namespace XlsxToXml
{
    /// <summary>
    /// CodeInfoControl.xaml 的交互逻辑
    /// </summary>
    public partial class CodeInfoControl : UserControl
    {
        string codeName = "";

        public CodeInfoControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 设置代码名称
        /// </summary>
        /// <param name="codeName"></param>
        public void SetCodeName(string codeName)
        {
            this.codeName = codeName;

            needGenCodeNameCheckBox.Content = codeName;
            exportXmlRootPathTextBox.Text = XlsxManager.GetExportXmlAbsolutePathByCodeName(codeName);
            exportCodeRootPathTextBox.Text = XlsxManager.GetExportCodeAbsolutePathByCodeName(codeName);
        }
        
        private void NeedGenCodeNameCheckBox_Click(object sender, RoutedEventArgs e)
        {
            XlsxManager.SetNeedExportByCodeName(codeName, (bool)needGenCodeNameCheckBox.IsChecked);
        }

        private void SelectExportXmlRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            dialog.InitialDirectory = exportXmlRootPathTextBox.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportXmlRootPathTextBox.Text = dialog.FileName;
                XlsxManager.SetExportXmlAbsolutePathByCodeName(codeName, dialog.FileName);
            }
        }

        private void OpenExportXmlRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            ProcessHelper.Run("explorer.exe", "", exportXmlRootPathTextBox.Text);
        }

        private void SelectExportCodeRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;//设置为选择文件夹
            dialog.InitialDirectory = exportCodeRootPathTextBox.Text;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                exportCodeRootPathTextBox.Text = dialog.FileName;
                XlsxManager.SetExportCodeAbsolutePathByCodeName(codeName, dialog.FileName);
            }
        }

        private void OpenExportCodeRootPathButton_Click(object sender, RoutedEventArgs e)
        {
            ProcessHelper.Run("explorer.exe", "", exportCodeRootPathTextBox.Text);
        }
    }
}
