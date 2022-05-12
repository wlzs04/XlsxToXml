using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace XlsxToXml
{
    /// <summary>
    /// 进程帮助类
    /// </summary>
    public class ProcessHelper
    {
        /// <summary>
        /// 获得指定程序运行后的输出结果
        /// </summary>
        /// <param name="processPath"></param>
        /// <param name="workDirectory"></param>
        /// <param name="arguments"></param>
        /// <returns></returns>
        public static string RunWithResult(string processPath, string workDirectory, string arguments)
        {
            //mac在获取环境变量时缺少部分路径，所以临时添加，然后删除
            //string oldPath = Environment.GetEnvironmentVariable("PATH");
            //Environment.SetEnvironmentVariable("PATH", oldPath + ":/usr/local/bin");
            //Environment.SetEnvironmentVariable("PATH", oldPath);
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = processPath;
            info.WorkingDirectory = workDirectory;
            info.Arguments = arguments;
            info.CreateNoWindow = true;
            info.RedirectStandardOutput = true;
            info.RedirectStandardInput = true;
            info.RedirectStandardError = true;
            info.UseShellExecute = false;

            Process process = Process.Start(info);
            process.WaitForExit();
            string content = process.StandardOutput.ReadToEnd();
            process.Close();
            return content;
        }
    }
}
