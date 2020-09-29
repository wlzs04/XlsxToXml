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
        public static string Run(string processPath,string workDirectory, string arguments)
        {
            var info = new ProcessStartInfo(processPath, arguments)
            {
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardInput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                WorkingDirectory = workDirectory
            };
            var process = new Process
            {
                StartInfo = info,
            };
            process.Start();
            string content = process.StandardOutput.ReadToEnd();
            process.Close();
            return content;
        }
    }
}
