using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace XlsxToXmlDll
{
    public class XlsxManager
    {
        static Action<bool,string> logCallback = null;

        /// <summary>
        /// 工具根路径
        /// </summary>
        static string toolRootPath = "";

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="toolRootPath"></param>
        /// <param name="logCallback"></param>
        public static void Init(string toolRootPath, Action<bool,string> logCallback)
        {
            UnInit();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            XlsxManager.toolRootPath = toolRootPath;
            XlsxManager.logCallback = logCallback;
            ConfigData.Init(toolRootPath+"/Config.xml");
        }

        /// <summary>
        /// 关闭
        /// </summary>
        public static void UnInit()
        {
            ConfigData.UnInit();
        }

        /// <summary>
        /// 日志
        /// </summary>
        /// <param name="isNormal"></param>
        /// <param name="content"></param>
        public static void Log(bool isNormal,string content)
        {
            logCallback?.Invoke(isNormal,content);
        }

        public static void SetImportXlsxAbsolutePath(string path)
        {
            ConfigData configData = ConfigData.GetSingle();
            configData.ImportXlsxAbsolutePath = path;
            configData.NeedSave = true;
            configData.ImportXlsxRelativePath = $"/{GetRelativePath(GetToolRootPath(), configData.ImportXlsxAbsolutePath)}/";
        }

        public static string GetImportXlsxAbsolutePath()
        {
            ConfigData configData = ConfigData.GetSingle();
            return configData.ImportXlsxAbsolutePath;
        }

        /// <summary>
        /// 获得code名称list
        /// </summary>
        /// <returns></returns>
        public static List<string> GetCodeNameList()
        {
            ConfigData configData = ConfigData.GetSingle();
            return configData.CodeConfigDataMap.Keys.ToList();
        }

        /// <summary>
        /// 检测是否是xlsx文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="showLog"></param>
        /// <returns></returns>
        public static bool CheckIsXlsxFile(string filePath,bool showLog=true)
        {
            if (!File.Exists(filePath))
            {
                if (showLog)
                {
                    Log(false,$"选择的文件：{filePath}不存在。");
                }
                return false;
            }
            else if (!filePath.EndsWith(".xlsx"))
            {
                if (showLog)
                {
                    Log(false, $"选择的文件：{filePath}不是xlsx文件。");
                }
                return false;
            }
            else if (filePath.Contains("~$"))
            {
                if (showLog)
                {
                    Log(true, $"选择的文件：{filePath}是~$临时文件。");
                }
                return false;
            }
            return true;
        }

        /// <summary>
        /// 获得差异文件的相对路径
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDifferentFileRelativePathList()
        {
            ConfigData configData = ConfigData.GetSingle();
            List<string> fileRelaticePathList = new List<string>();
            try
            {
                string differentFileListString = "";
                if (configData.ProjectVersionTool == "git")
                {
                    differentFileListString = ProcessHelper.Run("git.exe", GetImportXlsxAbsolutePath(), $"status {GetImportXlsxAbsolutePath()} -s");
                }
                else if (configData.ProjectVersionTool == "svn")
                {
                    differentFileListString = ProcessHelper.Run("svn.exe", GetImportXlsxAbsolutePath(), $"status");
                }
                if (string.IsNullOrEmpty(differentFileListString))
                {
                    Log(true, "没有差异文件！");
                }
                else
                {
                    string[] differentFileList = differentFileListString.Split('\n');
                    foreach (string differentFileString in differentFileList)
                    {
                        string differentFilePath = differentFileString.Trim();
                        if (differentFilePath.StartsWith("M") || differentFilePath.StartsWith("?"))
                        {
                            string[] differentFilePathParamList = differentFilePath.Split(' ');
                            string filePath = GetImportXlsxAbsolutePath() + "/" + differentFilePathParamList[differentFilePathParamList.Length - 1];
                            if (CheckIsXlsxFile(filePath, false))
                            {
                                fileRelaticePathList.Add(differentFilePathParamList[differentFilePathParamList.Length - 1]);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Log(false, "选择差异文件失败！可能是没有在配置文件Config.xml中的ProjectVersionTool属性设置svn或git，又或者是安装svn或git时没添加命令行工具。");
                Log(false, exception.Message);
                Log(false, exception.StackTrace);
                throw;
            }
            return fileRelaticePathList;
        }

        /// <summary>
        /// 设置指定代码名称是否需要导出
        /// </summary>
        /// <param name="codeName"></param>
        /// <param name="needExport"></param>
        public static void SetNeedExportByCodeName(string codeName,bool needExport)
        {
            ConfigData configData = ConfigData.GetSingle();
            if (configData.CodeConfigDataMap.ContainsKey(codeName))
            {
                configData.CodeConfigDataMap[codeName].NeedExport = needExport;
            }
            else
            {
                Log(false, $"SetNeedExportByCodeName:{codeName} 代码名称类型不存在");
            }
        }

        /// <summary>
        /// 获得指定代码名称是否需要导出
        /// </summary>
        /// <param name="codeName"></param>
        public static bool GetNeedExportByCodeName(string codeName)
        {
            ConfigData configData = ConfigData.GetSingle();
            if (configData.CodeConfigDataMap.ContainsKey(codeName))
            {
                return configData.CodeConfigDataMap[codeName].NeedExport;
            }
            else
            {
                Log(false, $"GetNeedExportByCodeName:{codeName} 代码名称类型不存在");
            }
            return false;
        }

        public static void SetExportXmlAbsolutePathByCodeName(string codeName,string path)
        {
            ConfigData configData = ConfigData.GetSingle();
            if (configData.CodeConfigDataMap.ContainsKey(codeName))
            {
                configData.CodeConfigDataMap[codeName].ExportXmlAbsolutePath = path;
                configData.CodeConfigDataMap[codeName].ExportXmlRelativePath = $"/{GetRelativePath(XlsxManager.GetToolRootPath(), path)}/";
                configData.NeedSave = true;
            }
            else
            {
                Log(false, $"SetNeedExportByCodeName:{codeName} 代码名称类型不存在");
            }
        }
        public static string GetExportXmlAbsolutePathByCodeName(string codeName)
        {
            ConfigData configData = ConfigData.GetSingle();
            if (configData.CodeConfigDataMap.ContainsKey(codeName))
            {
                return configData.CodeConfigDataMap[codeName].ExportXmlAbsolutePath;
            }
            else
            {
                Log(false, $"GetExportXmlAbsolutePathByCodeName:{codeName} 代码名称类型不存在");
            }
            return "";
        }
        public static void SetExportCodeAbsolutePathByCodeName(string codeName, string path)
        {
            ConfigData configData = ConfigData.GetSingle();
            if (configData.CodeConfigDataMap.ContainsKey(codeName))
            {
                configData.CodeConfigDataMap[codeName].ExportCodeAbsolutePath = path;
                configData.CodeConfigDataMap[codeName].ExportCodeRelativePath = $"/{GetRelativePath(XlsxManager.GetToolRootPath(), path)}/";
                configData.NeedSave = true;
            }
            else
            {
                Log(false, $"SetExportCodeAbsolutePathByCodeName:{codeName} 代码名称类型不存在");
            }
        }
        public static string GetExportCodeAbsolutePathByCodeName(string codeName)
        {
            ConfigData configData = ConfigData.GetSingle();
            if (configData.CodeConfigDataMap.ContainsKey(codeName))
            {
                return configData.CodeConfigDataMap[codeName].ExportCodeAbsolutePath;
            }
            else
            {
                Log(false, $"GetExportCodeAbsolutePathByCodeName:{codeName} 代码名称类型不存在");
            }
            return "";
        }

        /// <summary>
        /// 生成文件
        /// </summary>
        /// <param name="fileRelaticePathList"></param>
        /// <param name="resultCallback"></param>
        /// <param name="progressCallback"></param>
        public static void GenFile(List<string> fileRelaticePathList,Action<bool> resultCallback, Action<float> progressCallback)
        {
            ConfigData configData = ConfigData.GetSingle();
            foreach (var item in configData.CodeConfigDataMap)
            {
                if (item.Value.NeedExport)
                {
                    if (!Directory.Exists(configData.CodeConfigDataMap[item.Key].ExportXmlAbsolutePath))
                    {
                        Log(false, $"xml配置文件根路径:{configData.CodeConfigDataMap[item.Key].ExportXmlAbsolutePath}不存在！");
                        resultCallback?.Invoke(false);
                        return;
                    }
                    if (!Directory.Exists(configData.CodeConfigDataMap[item.Key].ExportCodeAbsolutePath))
                    {
                        Log(false, $"{item.Key}代码文件根路径:{configData.CodeConfigDataMap[item.Key].ExportCodeAbsolutePath}不存在！");
                        resultCallback?.Invoke(false);
                        return;
                    }
                }
            }
            progressCallback?.Invoke(0);
            int finishFileNumber = 0;
            string currentXlsxFilePath = "";
            try
            {
                Task.Run(() =>
                {
                    Log(true, $"开始生成文件！");
                    foreach (string xlsxFileRelativePath in fileRelaticePathList)
                    {
                        string xlsxFilePath = GetImportXlsxAbsolutePath() + "/" + xlsxFileRelativePath;
                        currentXlsxFilePath = xlsxFilePath;

                        XlsxFile xlsxFile = new XlsxFile(xlsxFilePath);
                        foreach (var item in configData.CodeConfigDataMap)
                        {
                            if (item.Value.NeedExport)
                            {
                                xlsxFile.Export(item.Key, xlsxFileRelativePath);
                            }
                        }
                        finishFileNumber++;
                        progressCallback?.Invoke((float)finishFileNumber / fileRelaticePathList.Count);
                    }
                    Log(true, $"生成文件结束！");
                    resultCallback?.Invoke(true);
                });
            }
            catch (CustomException customException)
            {
                Log(false, $"生成文件失败！{currentXlsxFilePath}");
                Log(false, $"{customException.customMessage}");
                resultCallback?.Invoke(false);
            }
            catch (Exception exception)
            {
                Log(false, $"生成文件失败！{currentXlsxFilePath}");
                Log(false, exception.Message);
                Log(false, exception.StackTrace);
                resultCallback?.Invoke(false);
            }
        }

        /// <summary>
        /// 导出所有配置的总览
        /// </summary>
        /// <param name="allRecorderOverviewFilePath"></param>
        public static void ExportAllRecorderOverview(string allRecorderOverviewFilePath)
        {
            ConfigData configData = ConfigData.GetSingle();
            DirectoryInfo xlsxRootDirectoryInfo = new DirectoryInfo(configData.ImportXlsxAbsolutePath);
            XDocument doc = new XDocument();
            XElement rootElement = ExportAllRecorderOverviewInDirectory(xlsxRootDirectoryInfo);
            doc.Add(rootElement);
            //保存时忽略声明
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;
            using (FileStream fileStream = new FileStream(allRecorderOverviewFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
                {
                    doc.Save(xmlWriter);
                }
            }
        }

        /// <summary>
        /// 导出文件夹下所有配置的总览
        /// </summary>
        /// <param name="rootDirectoryInfo"></param>
        static XElement ExportAllRecorderOverviewInDirectory(DirectoryInfo rootDirectoryInfo)
        {
            XElement node = new XElement(rootDirectoryInfo.Name);
            foreach (DirectoryInfo directorInfo in rootDirectoryInfo.GetDirectories())
            {
                XElement chileNode = ExportAllRecorderOverviewInDirectory(directorInfo);
                node.Add(chileNode);
            }
            foreach (FileInfo fileInfo in rootDirectoryInfo.GetFiles())
            {
                if (!CheckIsXlsxFile(fileInfo.FullName, false))
                {
                    continue;
                }
                XlsxFile xlsxFile = new XlsxFile(fileInfo.FullName);
                XElement chileNode = xlsxFile.ExportRecorderOverview();
                node.Add(chileNode);
            }
            return node;
        }

        /// <summary>
        /// 获得两个文件的相对路径，因为Path.GetRelativePath在.net core2之后再有，所以临时使用，判断规则和返回值有可能有问题
        /// </summary>
        /// <param name="fromPath"></param>
        /// <param name="toPath"></param>
        /// <returns></returns>
        public static string GetRelativePath(string fromPath, string toPath)
        {
            if (String.IsNullOrEmpty(fromPath)) throw new ArgumentNullException("fromPath");
            if (String.IsNullOrEmpty(toPath)) throw new ArgumentNullException("toPath");

            if(Directory.Exists(fromPath) && !(fromPath.EndsWith("\\")||fromPath.EndsWith("/")))
            {
                fromPath += "\\";
            }
            if (Directory.Exists(toPath) && (!toPath.EndsWith("\\") || !toPath.EndsWith("/")))
            {
                toPath += "\\";
            }

            Uri fromUri = new Uri(fromPath);
            Uri toUri = new Uri(toPath);

            if (fromUri.Scheme != toUri.Scheme) { return toPath; } // path can't be made relative.

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            String relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (toUri.Scheme.Equals("file", StringComparison.InvariantCultureIgnoreCase))
            {
                relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            }
            if(relativePath.EndsWith("\\") || relativePath.EndsWith("/"))
            {
                relativePath = relativePath.Substring(0,relativePath.Length-1);
            }
            return relativePath;
        }

        public static string GetToolRootPath()
        {
            return toolRootPath;
        }
    }
}
