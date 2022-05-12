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
        static Action<bool,object> logCallback = null;

        /// <summary>
        /// 工具根路径
        /// </summary>
        static string toolRootPath = "";

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="toolRootPath"></param>
        /// <param name="logCallback"></param>
        public static void Init(string toolRootPath, Action<bool,object> logCallback)
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
        public static void Log(bool isNormal,object content)
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
                configData.CodeConfigDataMap[codeName].ExportXmlRelativePath = $"/{GetRelativePath(GetToolRootPath(), path)}/";
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
                configData.CodeConfigDataMap[codeName].ExportCodeRelativePath = $"/{GetRelativePath(GetToolRootPath(), path)}/";
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
        /// 生成指定文件
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
                    if (!Directory.Exists(item.Value.ExportXmlAbsolutePath))
                    {
                        Log(false, $"xml配置文件根路径:{item.Value.ExportXmlAbsolutePath}不存在！");
                        resultCallback?.Invoke(false);
						return;
                    }
                    if (!Directory.Exists(item.Value.ExportCodeAbsolutePath))
                    {
                        Log(false, $"{item.Key}代码文件根路径:{item.Value.ExportCodeAbsolutePath}不存在！");
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
                Log(false, exception);
                resultCallback?.Invoke(false);
            }
        }

        /// <summary>
        /// 生成所有文件
        /// </summary>
        /// <param name="resultCallback"></param>
        /// <param name="progressCallback"></param>
        public static void GenAllFile(Action<bool> resultCallback, Action<float> progressCallback)
        {
            ConfigData configData = ConfigData.GetSingle();
            foreach (var item in configData.CodeConfigDataMap)
            {
                if (item.Value.NeedExport)
                {
                    if (!Directory.Exists(item.Value.ExportXmlAbsolutePath))
                    {
                        Log(false, $"xml配置文件根路径:{item.Value.ExportXmlAbsolutePath}不存在！");
                        resultCallback?.Invoke(false);
                        return;
                    }
                    if (!Directory.Exists(item.Value.ExportCodeAbsolutePath))
                    {
                        Log(false, $"{item.Key}代码文件根路径:{item.Value.ExportCodeAbsolutePath}不存在！");
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
                    Log(true, $"开始删除文件所有文件!");
                    foreach (var item in configData.CodeConfigDataMap)
                    {
                        if (item.Value.NeedExport)
                        {
                            Directory.Delete(item.Value.ExportXmlAbsolutePath,true);
                            Directory.CreateDirectory(item.Value.ExportXmlAbsolutePath);
                            Directory.Delete(item.Value.ExportCodeAbsolutePath, true);
                            Directory.CreateDirectory(item.Value.ExportCodeAbsolutePath);
                        }
                    }
                    Log(true, $"开始生成文件！");
                    List<string> fileRelaticePathList = GetDirectoryXlsxFileRelativeList(GetImportXlsxAbsolutePath());
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
                Log(false, exception);
                resultCallback?.Invoke(false);
            }

        }

        /// <summary>
        /// 添加文件夹到列表中
        /// </summary>
        /// <param name="directoryPath"></param>
        public static List<string> GetDirectoryXlsxFileRelativeList(string directoryPath)
        {
            List<string> fileRelativeList = new List<string>();
            if (Directory.Exists(directoryPath))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(directoryPath);
                foreach (FileInfo fileInfo in directoryInfo.GetFiles())
                {
                    if (CheckIsXlsxFile(fileInfo.FullName))
                    {
                        string fileRelativePath = GetRelativePath(GetImportXlsxAbsolutePath(), fileInfo.FullName);
                        fileRelativeList.Add(fileRelativePath);
                    }
                }
                foreach (DirectoryInfo childDirectoryInfo in directoryInfo.GetDirectories())
                {
                    fileRelativeList.AddRange(GetDirectoryXlsxFileRelativeList(childDirectoryInfo.FullName));
                }
            }
            return fileRelativeList;
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
        /// 获得两个文件的相对路径，为了保证平台之间的一致性，使用'/'作为路径分隔符
        /// </summary>
        /// <param name="fromPath"></param>
        /// <param name="toPath"></param>
        /// <returns></returns>
        public static string GetRelativePath(string fromPath, string toPath)
        {
            return Path.GetRelativePath(fromPath, toPath).Replace('\\', '/');
        }

        public static string GetToolRootPath()
        {
            return toolRootPath;
        }
    }
}
