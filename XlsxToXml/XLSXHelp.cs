using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace XlsxToXml
{
    /// <summary>
    /// xlsx文件类
    /// </summary>
    class XLSXFile
    {
        static Action<string> fileLogCallback = null;
        static string csClassTemplateContent = "";
        string xlsxFilePath = "";
        string fileName = "";
        DataRowCollection xlsxDataRowCollection = null;

        /// <summary>
        /// 属性名称
        /// </summary>
        List<string> propertyValueNameList = new List<string>();
        /// <summary>
        /// 是否需要导出
        /// </summary>
        List<bool> needExportList = new List<bool>();
        /// <summary>
        /// 类型名称
        /// </summary>
        List<string> propertyClassNameList = new List<string>();
        /// <summary>
        /// 规则描述，一般为空
        /// </summary>
        List<string> propertyDescriptionList = new List<string>();
        /// <summary>
        /// 配置名称，作为属性名称的注释
        /// </summary>
        List<string> propertyConfigNameList = new List<string>();

        public XLSXFile(string xlsxFilePath)
        {
            this.xlsxFilePath = xlsxFilePath;
            ReadExcel();
        }

        /// <summary>
        /// 设置log回调
        /// </summary>
        /// <param name="logCallback"></param>
        public static void SetLogCallback(Action<string> logCallback)
        {
            fileLogCallback = logCallback;
        }

        /// <summary>
        /// 设置c#类的模板内容
        /// </summary>
        /// <param name="csClassTemplateContent"></param>
        public static void SetCSClassTemplateContent(string content)
        {
            csClassTemplateContent = content;
        }

        /// <summary>
        /// 读取xlsx文件，只读取第一页(sheet1)的数据
        /// </summary>
        /// <returns></returns>
        void ReadExcel()
        {
            FileInfo xlsxFileInfo = new FileInfo(xlsxFilePath);
            fileName = xlsxFileInfo.Name.Substring(0, xlsxFileInfo.Name.LastIndexOf('.'));

            FileStream stream = File.Open(xlsxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            xlsxDataRowCollection = result.Tables[0].Rows;

            int rowCount = xlsxDataRowCollection.Count;
            if (rowCount <= 0)
            {
                fileLogCallback?.Invoke($"xlsx文件中没有内容：{xlsxFilePath}");
                return;
            }
            else if (xlsxDataRowCollection[1].ItemArray.Length > 100)
            {
                fileLogCallback?.Invoke($"xlsx文件中列数太多，超过100，请检查。如果需要扩充请修改代码。");
                return;
            }
            //xlsx文件格式：第一行，为属性名称
            foreach (object item in xlsxDataRowCollection[0].ItemArray)
            {
                propertyValueNameList.Add(item.ToString());
            }
            //xlsx文件格式：第二行，是否需要导出
            foreach (object item in xlsxDataRowCollection[1].ItemArray)
            {
                bool needExport = Convert.ToBoolean(item);
                needExportList.Add(needExport);
            }
            //xlsx文件格式：第三行，为类型名称
            foreach (object item in xlsxDataRowCollection[2].ItemArray)
            {
                propertyClassNameList.Add(item.ToString());
            }
            //xlsx文件格式：第四行，为规则描述，一般为空
            foreach (object item in xlsxDataRowCollection[3].ItemArray)
            {
                propertyDescriptionList.Add(item.ToString());
            }
            //xlsx文件格式：第五行，为配置名称，作为属性名称的注释
            foreach (object item in xlsxDataRowCollection[4].ItemArray)
            {
                propertyConfigNameList.Add(item.ToString());
            }
        }

        /// <summary>
        /// 将XLSX文件导出到XML文件
        /// </summary>
        /// <param name="exportXMLFilePath"></param>
        public void ExportXML(string exportXMLFilePath)
        {
            //先将文件名替换为配置名称
            string xmlFileName = ConfigData.GetSingle().XmlFileName.Replace("{recorderName}", fileName);
            exportXMLFilePath = Path.GetDirectoryName(exportXMLFilePath) + xmlFileName;
            fileLogCallback?.Invoke($"xlsx文件开始导出：{xlsxFilePath} -> {exportXMLFilePath}");

            int rowCount = xlsxDataRowCollection.Count;
            XDocument doc = new XDocument(new XElement(fileName));
            for (int rowIndex = 5; rowIndex < rowCount; rowIndex++)
            {
                object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                //有的文件会出现某一行内容为空的情况
                if (itemArray[0].ToString() == "")
                {
                    continue;
                }
                XElement recordNode = new XElement("Recorder");
                for (int i = 0; i < itemArray.Length; i++)
                {
                    if (needExportList[i])
                    {
                        recordNode.Add(new XAttribute(propertyValueNameList[i], itemArray[i]));
                    }
                }
                doc.Root.Add(recordNode);
            }

            //保存时忽略声明
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;
            FileInfo fileInfo = new FileInfo(exportXMLFilePath);
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            FileStream fileStream = new FileStream(exportXMLFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
            {
                doc.Save(xmlWriter);
            }

            fileLogCallback?.Invoke($"xml文件生成成功：{exportXMLFilePath}");
        }

        /// <summary>
        /// 将XLSX文件导出到CS文件
        /// </summary>
        /// <param name="exportCSFilePath"></param>
        public void ExportCS(string exportCSFilePath)
        {
            //先将文件名替换为配置名称
            string csFileName = ConfigData.GetSingle().CSClassFileName.Replace("{recorderName}", fileName);
            exportCSFilePath = Path.GetDirectoryName(exportCSFilePath) + csFileName;
            fileLogCallback?.Invoke($"xlsx文件开始导出：{xlsxFilePath} -> {exportCSFilePath}");

            FileInfo fileInfo = new FileInfo(exportCSFilePath);
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            FileStream fileStream = new FileStream(exportCSFilePath, FileMode.Create, FileAccess.ReadWrite);
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            {
                StringBuilder csClassContent = new StringBuilder(csClassTemplateContent);
                //替换类名
                csClassContent.Replace("{recorderName}", fileName);
                //替换属性模板
                Dictionary<string, string> propertyTemplateMap = ConfigData.GetSingle().CSClassPropertyTemplateMap;
                Dictionary<string, string> convertFunctionTemplateMap = ConfigData.GetSingle().ConvertFunctionTemplateMap;
                foreach (var property in propertyTemplateMap)
                {
                    StringBuilder propertyTotalContent = new StringBuilder();
                    for (int i = 0; i < propertyValueNameList.Count; i++)
                    {
                        StringBuilder propertyEveryContent = new StringBuilder(property.Value);
                        //根据类型替换转换方法模板
                        if(convertFunctionTemplateMap.ContainsKey(propertyClassNameList[i]))
                        {
                            propertyEveryContent.Replace("{convertFunction}", convertFunctionTemplateMap[propertyClassNameList[i]]);
                        }
                        else
                        {
                            propertyEveryContent.Replace("{convertFunction}", convertFunctionTemplateMap["default"]);
                        }
                        propertyEveryContent.Replace("{propertyConfigName}",propertyConfigNameList[i]);
                        propertyEveryContent.Replace("{propertyDescription}", propertyDescriptionList[i]);
                        propertyEveryContent.Replace("{propertyClassName}", propertyClassNameList[i]);
                        propertyEveryContent.Replace("{propertyValueName}", propertyValueNameList[i]);
                        propertyTotalContent.Append(propertyEveryContent.ToString());
                        if (i != propertyValueNameList.Count-1)
                        {
                            propertyTotalContent.AppendLine();
                        }
                    }
                    csClassContent.Replace($"{{{property.Key}}}", propertyTotalContent.ToString());
                }
                streamWriter.WriteLine(csClassContent.ToString());
                streamWriter.Flush();
            }
        }
    }
}
