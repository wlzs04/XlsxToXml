using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
        struct XlsxPropertyClass
        {
            /// <summary>
            /// 属性类型
            /// </summary>
            public string classType;
            /// <summary>
            /// 属性类名
            /// </summary>
            public string className;
            /// <summary>
            /// 属性参数
            /// </summary>
            public string classParam;
        }

        static Action<string> fileLogCallback = null;
        static string csClassTemplateContent = "";
        string xlsxFilePath = "";
        string fileName = "";
        DataRowCollection xlsxDataRowCollection = null;
        /// <summary>
        /// 需要导出的列
        /// </summary>
        List<int> needExportIndexList = new List<int>();
        /// <summary>
        /// 属性名称
        /// </summary>
        List<string> propertyValueNameList = new List<string>();
        /// <summary>
        /// 属性类名
        /// </summary>
        List<XlsxPropertyClass> propertyClassList = new List<XlsxPropertyClass>();
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
            int colCount = xlsxDataRowCollection[0].ItemArray.Length;
            if (rowCount < 5)
            {
                fileLogCallback?.Invoke($"xlsx文件中行数为{rowCount}，小于5，没有正确定义：{xlsxFilePath}");
                return;
            }
            else if (colCount > 100)
            {
                fileLogCallback?.Invoke($"xlsx文件中列数太多，为{colCount}，超过100，请检查。如果需要扩充请修改代码。");
                return;
            }
            //因为在xlsx配置中有可能出现空内容等问题，使属性列比实际要使用的列数多
            //所以先算出实际使用的列个数
            //xlsx文件格式：第一行，获得需要导出的列
            for (int i = 0; i < colCount; i++)
            {
                object item = xlsxDataRowCollection[0][i];
                if (item!=DBNull.Value && Convert.ToBoolean(item))
                {
                    needExportIndexList.Add(i);
                }
            }
            foreach (var index in needExportIndexList)
            {
                //xlsx文件格式：第二行，为属性名称
                propertyValueNameList.Add(xlsxDataRowCollection[1][index].ToString());

                //xlsx文件格式：第三行，为类型名称
                string propertyClassString = xlsxDataRowCollection[2][index].ToString();
                XlsxPropertyClass propertyClass = new XlsxPropertyClass();
                if (propertyClassString.Contains(' '))
                {
                    string[] propertyClassList = propertyClassString.Split(' ');
                    propertyClass.classType = propertyClassList[0];
                    if (propertyClassList.Length > 1)
                    {
                        propertyClass.className = propertyClassList[1];
                    }
                    if (propertyClassList.Length > 2)
                    {
                        propertyClass.classParam = propertyClassList[2];
                    }
                }
                else
                {
                    propertyClass.classType = propertyClassString;
                    propertyClass.className = propertyClassString;
                    propertyClass.classParam = "";
                }
                propertyClassList.Add(propertyClass);

                //xlsx文件格式：第四行，为规则描述，一般为空
                propertyDescriptionList.Add(xlsxDataRowCollection[3][index].ToString());

                //xlsx文件格式：第五行，为配置名称，作为属性名称的注释
                propertyConfigNameList.Add(xlsxDataRowCollection[4][index].ToString());
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
            exportXMLFilePath = Path.GetDirectoryName(exportXMLFilePath) + "/" + xmlFileName;
            fileLogCallback?.Invoke($"xlsx文件开始导出：{xlsxFilePath} -> {exportXMLFilePath}");

            int rowCount = xlsxDataRowCollection.Count;
            XDocument doc = new XDocument(new XElement(fileName));
            for (int rowIndex = 5; rowIndex < rowCount; rowIndex++)
            {
                object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                XElement recordNode = new XElement("Recorder");
                for (int i = 0; i < needExportIndexList.Count; i++)
                {
                    recordNode.Add(new XAttribute(propertyValueNameList[i], itemArray[needExportIndexList[i]]));
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
            exportCSFilePath = Path.GetDirectoryName(exportCSFilePath) + "/" + csFileName;
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
                foreach (var property in propertyTemplateMap)
                {
                    StringBuilder propertyTotalContent = new StringBuilder();
                    for (int i = 0; i < propertyValueNameList.Count; i++)
                    {
                        StringBuilder propertyEveryContent = new StringBuilder(property.Value);
                        //根据类型替换转换方法模板
                        propertyEveryContent.Replace("{convertFunction}", GetConvertFunctionByClassType(propertyClassList[i].classType));
                        if (propertyClassList[i].classType=="map")
                        {
                            string[] propertyClassNameList = propertyClassList[i].className.Split(',');
                            string propertyClassParam1 = propertyClassList[i].classParam[0].ToString();
                            string propertyClassParam2 = propertyClassList[i].classParam[1].ToString();
                            propertyEveryContent.Replace("{propertyClassName1}", propertyClassNameList[0]);
                            propertyEveryContent.Replace("{propertyClassName2}", propertyClassNameList[1]);
                            propertyEveryContent.Replace("{propertyClassParam1}", propertyClassParam1);
                            propertyEveryContent.Replace("{propertyClassParam2}", propertyClassParam2);
                            propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassNameList[0]).Replace("{propertyClassName}", propertyClassNameList[0]));
                            propertyEveryContent.Replace("{convertFunction2}", GetConvertFunctionByClassType(propertyClassNameList[1]).Replace("{propertyClassName}", propertyClassNameList[1]));
                            propertyEveryContent.Replace("{propertyClassName}", $"Dictionary<{propertyClassList[i].className}>");
                        }
                        else if(propertyClassList[i].classType == "list")
                        {
                            propertyEveryContent.Replace("{propertyClassParam1}", propertyClassList[i].classParam);
                            propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassList[i].className).Replace("{propertyClassName}", propertyClassList[i].className));
                            propertyEveryContent.Replace("{propertyClassName1}", propertyClassList[i].className);
                            propertyEveryContent.Replace("{propertyClassName}", $"List<{propertyClassList[i].className}>");
                        }
                        else
                        {
                            propertyEveryContent.Replace("{propertyClassName}", propertyClassList[i].className);
                        }
                        propertyEveryContent.Replace("{propertyConfigName}",propertyConfigNameList[i]);
                        propertyEveryContent.Replace("{propertyDescription}", propertyDescriptionList[i]);
                        propertyEveryContent.Replace("{propertyValueName}", propertyValueNameList[i]);
                        propertyTotalContent.Append(propertyEveryContent.ToString());
                        if (i != propertyValueNameList.Count-1)
                        {
                            propertyTotalContent.Append('\n');
                        }
                    }
                    csClassContent.Replace($"{{{property.Key}}}", propertyTotalContent.ToString());
                }
                streamWriter.Write(csClassContent.ToString());
                streamWriter.Flush();
            }
        }

        /// <summary>
        /// 通过类型获得转换方法
        /// </summary>
        /// <param name="classType"></param>
        string GetConvertFunctionByClassType(string classType)
        {
            if(ConfigData.GetSingle().ConvertFunctionTemplateMap.ContainsKey(classType))
            {
                return ConfigData.GetSingle().ConvertFunctionTemplateMap[classType];
            }
            else
            {
                return ConfigData.GetSingle().ConvertFunctionTemplateMap["custom"];
            }
        }
    }
}
