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
        /// <summary>
        /// xlsx类型
        /// </summary>
        enum XlsxEnum
        {
            /// <summary>
            /// 配置
            /// </summary>
            Recorder,
            /// <summary>
            /// 枚举
            /// </summary>
            Enum,
            /// <summary>
            /// 结构体
            /// </summary>
            Struct,
        }

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
        static string csRecorderTemplateContent = "";
        static string csEnumTemplateContent = "";
        static string csStructTemplateContent = "";
        string xlsxFilePath = "";
        string fileName = "";
        string className = "";
        XlsxEnum xlsxType = XlsxEnum.Recorder;
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

        //结构体特有
        string prefix = "";//前缀
        string suffix = "";//后缀
        string split = ";";//分割字符
        /// <summary>
        /// 配置默认值，作为属性的默认值
        /// </summary>
        List<string> propertyDefaultValueList = new List<string>();

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
        /// 设置配置类的模板内容
        /// </summary>
        /// <param name="content"></param>
        public static void SetCSRecorderTemplateContent(string content)
        {
            csRecorderTemplateContent = content;
        }

        /// <summary>
        /// 设置枚举类的模板内容
        /// </summary>
        /// <param name="content"></param>
        public static void SetCSEnumTemplateContent(string content)
        {
            csEnumTemplateContent = content;
        }

        /// <summary>
        /// 设置结构体的模板内容
        /// </summary>
        /// <param name="content"></param>
        public static void SetCSStructTemplateContent(string content)
        {
            csStructTemplateContent = content;
        }

        /// <summary>
        /// 读取xlsx文件，只读取第一页(sheet1)的数据
        /// </summary>
        /// <returns></returns>
        void ReadExcel()
        {
            FileInfo xlsxFileInfo = new FileInfo(xlsxFilePath);
            fileName = xlsxFileInfo.Name.Substring(0, xlsxFileInfo.Name.LastIndexOf('.'));
            if(fileName.Contains('.'))
            {
                int pointIndex = fileName.LastIndexOf('.');
                className = fileName.Substring(0,pointIndex);
                xlsxType = XlsxEnum.Recorder;
            }
            else if(fileName.EndsWith("Recorder"))
            {
                xlsxType = XlsxEnum.Recorder;
                className = fileName;
            }
            else if(fileName.EndsWith("Enum"))
            {
                xlsxType = XlsxEnum.Enum;
                className = fileName;
            }
            else if (fileName.EndsWith("Struct"))
            {
                xlsxType = XlsxEnum.Struct;
                className = fileName;
            }
            else
            {
                fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}，只能使用Recorder、Enum、Struct结尾，代表配置、枚举、结构体！");
            }
            FileStream stream = File.Open(xlsxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            xlsxDataRowCollection = result.Tables[0].Rows;
            
            int rowCount = xlsxDataRowCollection.Count;
            int colCount = xlsxDataRowCollection[0].ItemArray.Length;
            
            if(xlsxType == XlsxEnum.Recorder)
            {
                if (rowCount < 5)
                {
                    fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}中行数为{rowCount}，小于5，没有正确定义");
                    return;
                }
                if (colCount > 100)
                {
                    fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}中列数太多，为{colCount}，超过100，请检查。如果需要扩充请修改代码。");
                    return;
                }
                //因为在xlsx配置中有可能出现空内容等问题，使属性列比实际要使用的列数多
                //所以先算出实际使用的列个数
                //xlsx文件格式：第一行，获得需要导出的列
                for (int i = 0; i < colCount; i++)
                {
                    object item = xlsxDataRowCollection[0][i];
                    if (item != DBNull.Value && Convert.ToBoolean(item))
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
            else if(xlsxType == XlsxEnum.Enum)
            {
                if (colCount < 2)
                {
                    fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}中列数小于2，为{colCount}，请检查。需要保证一列名称，一列含义。");
                    return;
                }
                if (rowCount < 2)
                {
                    fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}中行数小于2，为{rowCount}，请检查。需要保证至少一行值。");
                    return;
                }
                for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                {
                    object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                    if (itemArray.Length < 1 || itemArray[0].ToString() == "")
                    {
                        fileLogCallback?.Invoke($"xlsx文件: {xlsxFilePath}表内容出现空值，可能是有意为之或内容未清除彻底，总行数:{rowCount},请核实row:{rowIndex + 1}：");
                        break;
                    }
                    propertyValueNameList.Add(itemArray[0].ToString());
                    propertyConfigNameList.Add(itemArray[1].ToString().ToString());
                }
            }
            else if (xlsxType == XlsxEnum.Struct)
            {
                if (colCount < 4)
                {
                    fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}中列数小于4，为{colCount}，请检查。需要保证结构正确。");
                    return;
                }
                if (rowCount < 5)
                {
                    fileLogCallback?.Invoke($"xlsx文件：{xlsxFilePath}中行数小于5，为{rowCount}，请检查。需要保证至少一行值。");
                    return;
                }
                prefix = xlsxDataRowCollection[1][0].ToString();
                suffix = xlsxDataRowCollection[1][1].ToString();
                split = xlsxDataRowCollection[1][2].ToString();
                for (int rowIndex = 4; rowIndex < rowCount; rowIndex++)
                {
                    object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                    if (itemArray.Length < 1 || itemArray[0].ToString() == "")
                    {
                        fileLogCallback?.Invoke($"xlsx文件: {xlsxFilePath}表内容出现空值，可能是有意为之或内容未清除彻底，总行数:{rowCount},请核实row:{rowIndex + 1}：");
                        break;
                    }
                    propertyValueNameList.Add(itemArray[0].ToString());
                    propertyConfigNameList.Add(itemArray[1].ToString());
                    XlsxPropertyClass propertyClass = new XlsxPropertyClass();
                    propertyClass.classType = itemArray[2].ToString();
                    propertyClass.className = propertyClass.classType;
                    propertyClass.classParam = "";
                    propertyClassList.Add(propertyClass);
                    propertyDefaultValueList.Add(itemArray[3].ToString().ToString());
                }
            }
        }

        /// <summary>
        /// 将XLSX文件导出到XML文件
        /// </summary>
        /// <param name="exportXMLFilePath"></param>
        public void ExportXML(string exportXMLFilePath)
        {
            if(xlsxType!=XlsxEnum.Recorder)
            {
                return;
            }
            //先将文件名替换为配置名称
            string xmlFileName = ConfigData.GetSingle().XmlFileName.Replace("{fileName}", fileName);
            exportXMLFilePath = Path.GetDirectoryName(exportXMLFilePath) + "/" + xmlFileName;
            //fileLogCallback?.Invoke($"xlsx文件开始导出：{xlsxFilePath} -> {exportXMLFilePath}");

            int rowCount = xlsxDataRowCollection.Count;
            XDocument doc = new XDocument(new XElement(className));
            for (int rowIndex = 5; rowIndex < rowCount; rowIndex++)
            {
                object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                if(itemArray.Length<1 || itemArray[0].ToString()=="")
                {
                    fileLogCallback?.Invoke($"xml文件{exportXMLFilePath}生成成功，但表内容出现空值，可能是有意为之或内容未清除彻底，总行数:{rowCount},请核实row:{rowIndex+1}：");
                    break;
                }
                XElement recordNode = new XElement("Recorder");
                for (int i = 0; i < needExportIndexList.Count; i++)
                {
                    if(propertyClassList[i].classType=="KeyValueMap")
                    {
                        XElement keyValueMap = new XElement(propertyValueNameList[i]);
                        recordNode.Add(keyValueMap);
                        int mapLength = Convert.ToInt32(propertyClassList[i].classParam);
                        for (int mapIndex = 0; mapIndex < mapLength; mapIndex++)
                        {
                            keyValueMap.Add(new XElement("KeyValue", new XAttribute("key", propertyValueNameList[i + mapIndex + 1]), new XAttribute("value", itemArray[needExportIndexList[i + mapIndex + 1]])));
                        }
                        i += mapLength;
                    }
                    else if (propertyClassList[i].classType == "ValueList")
                    {
                        XElement keyValueList = new XElement(propertyValueNameList[i]);
                        recordNode.Add(keyValueList);
                        int listLength = Convert.ToInt32(propertyClassList[i].classParam);
                        int realListLength = Convert.ToInt32(itemArray[needExportIndexList[i]]);
                        for (int listIndex = 0; listIndex < realListLength; listIndex++)
                        {
                            keyValueList.Add(new XElement("Value",new XAttribute("value", itemArray[needExportIndexList[i + listIndex + 1]])));
                        }
                        i += listLength;
                    }
                    else if(propertyClassList[i].classType == "StructList")
                    {
                        XElement structList = new XElement(propertyValueNameList[i]);
                        recordNode.Add(structList);
                        string[] paramStringList = propertyClassList[i].classParam.Split(',');
                        int structLength = Convert.ToInt32(paramStringList[0]);
                        int mapLength = Convert.ToInt32(paramStringList[1]);
                        int realMapLength = Convert.ToInt32(itemArray[needExportIndexList[i]]);
                        for (int mapIndex = 0; mapIndex < realMapLength; mapIndex++)
                        {
                            XElement structNode = new XElement(propertyClassList[i].className);
                            structList.Add(structNode);
                            for (int structIndex = 0; structIndex < structLength; structIndex++)
                            {
                                string attrName = propertyValueNameList[i + mapIndex * structLength + structIndex + 1];
                                object attrValue = itemArray[needExportIndexList[i + mapIndex * structLength + structIndex + 1]];
                                XAttribute attribute = new XAttribute(attrName, attrValue);
                                structNode.Add(attribute);
                            }
                        }
                        i += structLength * mapLength;
                    }
                    else if(propertyClassList[i].classType == "StructMap")
                    {
                        XElement structMap = new XElement(propertyValueNameList[i]);
                        recordNode.Add(structMap);
                        string[] paramStringList = propertyClassList[i].classParam.Split(',');
                        int structLength = Convert.ToInt32(paramStringList[0]);
                        int mapLength = Convert.ToInt32(paramStringList[1]);
                        int realMapLength = Convert.ToInt32(itemArray[needExportIndexList[i]]);
                        for (int mapIndex = 0; mapIndex < realMapLength; mapIndex++)
                        {
                            string nodeName = propertyClassList[i].className.Split(',')[1];
                            XElement structNode = new XElement(nodeName);
                            structMap.Add(structNode);
                            for (int structIndex = 0; structIndex < structLength; structIndex++)
                            {
                                string attrName = propertyValueNameList[i + mapIndex * structLength + structIndex + 1];
                                object attrValue = itemArray[needExportIndexList[i + mapIndex * structLength + structIndex + 1]];
                                XAttribute attribute = new XAttribute(attrName, attrValue);
                                structNode.Add(attribute);
                            }
                        }
                        i += structLength* mapLength;
                    }
                    else
                    {
                        recordNode.Add(new XAttribute(propertyValueNameList[i], itemArray[needExportIndexList[i]]));
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
            FileStream fileStream = new FileStream(exportXMLFilePath, FileMode.Create, FileAccess.ReadWrite);
            using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
            {
                doc.Save(xmlWriter);
            }

            //fileLogCallback?.Invoke($"xml文件生成成功：{exportXMLFilePath}");
        }

        /// <summary>
        /// 将XLSX文件导出到CS文件
        /// </summary>
        /// <param name="exportCSFilePath"></param>
        public void ExportCS(string exportCSFilePath)
        {
            //先将文件名替换为配置名称
            string csFileName = ConfigData.GetSingle().CSFileName.Replace("{fileName}", fileName);
            csFileName = csFileName.Replace("{className}", className);
            exportCSFilePath = Path.GetDirectoryName(exportCSFilePath) + "/" + csFileName;
            //fileLogCallback?.Invoke($"xlsx文件开始导出：{xlsxFilePath} -> {exportCSFilePath}");

            FileInfo fileInfo = new FileInfo(exportCSFilePath);
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            FileStream fileStream = new FileStream(exportCSFilePath, FileMode.Create, FileAccess.ReadWrite);
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            {
                StringBuilder csClassContent = new StringBuilder();
                if(xlsxType == XlsxEnum.Recorder)
                {
                    csClassContent.Append(csRecorderTemplateContent);
                }
                else if(xlsxType == XlsxEnum.Enum)
                {
                    csClassContent.Append(csEnumTemplateContent);
                }
                else if (xlsxType == XlsxEnum.Struct)
                {
                    csClassContent.Append(csStructTemplateContent);
                }
                //替换类名
                csClassContent.Replace("{fileName}", fileName);
                csClassContent.Replace("{className}", className);
                //替换命名空间
                string namespaceString = Path.GetRelativePath(ConfigData.GetSingle().ExportCSAbsolutePath, fileInfo.Directory.FullName);
                if(namespaceString!=".")
                {
                    namespaceString = "."+namespaceString.Replace("\\",".");
                }
                else
                {
                    namespaceString = "";
                }
                csClassContent.Replace("{namespace}", namespaceString);
                //替换索引
                if(propertyClassList.Count>0 && propertyValueNameList.Count>0)
                {
                    if(propertyClassList[0].classType=="int")
                    {
                        csClassContent.Replace("{key}", propertyValueNameList[0]);
                    }
                    else
                    {
                        csClassContent.Replace("{key}", "(int)"+propertyValueNameList[0]);
                    }
                }
                if(xlsxType == XlsxEnum.Struct)
                {
                    csClassContent.Replace("{prefix}", prefix);
                    csClassContent.Replace("{suffix}", suffix);
                    csClassContent.Replace("{split}", split);
                }
                //替换属性模板
                Dictionary<string, PropertyTemplateInfoStruct> propertyTemplateMap = ConfigData.GetSingle().CSClassPropertyTemplateMap;
                foreach (var property in propertyTemplateMap)
                {
                    PropertyTemplateInfoStruct propertyTemplateInfoStruct = property.Value;
                    StringBuilder propertyTotalContent = new StringBuilder();
                    bool isFirstAdd = true;
                    for (int i = 0; i < propertyValueNameList.Count; i++)
                    {
                        bool isInAttribute = true;
                        if(propertyClassList.Count>0)
                        {
                            isInAttribute = propertyClassList[i].classType != "KeyValueMap" 
                                && propertyClassList[i].classType != "ValueList"
                                && propertyClassList[i].classType != "StructList"
                                && propertyClassList[i].classType != "StructMap"
                                ;
                        }
                        bool isInElement = !isInAttribute;
                        bool needAdd = false;
                        if (isInAttribute && propertyTemplateInfoStruct.addInAttribute)
                        {
                            needAdd = true;
                        }
                        if(isInElement && propertyTemplateInfoStruct.addInElement)
                        {
                            needAdd = true;
                        }
                        if(needAdd)
                        {
                            //处理换行
                            if (!isFirstAdd)
                            {
                                if(propertyTemplateInfoStruct.split=="{split}")
                                {
                                    propertyTotalContent.Append(split);
                                }
                                else
                                {
                                    propertyTotalContent.Append(propertyTemplateInfoStruct.split);
                                }
                            }
                            isFirstAdd = false;
                            StringBuilder propertyEveryContent = new StringBuilder(propertyTemplateInfoStruct.content);
                            if (propertyClassList.Count > 0)
                            {
                                //根据类型替换转换方法模板
                                propertyEveryContent.Replace("{convertFunction}", GetConvertFunctionByClassType(propertyClassList[i].classType));
                                if (propertyClassList[i].classType == "SplitStringList")
                                {
                                    propertyEveryContent.Replace("{propertyClassParam1}", propertyClassList[i].classParam);
                                    propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassList[i].className).Replace("{propertyClassName}", propertyClassList[i].className));
                                    propertyEveryContent.Replace("{propertyClassName1}", propertyClassList[i].className);
                                    propertyEveryContent.Replace("{propertyClassName}", $"List<{propertyClassList[i].className}>");
                                }
                                else if (propertyClassList[i].classType == "SplitStringMap")
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
                                else if (propertyClassList[i].classType == "ValueList")
                                {
                                    propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassList[i].className).Replace("{propertyClassName}", propertyClassList[i].className));
                                    propertyEveryContent.Replace("{propertyClassName1}", propertyClassList[i].className);
                                    propertyEveryContent.Replace("{propertyClassName}", $"List<{propertyClassList[i].className}>");
                                }
                                else if (propertyClassList[i].classType == "KeyValueMap")
                                {
                                    string[] propertyClassNameList = propertyClassList[i].className.Split(',');
                                    propertyEveryContent.Replace("{propertyClassName1}", propertyClassNameList[0]);
                                    propertyEveryContent.Replace("{propertyClassName2}", propertyClassNameList[1]);
                                    propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassNameList[0]).Replace("{propertyClassName}", propertyClassNameList[0]));
                                    propertyEveryContent.Replace("{convertFunction2}", GetConvertFunctionByClassType(propertyClassNameList[1]).Replace("{propertyClassName}", propertyClassNameList[1]));
                                    propertyEveryContent.Replace("{propertyClassName}", $"Dictionary<{propertyClassList[i].className}>");
                                }
                                else if (propertyClassList[i].classType == "StructList")
                                {
                                    propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassList[i].className).Replace("{propertyClassName}", propertyClassList[i].className));
                                    propertyEveryContent.Replace("{propertyClassName1}", propertyClassList[i].className);
                                    propertyEveryContent.Replace("{propertyClassName}", $"List<{propertyClassList[i].className}>");
                                }
                                else if (propertyClassList[i].classType == "StructMap")
                                {
                                    string[] propertyClassNameList = propertyClassList[i].className.Split(',');
                                    propertyEveryContent.Replace("{structMapKeyName}", propertyValueNameList[i+1]);
                                    propertyEveryContent.Replace("{propertyClassName1}", propertyClassNameList[0]);
                                    propertyEveryContent.Replace("{propertyClassName2}", propertyClassNameList[1]);
                                    propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(propertyClassNameList[0]).Replace("{propertyClassName}", propertyClassNameList[0]));
                                    propertyEveryContent.Replace("{convertFunction2}", GetConvertFunctionByClassType(propertyClassNameList[1]).Replace("{propertyClassName}", propertyClassNameList[1]));
                                    propertyEveryContent.Replace("{propertyClassName}", $"Dictionary<{propertyClassList[i].className}>");
                                }
                                else
                                {
                                    propertyEveryContent.Replace("{propertyClassName}", propertyClassList[i].className);
                                }
                            }
                            if (propertyConfigNameList.Count > 0)
                            {
                                propertyEveryContent.Replace("{propertyConfigName}", propertyConfigNameList[i]);
                            }
                            if (propertyDescriptionList.Count > 0)
                            {
                                propertyEveryContent.Replace("{propertyDescription}", propertyDescriptionList[i]);
                            }
                            if (propertyValueNameList.Count > 0)
                            {
                                propertyEveryContent.Replace("{propertyValueName}", propertyValueNameList[i]);
                            }
                            if(xlsxType == XlsxEnum.Struct)
                            {
                                propertyEveryContent.Replace("{propertyDefaultValue}", propertyDefaultValueList[i]);
                                propertyEveryContent.Replace("{propertyIndex}", i.ToString());
                            }
                            propertyTotalContent.Append(propertyEveryContent.ToString());
                        }
                        if (propertyClassList.Count>0 && (propertyClassList[i].classType == "KeyValueMap"|| propertyClassList[i].classType == "ValueList"))
                        {
                            int mapLength = Convert.ToInt32(propertyClassList[i].classParam);
                            i += mapLength;
                        }
                        else if(propertyClassList.Count > 0 && (propertyClassList[i].classType == "StructList" || propertyClassList[i].classType == "StructMap"))
                        {
                            string[] paramStringList = propertyClassList[i].classParam.Split(',');
                            int structLength = Convert.ToInt32(paramStringList[0]);
                            int mapLength = Convert.ToInt32(paramStringList[1]);
                            i += mapLength * structLength;
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
