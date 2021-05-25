using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace XlsxToXmlDll
{
    /// <summary>
    /// xlsx文件类
    /// </summary>
    class XlsxFile
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

        /// <summary>
        /// 导出信息结构体
        /// </summary>
        class ExportInfo
        {
            public int index = 0;
            public string originString = "";
            public List<string> exportCodeNameList = new List<string>();
        }

        /// <summary>
        /// Xlsx属性类型
        /// </summary>
        class XlsxPropertyClass
        {
            /// <summary>
            /// 原本内容
            /// </summary>
            public string originString;
            /// <summary>
            /// 属性类型
            /// </summary>
            public string classType;
            /// <summary>
            /// 属性类名
            /// </summary>
            public string className;
            /// <summary>
            /// 子属性类型，一般用于ListMap等泛型类
            /// </summary>
            public List<XlsxPropertyClass> childXlsxPropertyClassList = new List<XlsxPropertyClass>();
            /// <summary>
            /// 属性参数，因为是字符串，所以可以当做List使用
            /// </summary>
            public string classParameter = "";

            /// <summary>
            /// 从字符串中读取
            /// </summary>
            /// <param name="stringValue"></param>
            public void LoadFormString(XlsxFile xlsxFile, string stringValue)
            {
                originString = stringValue;
                if (stringValue.Contains(' '))
                {
                    string[] stringValueList = stringValue.Split(' ');
                    classType = stringValueList[0];
                    if (stringValueList.Length > 1)
                    {
                        string[] propertyChildClassList = stringValueList[1].Split(',');
                        foreach (var propertyChildClass in propertyChildClassList)
                        {
                            XlsxPropertyClass childXlsxPropertyClass = new XlsxPropertyClass();
                            childXlsxPropertyClass.LoadFormString(xlsxFile, propertyChildClass);
                            childXlsxPropertyClassList.Add(childXlsxPropertyClass);
                        }
                    }
                    if (stringValueList.Length > 2)
                    {
                        classParameter = stringValueList[2];
                    }
                }
                else
                {
                    classType = stringValue;
                }
            }

            /// <summary>
            /// 刷新信息
            /// </summary>
            public void RefreshInfo(XlsxFile xlsxFile, bool isClassValue)
            {
                className = xlsxFile.GetClassNameByClassType(classType, isClassValue);
                foreach (var childXlsxPropertyClass in childXlsxPropertyClassList)
                {
                    childXlsxPropertyClass.RefreshInfo(xlsxFile, true);
                }
            }
        }

        string xlsxFilePath = "";
        string fileName = "";
        string className = "";
        string nameSpaceRelativeName = "";
        XlsxEnum xlsxType = XlsxEnum.Recorder;
        DataRowCollection xlsxDataRowCollection = null;
        /// <summary>
        /// 需要导出的列
        /// </summary>
        List<ExportInfo> needExportInfoList = new List<ExportInfo>();
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

        CodeConfigData codeConfigData = null;

        public XlsxFile(string xlsxFilePath)
        {
            this.xlsxFilePath = xlsxFilePath;
            ReadExcel();
        }

        /// <summary>
        /// 读取xlsx文件，只读取第一页(sheet1)的数据
        /// </summary>
        /// <returns></returns>
        void ReadExcel()
        {
            FileInfo xlsxFileInfo = new FileInfo(xlsxFilePath);
            fileName = xlsxFileInfo.Name.Substring(0, xlsxFileInfo.Name.LastIndexOf('.'));
            if (fileName.Contains('.'))
            {
                int pointIndex = fileName.LastIndexOf('.');
                className = fileName.Substring(0, pointIndex);
                xlsxType = XlsxEnum.Recorder;
            }
            else if (fileName.EndsWith("Recorder"))
            {
                xlsxType = XlsxEnum.Recorder;
                className = fileName;
            }
            else if (fileName.EndsWith("Enum"))
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
                XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}，只能使用Recorder、Enum、Struct结尾，代表配置、枚举、结构体！");
            }
            string namespaceString = XlsxManager.GetRelativePath(ConfigData.GetSingle().ImportXlsxAbsolutePath, xlsxFileInfo.Directory.FullName);
            if (namespaceString != "" && namespaceString != ".")
            {
                namespaceString = "." + namespaceString.Replace("\\", ".");
            }
            else
            {
                namespaceString = "";
            }
            nameSpaceRelativeName = namespaceString;
            FileStream stream = File.Open(xlsxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            xlsxDataRowCollection = result.Tables[0].Rows;

            int rowCount = xlsxDataRowCollection.Count;
            int colCount = xlsxDataRowCollection[0].ItemArray.Length;

            if (xlsxType == XlsxEnum.Recorder)
            {
                if (rowCount < 5)
                {
                    XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}中行数为{rowCount}，小于5，没有正确定义");
                    return;
                }
                if (colCount > 100)
                {
                    XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}中列数太多，为{colCount}，超过100，请检查。如果需要扩充请修改代码。");
                    return;
                }
                //因为在xlsx配置中有可能出现空内容等问题，使属性列比实际要使用的列数多
                //所以先算出实际使用的列个数
                //xlsx文件格式：第一行，获得需要导出的列
                for (int i = 0; i < colCount; i++)
                {
                    object item = xlsxDataRowCollection[0][i];
                    if (item != DBNull.Value && item.ToString() != "")
                    {
                        ExportInfo exportInfo = new ExportInfo();
                        exportInfo.originString = item.ToString();
                        exportInfo.index = i;
                        exportInfo.exportCodeNameList.AddRange(item.ToString().Split(';'));
                        needExportInfoList.Add(exportInfo);
                    }
                }
                foreach (ExportInfo exportInfo in needExportInfoList)
                {
                    //xlsx文件格式：第二行，为属性名称
                    propertyValueNameList.Add(xlsxDataRowCollection[1][exportInfo.index].ToString());

                    //xlsx文件格式：第三行，为类型名称
                    string propertyClassString = xlsxDataRowCollection[2][exportInfo.index].ToString();
                    XlsxPropertyClass propertyClass = new XlsxPropertyClass();
                    propertyClass.LoadFormString(this, propertyClassString);

                    propertyClassList.Add(propertyClass);

                    //xlsx文件格式：第四行，为规则描述，一般为空
                    propertyDescriptionList.Add(xlsxDataRowCollection[3][exportInfo.index].ToString());

                    //xlsx文件格式：第五行，为配置名称，作为属性名称的注释
                    propertyConfigNameList.Add(xlsxDataRowCollection[4][exportInfo.index].ToString());
                }
            }
            else if (xlsxType == XlsxEnum.Enum)
            {
                if (colCount < 2)
                {
                    XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}中列数小于2，为{colCount}，请检查。需要保证一列名称，一列含义。");
                    return;
                }
                if (rowCount < 2)
                {
                    XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}中行数小于2，为{rowCount}，请检查。需要保证至少一行值。");
                    return;
                }
                for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
                {
                    object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                    if (itemArray.Length < 1 || itemArray[0].ToString() == "")
                    {
                        XlsxManager.Log(false, $"xlsx文件: {xlsxFilePath}表内容出现空值，可能是有意为之或内容未清除彻底，总行数:{rowCount},请核实row:{rowIndex + 1}：");
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
                    XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}中列数小于4，为{colCount}，请检查。需要保证结构正确。");
                    return;
                }
                if (rowCount < 5)
                {
                    XlsxManager.Log(false, $"xlsx文件：{xlsxFilePath}中行数小于5，为{rowCount}，请检查。需要保证至少一行值。");
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
                        XlsxManager.Log(false, $"xlsx文件: {xlsxFilePath}表内容出现空值，可能是有意为之或内容未清除彻底，总行数:{rowCount},请核实row:{rowIndex + 1}：");
                        break;
                    }
                    propertyValueNameList.Add(itemArray[0].ToString());
                    propertyConfigNameList.Add(itemArray[1].ToString());
                    XlsxPropertyClass propertyClass = new XlsxPropertyClass();
                    string propertyClassString = itemArray[2].ToString();
                    propertyClass.LoadFormString(this, propertyClassString);
                    propertyClassList.Add(propertyClass);
                    propertyDefaultValueList.Add(itemArray[3].ToString().ToString());
                }
            }
        }

        /// <summary>
        /// 将XLSX文件导出
        /// </summary>
        /// <param name="codeName"></param>
        /// <param name="xlsxFileRelativePath"></param>
        public void Export(string codeName, string xlsxFileRelativePath)
        {
            codeConfigData = ConfigData.GetSingle().CodeConfigDataMap[codeName];
            foreach (var propertyClass in propertyClassList)
            {
                propertyClass.RefreshInfo(this, false);
            }
            //先判断当前配置的是否需要导出当前语言
            if (xlsxType == XlsxEnum.Recorder)
            {
                bool needExport = false;
                foreach (var item in needExportInfoList)
                {
                    if (item.exportCodeNameList.Contains(codeName))
                    {
                        needExport = true;
                        break;
                    }
                }
                //如果不需要导出，则删除原本应该生成的文件
                if(!needExport)
                {
                    DeleteXml(xlsxFileRelativePath);
                    DeleteCode(xlsxFileRelativePath);
                    return;
                }
            }
            ExportXML(xlsxFileRelativePath);
            ExportCode(xlsxFileRelativePath);
        }

        string GetXmlPathByXlsxRelativePath(string xlsxFileRelativePath)
        {
            string exportXmlFilePath = "";
            if (codeConfigData.ExportAllXmlToTheSamePath)
            {
                exportXmlFilePath = codeConfigData.ExportXmlAbsolutePath;
                //先将文件名替换为配置名称
                string xmlFileName = codeConfigData.XmlFileName.Replace("{fileName}", fileName);
                xmlFileName = xmlFileName.Replace("{className}", className);
                xmlFileName = xmlFileName.Replace("{nameSpaceRelativeName}", nameSpaceRelativeName);
                exportXmlFilePath = Path.GetDirectoryName(exportXmlFilePath) + "/" + xmlFileName;
            }
            else
            {
                exportXmlFilePath = codeConfigData.ExportXmlAbsolutePath + "/" + xlsxFileRelativePath;
                exportXmlFilePath = exportXmlFilePath.Substring(0, exportXmlFilePath.LastIndexOf('.')) + ".xml";
                //先将文件名替换为配置名称
                string xmlFileName = codeConfigData.XmlFileName.Replace("{fileName}", fileName);
                xmlFileName = xmlFileName.Replace("{className}", className);
                xmlFileName = xmlFileName.Replace("{nameSpaceRelativeName}", nameSpaceRelativeName);
                exportXmlFilePath = Path.GetDirectoryName(exportXmlFilePath) + "/" + xmlFileName;
            }
            return exportXmlFilePath;
        }

        string GetCodePathByXlsxRelativePath(string xlsxFileRelativePath)
        {
            string exportCodeFilePath = codeConfigData.ExportCodeAbsolutePath + "/" + xlsxFileRelativePath;
            //先将文件名替换为配置名称
            string codeFileName = codeConfigData.CodeFileName.Replace("{fileName}", fileName);
            codeFileName = codeFileName.Replace("{className}", className);
            exportCodeFilePath = Path.GetDirectoryName(exportCodeFilePath) + "/" + codeFileName;
            return exportCodeFilePath;
        }

        void DeleteXml(string xlsxFileRelativePath)
        {
            if (xlsxType != XlsxEnum.Recorder)
            {
                return;
            }
            string exportXmlFilePath = GetXmlPathByXlsxRelativePath(xlsxFileRelativePath);
            if (File.Exists(exportXmlFilePath))
            {
                File.Delete(exportXmlFilePath);
            }
        }

        void DeleteCode(string xlsxFileRelativePath)
        {
            string exportCodeFilePath = GetCodePathByXlsxRelativePath(xlsxFileRelativePath);
            if (File.Exists(exportCodeFilePath))
            {
                File.Delete(exportCodeFilePath);
            }
        }

        /// <summary>
        /// 将XLSX文件导出到XML文件
        /// </summary>
        /// <param name="exportXmlFilePath"></param>
        void ExportXML(string xlsxFileRelativePath)
        {
            if (xlsxType != XlsxEnum.Recorder)
            {
                return;
            }
            string exportXmlFilePath = GetXmlPathByXlsxRelativePath(xlsxFileRelativePath);

            int rowCount = xlsxDataRowCollection.Count;
            string xmlRootName = codeConfigData.XmlRootNodeName;
            xmlRootName = xmlRootName.Replace("{fileName}", fileName);
            xmlRootName = xmlRootName.Replace("{className}", className);
            xmlRootName = xmlRootName.Replace("{nameSpaceRelativeName}", nameSpaceRelativeName);
            XDocument doc = new XDocument(new XElement(xmlRootName));
            string recorderNodeName = codeConfigData.XmlRecorderNodeName;
            recorderNodeName = recorderNodeName.Replace("{fileName}", fileName);
            recorderNodeName = recorderNodeName.Replace("{className}", className);
            recorderNodeName = recorderNodeName.Replace("{nameSpaceRelativeName}", nameSpaceRelativeName);
            for (int rowIndex = 5; rowIndex < rowCount; rowIndex++)
            {
                object[] itemArray = xlsxDataRowCollection[rowIndex].ItemArray;
                if (itemArray.Length < 1 || itemArray[0].ToString() == "")
                {
                    XlsxManager.Log(false, $"xml文件{exportXmlFilePath}生成成功，但表内容出现空值，可能是有意为之或内容未清除彻底，总行数:{rowCount},请核实row:{rowIndex + 1}：");
                    break;
                }
                XElement recordNode = new XElement(recorderNodeName);
                for (int i = 0; i < needExportInfoList.Count; i++)
                {
                    if (!needExportInfoList[i].exportCodeNameList.Contains(codeConfigData.CodeName))
                    {
                        continue;
                    }
                    if (propertyClassList[i].classType == "ValueList")
                    {
                        XElement keyValueList = new XElement(propertyValueNameList[i]);
                        recordNode.Add(keyValueList);
                        int listLength = Convert.ToInt32(propertyClassList[i].classParameter);
                        int realListLength = Convert.ToInt32(itemArray[needExportInfoList[i].index]);
                        for (int listIndex = 0; listIndex < realListLength; listIndex++)
                        {
                            CheckValueTypeByIndex(rowIndex, i + listIndex + 1);
                            XElement valueElement = new XElement("Value");
                            valueElement.Add(new XAttribute("value", itemArray[needExportInfoList[i + listIndex + 1].index]));
                            keyValueList.Add(valueElement);
                        }
                        i += listLength;
                    }
                    else if (propertyClassList[i].classType == "KeyValueMap")
                    {
                        XElement keyValueMap = new XElement(propertyValueNameList[i]);
                        recordNode.Add(keyValueMap);
                        int mapLength = Convert.ToInt32(propertyClassList[i].classParameter);
                        for (int mapIndex = 0; mapIndex < mapLength; mapIndex++)
                        {
                            CheckValueTypeByIndex(rowIndex, i + mapIndex + 1);
                            XElement keyValueElement = new XElement("KeyValue");
                            keyValueElement.Add(new XAttribute("key", propertyValueNameList[i + mapIndex + 1]));
                            keyValueElement.Add(new XAttribute("value", itemArray[needExportInfoList[i + mapIndex + 1].index]));
                            keyValueMap.Add(keyValueElement);
                        }
                        i += mapLength;
                    }
                    else if (propertyClassList[i].classType == "StructList")
                    {
                        XElement structList = new XElement(propertyValueNameList[i]);
                        recordNode.Add(structList);
                        string[] paramStringList = propertyClassList[i].classParameter.Split(',');
                        int structLength = Convert.ToInt32(paramStringList[0]);
                        int mapLength = Convert.ToInt32(paramStringList[1]);
                        int realMapLength = Convert.ToInt32(itemArray[needExportInfoList[i].index]);
                        for (int mapIndex = 0; mapIndex < realMapLength; mapIndex++)
                        {
                            XElement structNode = new XElement(propertyClassList[i].childXlsxPropertyClassList[0].className);
                            structList.Add(structNode);
                            for (int structIndex = 0; structIndex < structLength; structIndex++)
                            {
                                string attrName = propertyValueNameList[i + mapIndex * structLength + structIndex + 1];
                                object attrValue = itemArray[needExportInfoList[i + mapIndex * structLength + structIndex + 1].index];
                                CheckValueTypeByIndex(rowIndex, i + mapIndex * structLength + structIndex + 1);
                                if (codeConfigData.XmlAttributeFirst)
                                {
                                    XAttribute attribute = new XAttribute(attrName, attrValue);
                                    structNode.Add(attribute);
                                }
                                else
                                {
                                    XElement element = new XElement(attrName, attrValue);
                                    structNode.Add(element);
                                }
                            }
                        }
                        i += structLength * mapLength;
                    }
                    else if (propertyClassList[i].classType == "StructMap")
                    {
                        XElement structMap = new XElement(propertyValueNameList[i]);
                        recordNode.Add(structMap);
                        string[] paramStringList = propertyClassList[i].classParameter.Split(',');
                        int structLength = Convert.ToInt32(paramStringList[0]);
                        int mapLength = Convert.ToInt32(paramStringList[1]);
                        int realMapLength = Convert.ToInt32(itemArray[needExportInfoList[i].index]);
                        for (int mapIndex = 0; mapIndex < realMapLength; mapIndex++)
                        {
                            XElement structNode = new XElement(propertyClassList[i].childXlsxPropertyClassList[1].className);
                            structMap.Add(structNode);
                            for (int structIndex = 0; structIndex < structLength; structIndex++)
                            {
                                string attrName = propertyValueNameList[i + mapIndex * structLength + structIndex + 1];
                                object attrValue = itemArray[needExportInfoList[i + mapIndex * structLength + structIndex + 1].index];
                                CheckValueTypeByIndex(rowIndex, i + mapIndex * structLength + structIndex + 1);
                                if (codeConfigData.XmlAttributeFirst)
                                {
                                    XAttribute attribute = new XAttribute(attrName, attrValue);
                                    structNode.Add(attribute);
                                }
                                else
                                {
                                    XElement element = new XElement(attrName, attrValue);
                                    structNode.Add(element);
                                }
                            }
                        }
                        i += structLength * mapLength;
                    }
                    else
                    {
                        CheckValueTypeByIndex(rowIndex, i);
                        if (codeConfigData.XmlAttributeFirst)
                        {
                            recordNode.Add(new XAttribute(propertyValueNameList[i], itemArray[needExportInfoList[i].index]));
                        }
                        else
                        {
                            recordNode.Add(new XElement(propertyValueNameList[i], itemArray[needExportInfoList[i].index]));
                        }
                    }
                }
                doc.Root.Add(recordNode);
            }

            //保存时忽略声明
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;
            FileInfo fileInfo = new FileInfo(exportXmlFilePath);
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            using (FileStream fileStream = new FileStream(exportXmlFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
                {
                    doc.Save(xmlWriter);
                }
            }
        }

        /// <summary>
        /// 按代码名称导出代码
        /// </summary>
        /// <param name="codeName"></param>
        /// <param name="xlsxFileRelativePath"></param>
        void ExportCode(string xlsxFileRelativePath)
        {
            string exportCodeFilePath = GetCodePathByXlsxRelativePath(xlsxFileRelativePath);

            FileInfo fileInfo = new FileInfo(exportCodeFilePath);
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
            using (FileStream fileStream = new FileStream(exportCodeFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                using (StreamWriter streamWriter = new StreamWriter(fileStream))
                {
                    StringBuilder csClassContent = new StringBuilder();
                    if (xlsxType == XlsxEnum.Recorder)
                    {
                        csClassContent.Append(codeConfigData.RecorderTemplateContent);
                    }
                    else if (xlsxType == XlsxEnum.Enum)
                    {
                        csClassContent.Append(codeConfigData.EnumTemplateContent);
                    }
                    else if (xlsxType == XlsxEnum.Struct)
                    {
                        csClassContent.Append(codeConfigData.StructTemplateContent);
                    }
                    //替换类名
                    csClassContent.Replace("{fileName}", fileName);
                    csClassContent.Replace("{className}", className);
                    //替换命名空间根名称
                    csClassContent.Replace("{nameSpaceRootName}", codeConfigData.NameSpaceRootName);
                    //替换命名空间
                    csClassContent.Replace("{nameSpaceRelativeName}", nameSpaceRelativeName);
                    //替换索引
                    if (propertyClassList.Count > 0 && propertyValueNameList.Count > 0)
                    {
                        if (propertyClassList[0].classType != "int")
                        {
                            csClassContent.Replace("{key}", codeConfigData.RecorderKeyUnknownClassTypeToInt);
                        }
                        csClassContent.Replace("{key}", propertyValueNameList[0]);
                    }
                    if (xlsxType == XlsxEnum.Struct)
                    {
                        csClassContent.Replace("{prefix}", prefix);
                        csClassContent.Replace("{suffix}", suffix);
                        csClassContent.Replace("{split}", split);
                    }
                    //替换属性模板
                    Dictionary<string, PropertyTemplateInfoStruct> propertyTemplateMap = codeConfigData.ClassPropertyTemplateMap;
                    foreach (var property in propertyTemplateMap)
                    {
                        PropertyTemplateInfoStruct propertyTemplateInfoStruct = property.Value;
                        StringBuilder propertyTotalContent = new StringBuilder();
                        bool isFirstAdd = true;
                        for (int i = 0; i < propertyValueNameList.Count; i++)
                        {
                            if (xlsxType == XlsxEnum.Recorder)
                            {
                                if (!needExportInfoList[i].exportCodeNameList.Contains(codeConfigData.CodeName))
                                {
                                    continue;
                                }
                            }
                            if (xlsxType != XlsxEnum.Enum)
                            {
                                //先判断是否有设置特殊类型
                                if ((IsSpecialByClassType(propertyClassList[i].classType) && property.Value.addIsSpecialClassType)
                                    || (!IsSpecialByClassType(propertyClassList[i].classType) && property.Value.addIsNotSpecialClassType))
                                {
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            bool isInAttribute = false;
                            if (codeConfigData.XmlAttributeFirst)
                            {
                                if (propertyClassList.Count > 0)
                                {
                                    isInAttribute = propertyClassList[i].classType != "KeyValueMap"
                                        && propertyClassList[i].classType != "ValueList"
                                        && propertyClassList[i].classType != "StructList"
                                        && propertyClassList[i].classType != "StructMap"
                                        ;
                                }
                            }
                            bool isInElement = !isInAttribute;
                            bool needAdd = false;
                            if (isInAttribute && propertyTemplateInfoStruct.addInAttribute)
                            {
                                needAdd = true;
                            }
                            if (isInElement && propertyTemplateInfoStruct.addInElement)
                            {
                                needAdd = true;
                            }
                            if (needAdd)
                            {
                                //处理换行
                                if (!isFirstAdd)
                                {
                                    if (propertyTemplateInfoStruct.split == "{split}")
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
                                    propertyEveryContent.Replace("{toStringFunction}", GetStringFunctionByClassType(propertyClassList[i].classType));
                                    propertyEveryContent.Replace("{split}", split);
                                    List<XlsxPropertyClass> childXlsxPropertyClassList = propertyClassList[i].childXlsxPropertyClassList;
                                    if (propertyClassList[i].classType == "SplitStringList")
                                    {
                                        propertyEveryContent.Replace("{propertyClassParam1}", propertyClassList[i].classParameter);
                                        propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(childXlsxPropertyClassList[0].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[0].className));
                                        propertyEveryContent.Replace("{propertyClassName1}", childXlsxPropertyClassList[0].className);
                                        propertyEveryContent.Replace("{propertyClassName}", $"{propertyClassList[i].className}<{childXlsxPropertyClassList[0].className}>");
                                    }
                                    else if (propertyClassList[i].classType == "SplitStringMap")
                                    {
                                        propertyEveryContent.Replace("{propertyClassName1}", childXlsxPropertyClassList[0].className);
                                        propertyEveryContent.Replace("{propertyClassName2}", childXlsxPropertyClassList[1].className);
                                        propertyEveryContent.Replace("{propertyClassParam1}", propertyClassList[i].classParameter[0].ToString());
                                        propertyEveryContent.Replace("{propertyClassParam2}", propertyClassList[i].classParameter[1].ToString());
                                        propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(childXlsxPropertyClassList[0].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[0].className));
                                        propertyEveryContent.Replace("{convertFunction2}", GetConvertFunctionByClassType(childXlsxPropertyClassList[1].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[1].className));
                                        propertyEveryContent.Replace("{propertyClassName}", $"{propertyClassList[i].className}<{childXlsxPropertyClassList[0].className},{childXlsxPropertyClassList[1].className}>");
                                    }
                                    else if (propertyClassList[i].classType == "ValueList")
                                    {
                                        propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(childXlsxPropertyClassList[0].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[0].className));
                                        propertyEveryContent.Replace("{propertyClassName1}", childXlsxPropertyClassList[0].className);
                                        propertyEveryContent.Replace("{propertyClassName}", $"{propertyClassList[i].className}<{childXlsxPropertyClassList[0].className}>");
                                    }
                                    else if (propertyClassList[i].classType == "KeyValueMap")
                                    {
                                        string[] propertyClassNameList = propertyClassList[i].className.Split(',');
                                        propertyEveryContent.Replace("{propertyClassName1}", childXlsxPropertyClassList[0].className);
                                        propertyEveryContent.Replace("{propertyClassName2}", childXlsxPropertyClassList[1].className);
                                        propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(childXlsxPropertyClassList[0].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[0].className));
                                        propertyEveryContent.Replace("{convertFunction2}", GetConvertFunctionByClassType(childXlsxPropertyClassList[1].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[1].className));
                                        propertyEveryContent.Replace("{propertyClassName}", $"{propertyClassList[i].className}<{childXlsxPropertyClassList[0].className},{childXlsxPropertyClassList[1].className}>");
                                    }
                                    else if (propertyClassList[i].classType == "StructList")
                                    {
                                        propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(childXlsxPropertyClassList[0].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[0].className));
                                        propertyEveryContent.Replace("{propertyClassName1}", childXlsxPropertyClassList[0].className);
                                        propertyEveryContent.Replace("{propertyClassName}", $"{propertyClassList[i].className}<{childXlsxPropertyClassList[0].className}>");
                                    }
                                    else if (propertyClassList[i].classType == "StructMap")
                                    {
                                        string[] propertyClassNameList = propertyClassList[i].className.Split(',');
                                        propertyEveryContent.Replace("{structMapKeyName}", propertyValueNameList[i + 1]);
                                        propertyEveryContent.Replace("{propertyClassName1}", childXlsxPropertyClassList[0].className);
                                        propertyEveryContent.Replace("{propertyClassName2}", childXlsxPropertyClassList[1].className);
                                        propertyEveryContent.Replace("{convertFunction1}", GetConvertFunctionByClassType(childXlsxPropertyClassList[0].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[0].className));
                                        propertyEveryContent.Replace("{convertFunction2}", GetConvertFunctionByClassType(childXlsxPropertyClassList[1].classType).Replace("{propertyClassName}", childXlsxPropertyClassList[1].className));
                                        propertyEveryContent.Replace("{propertyClassName}", $"{propertyClassList[i].className}<{childXlsxPropertyClassList[0].className},{childXlsxPropertyClassList[1].className}>");
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
                                if (xlsxType == XlsxEnum.Struct)
                                {
                                    propertyEveryContent.Replace("{propertyDefaultValue}", propertyDefaultValueList[i]);
                                    propertyEveryContent.Replace("{propertyIndex}", i.ToString());
                                }
                                propertyTotalContent.Append(propertyEveryContent.ToString());
                            }
                            //跳过占用多个列数的列
                            if (propertyClassList.Count > 0 && (propertyClassList[i].classType == "KeyValueMap" || propertyClassList[i].classType == "ValueList"))
                            {
                                int mapLength = Convert.ToInt32(propertyClassList[i].classParameter);
                                i += mapLength;
                            }
                            else if (propertyClassList.Count > 0 && (propertyClassList[i].classType == "StructList" || propertyClassList[i].classType == "StructMap"))
                            {
                                string[] paramStringList = propertyClassList[i].classParameter.Split(',');
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
        }

        /// <summary>
        /// 导出配置总览
        /// </summary>
        /// <returns></returns>
        public XElement ExportRecorderOverview()
        {
            XElement node = new XElement(fileName);
            node.Add(new XAttribute("className", className));
            node.Add(new XAttribute("xlsxType", xlsxType));
            if (xlsxType == XlsxEnum.Recorder)
            {
                for (int i = 0; i < needExportInfoList.Count; i++)
                {
                    XElement itemNode = new XElement(propertyValueNameList[i]);
                    itemNode.Add(new XAttribute("configName", propertyConfigNameList[i]));
                    itemNode.Add(new XAttribute("class", propertyClassList[i].originString));
                    itemNode.Add(new XAttribute("needExport", needExportInfoList[i].originString));
                    node.Add(itemNode);
                    //因为存在特殊list和map类型，需要判断跳过
                    if (propertyClassList[i].classType == "ValueList")
                    {
                        int listLength = Convert.ToInt32(propertyClassList[i].classParameter);
                        i += listLength;
                    }
                    else if (propertyClassList[i].classType == "KeyValueMap")
                    {
                        int mapLength = Convert.ToInt32(propertyClassList[i].classParameter);
                        i += mapLength;
                    }
                    else if (propertyClassList[i].classType == "StructList")
                    {
                        string[] paramStringList = propertyClassList[i].classParameter.Split(',');
                        int structLength = Convert.ToInt32(paramStringList[0]);
                        int mapLength = Convert.ToInt32(paramStringList[1]);
                        i += structLength * mapLength;
                    }
                    else if (propertyClassList[i].classType == "StructMap")
                    {
                        string[] paramStringList = propertyClassList[i].classParameter.Split(',');
                        int structLength = Convert.ToInt32(paramStringList[0]);
                        int mapLength = Convert.ToInt32(paramStringList[1]);
                        i += structLength * mapLength;
                    }
                }
            }
            else if (xlsxType == XlsxEnum.Struct)
            {
                node.Add(new XAttribute("prefix", prefix));
                node.Add(new XAttribute("suffix", suffix));
                node.Add(new XAttribute("split", split));
                for (int i = 0; i < propertyValueNameList.Count; i++)
                {
                    XElement itemNode = new XElement(propertyValueNameList[i]);
                    itemNode.Add(new XAttribute("configName", propertyConfigNameList[i]));
                    itemNode.Add(new XAttribute("class", propertyClassList[i].originString));
                    itemNode.Add(new XAttribute("defaultValue", propertyDefaultValueList[i]));
                    node.Add(itemNode);
                }
            }
            else if (xlsxType == XlsxEnum.Enum)
            {
                for (int i = 0; i < propertyValueNameList.Count; i++)
                {
                    node.Add(new XElement(propertyValueNameList[i], propertyConfigNameList[i]));
                }
            }
            return node;
        }

        /// <summary>
        /// 判断类型是否是特殊类型
        /// </summary>
        /// <param name="classType"></param>
        /// <returns></returns>
        bool IsSpecialByClassType(string classType)
        {
            if (codeConfigData.ClassTypeInfoMap.ContainsKey(classType))
            {
                return codeConfigData.ClassTypeInfoMap[classType].isSpecial;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 通过类型获得类型名称
        /// </summary>
        /// <param name="classType"></param>
        string GetClassNameByClassType(string classType, bool inClassValue = false)
        {
            if (codeConfigData.ClassTypeInfoMap.ContainsKey(classType))
            {
                if (!inClassValue)
                {
                    return codeConfigData.ClassTypeInfoMap[classType].normalValue;
                }
                else
                {
                    return codeConfigData.ClassTypeInfoMap[classType].classValue;
                }
            }
            else
            {
                return codeConfigData.UnknownClassTypeNamePrefix + classType;
            }
        }

        /// <summary>
        /// 通过类型获得转换方法
        /// </summary>
        /// <param name="classType"></param>
        string GetConvertFunctionByClassType(string classType)
        {
            if (codeConfigData.ConvertFunctionTemplateMap.ContainsKey(classType))
            {
                return codeConfigData.ConvertFunctionTemplateMap[classType];
            }
            else
            {
                if (codeConfigData.ConvertFunctionTemplateMap.ContainsKey("custom"))
                {
                    return codeConfigData.ConvertFunctionTemplateMap["custom"];
                }
                return "";
            }
        }

        /// <summary>
        /// 通过类型获得字符串方法
        /// </summary>
        /// <param name="classType"></param>
        string GetStringFunctionByClassType(string classType)
        {
            if (codeConfigData.ToStringFunctionTemplateMap.ContainsKey(classType))
            {
                return codeConfigData.ToStringFunctionTemplateMap[classType];
            }
            else
            {
                if (codeConfigData.ToStringFunctionTemplateMap.ContainsKey("custom"))
                {
                    return codeConfigData.ToStringFunctionTemplateMap["custom"];
                }
                return "";
            }
        }

        /// <summary>
        /// 检测指定位置的值是否符合类型，失败直接抛异常
        /// </summary>
        /// <returns></returns>
        void CheckValueTypeByIndex(int row, int col)
        {
            XlsxPropertyClass xlsxPropertyClass = propertyClassList[col];
            object value = xlsxDataRowCollection[row][needExportInfoList[col].index];
            try
            {
                if (xlsxPropertyClass.classType == "int")
                {
                    Convert.ToInt32(value);
                }
                else if (xlsxPropertyClass.classType == "float")
                {
                    Convert.ToDouble(value);
                }
                else if (xlsxPropertyClass.classType == "bool")
                {
                    Convert.ToBoolean(value);
                }
            }
            catch
            {
                throw new CustomException($"配置:{xlsxFilePath}的第{row}行，第{col}列，名称:{propertyValueNameList[col]}，类型:{xlsxPropertyClass.classType}，值:{value}，类型检测失败!");
            }
        }
    }
}
