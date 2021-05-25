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
    /// <summary>
    /// 类型名称结构体，主要是因为java使用Map等泛型时，int需要改成Integer，所以使用配置配具体类名
    /// </summary>
    struct ClassTypeInfoStruct
    {
        /// <summary>
        /// 普通类型值
        /// </summary>
        public string normalValue;
        /// <summary>
        /// 类类型值
        /// </summary>
        public string classValue;
        /// <summary>
        /// 是否是特殊类型
        /// </summary>
        public bool isSpecial;
    }

    /// <summary>
    /// 属性模板结构体
    /// </summary>
    struct PropertyTemplateInfoStruct
    {
        /// <summary>
        /// 是否添加到attribute属性部分
        /// </summary>
        public bool addInAttribute;
        /// <summary>
        /// 是否添加到element节点部分
        /// </summary>
        public bool addInElement;
        /// <summary>
        /// 是否添加特殊类型
        /// </summary>
        public bool addIsSpecialClassType;
        /// <summary>
        /// 是否添加非特殊类型
        /// </summary>
        public bool addIsNotSpecialClassType;
        /// <summary>
        /// 内容
        /// </summary>
        public string content;
        /// <summary>
        /// 分割符
        /// </summary>
        public string split;
    }

    /// <summary>
    /// 代码配置类
    /// </summary>
    class CodeConfigData
    {
        public string CodeName { get; set; } = "";
        public bool NeedExport { get; set; } = true;
        public string XmlFileName { get; private set; } = "Recorder.xml";
        public bool XmlAttributeFirst { get; private set; } = true;
        public string XmlRootNodeName { get; private set; } = "{fileName}";
        public string XmlRecorderNodeName { get; private set; } = "Recorder";
        public bool ExportAllXmlToTheSamePath { get; private set; } = false;
        public string ExportXmlRelativePath { get; set; } = "/../";
        public string ExportXmlAbsolutePath { get; set; } = "/../";
        public string CodeFileName { get; set; } = "Recorder.cs";
        public string ExportCodeRelativePath { get; set; } = "/../";
        public string ExportCodeAbsolutePath { get; set; } = "/../";
        public string RecorderTemplateFileRelativePath { get; set; } = "/../";
        public string EnumTemplateFileRelativePath { get; set; } = "/../";
        public string StructTemplateFileRelativePath { get; set; } = "/../";
        public string UnknownClassTypeNamePrefix { get; set; } = "";
        public string NameSpaceRootName { get; set; } = "";
        public string RecorderKeyUnknownClassTypeToInt { get; set; } = "";
        public Dictionary<string, ClassTypeInfoStruct> ClassTypeInfoMap = new Dictionary<string, ClassTypeInfoStruct>();
        public Dictionary<string, PropertyTemplateInfoStruct> ClassPropertyTemplateMap { get; private set; } = new Dictionary<string, PropertyTemplateInfoStruct>();
        public Dictionary<string, string> ConvertFunctionTemplateMap { get; private set; } = new Dictionary<string, string>();
        public Dictionary<string, string> ToStringFunctionTemplateMap { get; private set; } = new Dictionary<string, string>();

        public string RecorderTemplateContent { get; set; } = "";
        public string EnumTemplateContent { get; set; } = "";
        public string StructTemplateContent { get; set; } = "";

        public void Load(XElement root)
        {
            foreach (XElement xElement in root.Elements())
            {
                string attributeName = xElement.Name.LocalName;
                string attributeValue = xElement.Value;

                if (attributeName == "CodeName")
                {
                    CodeName = attributeValue;
                }
                else if (attributeName == "XmlFileName")
                {
                    XmlFileName = attributeValue;
                }
                else if (attributeName == "XmlAttributeFirst")
                {
                    XmlAttributeFirst = Convert.ToBoolean(attributeValue);
                }
                else if (attributeName == "XmlRootNodeName")
                {
                    XmlRootNodeName = attributeValue;
                }
                else if (attributeName == "XmlRecorderNodeName")
                {
                    XmlRecorderNodeName = attributeValue;
                }
                else if (attributeName == "ExportAllXmlToTheSamePath")
                {
                    ExportAllXmlToTheSamePath = Convert.ToBoolean(attributeValue);
                }
                else if (attributeName == "ExportXmlRelativePath")
                {
                    ExportXmlRelativePath = attributeValue;
                }
                else if (attributeName == "CodeFileName")
                {
                    CodeFileName = attributeValue;
                }
                else if (attributeName == "ExportCodeRelativePath")
                {
                    ExportCodeRelativePath = attributeValue;
                }
                if (attributeName == "RecorderTemplateFileRelativePath")
                {
                    RecorderTemplateFileRelativePath = attributeValue;
                }
                else if (attributeName == "EnumTemplateFileRelativePath")
                {
                    EnumTemplateFileRelativePath = attributeValue;
                }
                else if (attributeName == "StructTemplateFileRelativePath")
                {
                    StructTemplateFileRelativePath = attributeValue;
                }
                else if (attributeName == "UnknownClassTypeNamePrefix")
                {
                    UnknownClassTypeNamePrefix = attributeValue;
                }
                else if (attributeName == "NameSpaceRootName")
                {
                    NameSpaceRootName = attributeValue;
                }
                else if (attributeName == "RecorderKeyUnknownClassTypeToInt")
                {
                    RecorderKeyUnknownClassTypeToInt = attributeValue;
                }
                else if (attributeName == "ClassTypeInfoMap")
                {
                    ClassTypeInfoMap.Clear();
                    foreach (var classTypeElement in xElement.Elements())
                    {
                        ClassTypeInfoStruct classTypeInfoStruct = new ClassTypeInfoStruct();
                        classTypeInfoStruct.normalValue = classTypeElement.Attribute("normalValue").Value;
                        classTypeInfoStruct.classValue = classTypeElement.Attribute("classValue").Value;
                        classTypeInfoStruct.isSpecial = Convert.ToBoolean(classTypeElement.Attribute("isSpecial").Value);
                        ClassTypeInfoMap.Add(classTypeElement.Name.LocalName, classTypeInfoStruct);
                    }
                }
                else if (attributeName == "ClassPropertyTemplateMap")
                {
                    ClassPropertyTemplateMap.Clear();
                    foreach (var CSClassPropertyTemplateElement in xElement.Elements())
                    {
                        PropertyTemplateInfoStruct propertyTemplateInfoStruct = new PropertyTemplateInfoStruct();
                        propertyTemplateInfoStruct.addInAttribute = Convert.ToBoolean(CSClassPropertyTemplateElement.Attribute("addInAttribute").Value);
                        propertyTemplateInfoStruct.addInElement = Convert.ToBoolean(CSClassPropertyTemplateElement.Attribute("addInElement").Value);
                        propertyTemplateInfoStruct.content = CSClassPropertyTemplateElement.Value;
                        propertyTemplateInfoStruct.addIsSpecialClassType = Convert.ToBoolean(CSClassPropertyTemplateElement.Attribute("addIsSpecialClassType").Value);
                        propertyTemplateInfoStruct.addIsNotSpecialClassType = Convert.ToBoolean(CSClassPropertyTemplateElement.Attribute("addIsNotSpecialClassType").Value);
                        XAttribute splitAttribute = CSClassPropertyTemplateElement.Attribute("split");
                        if (splitAttribute != null)
                        {
                            propertyTemplateInfoStruct.split = splitAttribute.Value;
                        }
                        else
                        {
                            propertyTemplateInfoStruct.split = "\n";
                        }
                        ClassPropertyTemplateMap.Add(CSClassPropertyTemplateElement.Name.LocalName, propertyTemplateInfoStruct);
                    }
                }
                else if (attributeName == "ConvertFunctionTemplateMap")
                {
                    ConvertFunctionTemplateMap.Clear();
                    foreach (var ConvertFunctionTemplateElement in xElement.Elements())
                    {
                        ConvertFunctionTemplateMap[ConvertFunctionTemplateElement.Name.LocalName] = ConvertFunctionTemplateElement.Value;
                    }
                }
                else if (attributeName == "ToStringFunctionTemplateMap")
                {
                    ToStringFunctionTemplateMap.Clear();
                    foreach (var ToStringFunctionTemplateElement in xElement.Elements())
                    {
                        ToStringFunctionTemplateMap[ToStringFunctionTemplateElement.Name.LocalName] = ToStringFunctionTemplateElement.Value;
                    }
                }
            }
        }
    }

    /// <summary>
    /// 配置类
    /// </summary>
    class ConfigData
    {
        public string ProjectVersionTool { get; set; } = "git";
        public string ImportXlsxRelativePath { get; set; } = "/../";
        public string ImportXlsxAbsolutePath { get; set; } = "/../";

        public Dictionary<string, CodeConfigData> CodeConfigDataMap { get; private set; } = new Dictionary<string, CodeConfigData>();

        string configPath = "";
        /// <summary>
        /// 是否需要保存配置
        /// </summary>
        public bool NeedSave { get; set; } = false;

        static ConfigData configData = null;

        public static void Init(string configPath)
        {
            if (configData == null)
            {
                configData = new ConfigData();
                configData.configPath = configPath;
                configData.Load();
                configData.Check();
            }
        }

        public static void UnInit()
        {
            if(configData!=null)
            {
                if (configData.NeedSave)
                {
                    configData.Save();
                }
                configData = null;
            }
        }

        /// <summary>
        /// 获得单例
        /// </summary>
        /// <returns></returns>
        public static ConfigData GetSingle()
        {
            return configData;
        }

        /// <summary>
        /// 加载
        /// </summary>
        void Load()
        {
            if (!File.Exists(configPath))
            {
                XlsxManager.Log(false, $"配置文件不存在!{configPath}");
                return;
            }
            XDocument doc = XDocument.Load(configPath);
            if (doc == null)
            {
                XlsxManager.Log(false, $"配置文件加载失败!{configPath}");
                return;
            }
            foreach (XElement xElement in doc.Root.Elements())
            {
                string attributeName = xElement.Name.LocalName;
                string attributeValue = xElement.Value;
                if (attributeName == "ImportXlsxRelativePath")
                {
                    ImportXlsxRelativePath = attributeValue;
                }
                else if (attributeName == "ProjectVersionTool")
                {
                    ProjectVersionTool = attributeValue;
                }
                else if (attributeName == "CodeConfigDataMap")
                {
                    CodeConfigDataMap.Clear();
                    foreach (var item in xElement.Elements())
                    {
                        CodeConfigData codeConfigData = new CodeConfigData();
                        codeConfigData.Load(item);
                        CodeConfigDataMap.Add(codeConfigData.CodeName, codeConfigData);
                    }
                }
            }
        }

        /// <summary>
        /// 检测
        /// </summary>
        public void Check()
        {
            ImportXlsxAbsolutePath = System.IO.Path.GetFullPath(XlsxManager.GetToolRootPath() + ImportXlsxRelativePath);
            foreach (var item in configData.CodeConfigDataMap)
            {
                item.Value.ExportXmlAbsolutePath = System.IO.Path.GetFullPath(XlsxManager.GetToolRootPath() + item.Value.ExportXmlRelativePath);
                item.Value.ExportCodeAbsolutePath = System.IO.Path.GetFullPath(XlsxManager.GetToolRootPath() + item.Value.ExportCodeRelativePath);
                if (!File.Exists(XlsxManager.GetToolRootPath() + item.Value.RecorderTemplateFileRelativePath))
                {
                    XlsxManager.Log(false, $"代码类:{item.Key}缺少Recorder模板！");
                }
                else
                {
                    using (StreamReader streamReader = new StreamReader(XlsxManager.GetToolRootPath() + item.Value.RecorderTemplateFileRelativePath))
                    {
                        item.Value.RecorderTemplateContent = streamReader.ReadToEnd();
                    }
                }
                if (!File.Exists(XlsxManager.GetToolRootPath() + item.Value.EnumTemplateFileRelativePath))
                {
                    XlsxManager.Log(false, $"代码类:{item.Key}缺少CSEnum模板！");
                }
                else
                {
                    using (StreamReader streamReader = new StreamReader(XlsxManager.GetToolRootPath() + item.Value.EnumTemplateFileRelativePath))
                    {
                        item.Value.EnumTemplateContent = streamReader.ReadToEnd();
                    }
                }
                if (!File.Exists(XlsxManager.GetToolRootPath() + item.Value.StructTemplateFileRelativePath))
                {
                    XlsxManager.Log(false, $"代码类:{item.Key}缺少CSStruct模板！");
                }
                else
                {
                    using (StreamReader streamReader = new StreamReader(XlsxManager.GetToolRootPath() + item.Value.StructTemplateFileRelativePath))
                    {
                        item.Value.StructTemplateContent = streamReader.ReadToEnd();
                    }
                }
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            if (!File.Exists(configPath))
            {
                return;
            }
            XDocument doc = XDocument.Load(configPath);
            if (doc == null)
            {
                return;
            }
            doc.Root.Element("ImportXlsxRelativePath").Value = ImportXlsxRelativePath;
            foreach (var item in doc.Root.Element("CodeConfigDataMap").Elements())
            {
                item.Element("ExportXmlRelativePath").Value = CodeConfigDataMap[item.Element("CodeName").Value].ExportXmlRelativePath;
                item.Element("ExportCodeRelativePath").Value = CodeConfigDataMap[item.Element("CodeName").Value].ExportCodeRelativePath;
            }

            //保存时忽略声明
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;
            xws.NewLineChars = "\n";
            FileStream fileStream = new FileStream(configPath, FileMode.Create, FileAccess.ReadWrite);
            using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
            {
                doc.Save(xmlWriter);
            }
        }
    }
}
