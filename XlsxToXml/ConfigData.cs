using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace XlsxToXml
{
    /// <summary>
    /// 配置类
    /// </summary>
    class ConfigData
    {
        public string ImportXlsxRelativePath { get; set; } = "/../";
        public string ExportXmlRelativePath { get; set; } = "/../";
        public string ExportCSRelativePath { get; set; } = "/../";
        public string ProjectVersionTool { get; set; } = "git";
        public string CSClassTemplateFileRelativePath { get; private set; } = "/CSClassTemplate.txt";
        public string XmlFileName { get; private set; } = "Recorder.xml";
        public string CSClassFileName { get; private set; } = "Recorder.cs";
        
        public Dictionary<string,string> CSClassPropertyTemplateMap { get; private set; } = new Dictionary<string, string>();
        public Dictionary<string,string> ConvertFunctionTemplateMap { get; private set; } = new Dictionary<string, string>();

        string configPath = "";

        static ConfigData configData = new ConfigData(Environment.CurrentDirectory + "/Config.xml");

        private ConfigData(string configPath)
        {
            this.configPath = configPath;
            Load();
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
                return;
            }
            XDocument doc = XDocument.Load(configPath);
            if (doc == null)
            {
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
                else if (attributeName == "ExportXmlRelativePath")
                {
                    ExportXmlRelativePath = attributeValue;
                }
                else if (attributeName == "ExportCSRelativePath")
                {
                    ExportCSRelativePath = attributeValue;
                }
                else if (attributeName == "ProjectVersionTool")
                {
                    ProjectVersionTool = attributeValue;
                }
                else if (attributeName == "CSClassTemplateFileRelativePath")
                {
                    CSClassTemplateFileRelativePath = attributeValue;
                }
                else if (attributeName == "XmlFileName")
                {
                    XmlFileName = attributeValue;
                }
                else if (attributeName == "CSClassFileName")
                {
                    CSClassFileName = attributeValue;
                }
                else if (attributeName == "CSClassPropertyTemplateMap")
                {
                    CSClassPropertyTemplateMap.Clear();
                    foreach (var CSClassPropertyTemplateElement in xElement.Elements())
                    {
                        CSClassPropertyTemplateMap.Add(CSClassPropertyTemplateElement.Name.LocalName, CSClassPropertyTemplateElement.Value);
                    }
                }
                else if (attributeName == "ConvertFunctionTemplateMap")
                {
                    ConvertFunctionTemplateMap.Clear();
                    foreach (var ConvertFunctionTemplateElement in xElement.Elements())
                    {
                        ConvertFunctionTemplateMap.Add(ConvertFunctionTemplateElement.Name.LocalName, ConvertFunctionTemplateElement.Value);
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
            doc.Root.Element("ExportXmlRelativePath").Value = ExportXmlRelativePath;
            doc.Root.Element("ExportCSRelativePath").Value = ExportCSRelativePath;
            
            //保存时忽略声明
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;
            FileStream fileStream = new FileStream(configPath, FileMode.Create, FileAccess.ReadWrite);
            using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
            {
                doc.Save(xmlWriter);
            }
        }
    }
}
