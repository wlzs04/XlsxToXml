using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace XlsxToXml
{
    class ConfigData
    {
        public string ImportXlsxRelativePath { get; set; } = "/../";
        public string ExportXmlRelativePath { get; set; } = "/../";
        public string ExportCSRelativePath { get; set; } = "/../";
        public string CSClassTemplateFileRelativePath { get; private set; } = "/CSClassTemplate.txt";
        
        string configPath = "";

        public ConfigData(string configPath)
        {
            this.configPath = configPath;
            Load();
        }

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
                else if (attributeName == "CSClassTemplateFileRelativePath")
                {
                    CSClassTemplateFileRelativePath = attributeValue;
                }
            }
        }

        public void Save()
        {
            XDocument doc = new XDocument(
                new XElement("Config",
                    new XElement("ImportXlsxRelativePath", ImportXlsxRelativePath),
                    new XElement("ExportXmlRelativePath", ExportXmlRelativePath),
                    new XElement("ExportCSRelativePath", ExportCSRelativePath),
                    new XElement("CSClassTemplateFileRelativePath", CSClassTemplateFileRelativePath)
                )
            );
            //保存时忽略声明
            XmlWriterSettings xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = true;
            xws.Indent = true;
            FileStream fileStream = new FileStream(configPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            using (XmlWriter xmlWriter = XmlWriter.Create(fileStream, xws))
            {
                doc.Save(xmlWriter);
            }
        }
    }
}
