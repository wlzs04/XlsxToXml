本项目用于导出配置。
可以将Excel(.xlsx)文档，导出为Xml(.xml)配置，和读取xml对应的C#(.cs)配置类。

项目目录下的Config.xml文件，用于配置项目功能。

ImportXlsxRelativePath 导入的Excel根路径和本项目的相对路径，程序关闭自动保存当前路径。
ExportXmlRelativePath 导出的Xml根路径和本项目的相对路径，程序关闭自动保存当前路径。
ExportCSRelativePath 导出的C#根路径和本项目的相对路径，程序关闭自动保存当前路径。
ProjectVersionTool 版本管理工具，目前支持svn和git，点击添加差异文件会自动添加未提交的.xlsx文件。

CSClassTemplateFileRelativePath C#模板类文件的相对路径
XmlFileName 导出的.xml配置文件的命名规则，可使用参数。
CSClassFileName 导出的.cs配置类文件的命名规则，可使用参数。
CSClassPropertyTemplateMap 属类性模板map，在C#模板类文件中对应位置将配置中的字符串替换给所有属性。
ConvertFunctionTemplateMap 根据类型进行转换方法模板map，配置中的字符串会直接替换到{convertFunction}中。

模板类文件和配置中可使用的参数
{recorderName}，配置文件名称，对应.xlsx文件的名称
{propertyValueName}，配置属性名称，对应.xlsx文件的第2行
{propertyClassName}，配置类型名称，对应.xlsx文件的第3行
{propertyDescription}，配置描述，对应.xlsx文件的第4行
{propertyConfigName}，配置名称，对应.xlsx文件的第5行
{convertFunction}，根据{propertyClassName}类型，在ConvertFunctionTemplateMap中替换对应的转换方法

.xlsx文件格式要求
第1行：是否需要导出，TRUE,FALSE
第2行：配置属性名称，一般用于代码的属性名称
第3行：配置类型名称，一般用于代码的属性类型
第4行：配置描述，一般用于描述配置的复杂规则，代码的注释
第5行：配置名称，一般用于描述配置的名称，代码的注释
第6及以下行：配置内容


