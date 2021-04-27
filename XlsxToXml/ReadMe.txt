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
普通参数：
{namespace}，命名空间，对应.cs文件根路径的相对路径
{fileName}，配置文件名称，对应.xlsx文件的名称
{key}，配置的第一列属性名称，用来作为索引，如果类型不是int会强转为int，需要保证配置文件中至少有一列导出
属性参数：会根据配置列数循环处理
{propertyValueName}，配置属性名称，对应????Recorder.xlsx文件的第2行，????Enum.xlsx文件的第1列
{propertyClassName}，配置类型名称，对应.xlsx文件的第3行
{propertyDescription}，配置描述，对应.xlsx文件的第4行
{propertyConfigName}，配置名称，对应????Recorder.xlsx文件的第5行，????Enum.xlsx文件的第2列
{convertFunction}，根据{propertyClassName}类型，在ConvertFunctionTemplateMap中替换对应的转换方法，默认使用custom类型的转换方式。
{structMapKeyName}，StructMap类型特有，将StructMap类型之后的Struct第一个字段的名称作为map的key

Struct文件特有
普通参数：
{prefix}，前缀，string类型，可空
{suffix}，后缀，string类型，可空
{split}，分割字符，char类型，不可空
属性参数：会根据配置列数循环处理
{propertyIndex}，配置属性索引，自动计算，从0开始
{propertyDefaultValue}，配置属性默认值

特殊类型：                    说明                             格式要求，例子：                             说明
SplitStringList    将string类型安分割字符分割组成List    SplitStringList int ;         字符串使用';'进行分割，每个子字符串都是int类型，将值组成list
SplitStringMap     将string类型安分割字符分割组成Map     SplitStringMap int,bool ;#    字符串先使用';'进行分割组成list，每个子字符串再使用'#'分割成key(int类型)和value(bool类型)，将key和value组成Map
ValueList          此列后指定个数的列都作为List的子节点   ValueList int 4               当前位置添加此list的长度，不超过4，此列的后4列为子节点，他们的内容作为value，安顺序组成List
KeyValueMap        此列后指定个数的列都作为Map的子节点    KeyValueMap int,bool 4        此列的后4列为子节点，他们的配置属性名称作为key，内容作为value，组成Map
StructList         此列后指定个数的Struct为list的节点，  StructList InstructionStruct 6 2  当前位置添加此list的长度，不超过2，此列的每6列为一个Struct，Struct作为子节点，共2个Struct，组成List
StructMap          此列后指定个数的Struct为map的节点，   StructMap BuffInfoStruct 6 2  当前位置添加此map的长度，不超过2，此列的每6列为一个Struct，Struct作为子节点，每个Struct的第一列作为key，Struct作为value，共2个Struct，组成Map

.xlsx文件要求

使用????Recorder.xlsx结尾代表配置内容文件，格式要求：
第1行：是否需要导出，TRUE,FALSE
第2行：配置属性名称，一般用于代码的属性名称
第3行：配置类型名称，一般用于代码的属性类型，类型包括其中list和map的转换方法比较特殊
第4行：配置描述，一般用于描述配置的复杂规则，代码的注释
第5行：配置名称，一般用于描述配置的名称，代码的注释
第6及以下行：配置内容

使用????Enum.xlsx结尾代表枚举类型文件，格式要求：
第1行：说明
第2及以下行：枚举内容
第1列：名称 第2列：含义

使用????Struct.xlsx结尾代表结构体类型文件，格式要求：
第1、2行：说明、内容
第1列：前缀，string类型，可空
第2列：后缀，string类型，可空
第3列：分割字符，char类型，不可空
第3行：空行
第4及以下行：结构体内容
第1列：名称 第2列：含义 第3列：类型 第4列：默认值
