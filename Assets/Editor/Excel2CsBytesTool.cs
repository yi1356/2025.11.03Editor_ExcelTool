using System.IO;
using Excel;//Z：Excel读取库（解析xlsx文件核心）
using System.Data;//Z：数据集合处理（存储Excel读取后的表格数据）
using UnityEditor;//Z：Unity编辑器扩展（创建菜单、刷新资源）
using UnityEngine;
using System.Collections.Generic;//Z：泛型集合（List存储数据）
using System.Text;//Z：字符串构建（高效拼接C#/XML文本）
using System.Reflection;
using Table;//Z：自定义命名空间（存放生成的C#数据类）
using System;
using System.Xml;//Z：XML处理（辅助序列化）
using System.Xml.Serialization;//Z：XML序列化（将对象转为XML/从XML还原对象）
using System.Runtime.Serialization.Formatters.Binary;//Z：二进制序列化（将对象转为bytes）

/// <summary>
/// Excel生成bytes和cs工具
/// Z：仅生成 C# 脚本：提取 Excel 的字段名、类型、描述，生成带序列化标签的类（含数据加载方法）
///    生成 bytes 文件：先将 Excel 数据转为临时 XML，再反序列化XML为对象，将这个对象序列化为二进制文件，同时可自动删除临时 XML
/// </summary>
public class Excel2CsBytesTool
{
    static string ExcelDataPath = Application.dataPath + "/../ExcelData";//源Excel文件夹,xlsx格式(Z：/../ 是路径中的 “上级目录” 跳转符)
    static string BytesDataPath = Application.dataPath + "/Resources/DataTable";//生成的bytes文件夹
    static string CsClassPath = Application.dataPath + "/Scripts/DataTable";//生成的c#脚本文件夹
    static string XmlDataPath = ExcelDataPath + "/tempXmlData";//生成的xml(临时)文件夹..
    static string AllCsHead = "all";//序列化结构体的数组类.类名前缀 （Z：生成的 “数组容器类” 前缀（如 Excel 名为weapon，则容器类为allweapon））

    static char ArrayTypeSplitChar = '#';//数组类型值拆分符: int[] 1#2#34（Z：也就是这个数组中存的是[1,2,34]） ...
    static bool IsDeleteXmlInFinish = false ;//生成bytes后是否删除中间文件xml

    [MenuItem("SDGSupporter/Excel/Excel2Cs")]
    static void Excel2Cs()
    {
        Init();
        Excel2CsOrXml(true);
    }

    [MenuItem("SDGSupporter/Excel/Excel2Bytes")]
    static void Excel2Xml2Bytes()
    {
        Init();
        //生成中间文件xml
        Excel2CsOrXml(false);
        //生成bytes
        WriteBytes();
    }

    //Z：初始化、作用是检测文件夹是否存在，不存在则需要创建
    static void Init()
    {
        if (!Directory.Exists(CsClassPath))
        {
            Directory.CreateDirectory(CsClassPath);
        }
        if (!Directory.Exists(XmlDataPath))
        {
            Directory.CreateDirectory(XmlDataPath);
        }
        if (!Directory.Exists(BytesDataPath))
        {
            Directory.CreateDirectory(BytesDataPath);
        }
    }

    /// <summary>
    /// Z：生成CS脚本的方法； 根据Excel 提取的类名、字段名、字段类型、字段描述，生成带序列化标签的C# 类文件。
    /// </summary>
    /// <param name="className"></param>
    /// <param name="names"></param>
    /// <param name="types"></param>
    /// <param name="descs"></param>
    static void WriteCs(string className, string[] names, string[] types, string[] descs)
    {
        try
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("using System;");
            stringBuilder.AppendLine("using System.Collections.Generic;");
            stringBuilder.AppendLine("using System.IO;");
            stringBuilder.AppendLine("using System.Runtime.Serialization.Formatters.Binary;");
            stringBuilder.AppendLine("using System.Xml.Serialization;");
            stringBuilder.Append("\n");
            stringBuilder.AppendLine("namespace Table");//Z：定义了Table命名空间，将生成的类封装在该命名空间下
            stringBuilder.AppendLine("{");
            stringBuilder.AppendLine("    [Serializable]");//Z：标记类为[Serializable]，使其支持二进制序列化，以便后续转换为 bytes 文件
            stringBuilder.AppendLine("    public class " + className);
            stringBuilder.AppendLine("    {");

            //Z：写属性
            for (int i = 0; i < names.Length; i++)//Z：遍历 Excel 中提取的字段名（names）、类型（types）、描述（descs）数组？？？(其实就是列属性，id、名字、预制体名、描述 ；但不太理解的地方在于为什么有的竟然是用列表存的？？？)
            {
                stringBuilder.AppendLine("        /// <summary>");
                stringBuilder.AppendLine("        /// " + descs[i]);
                stringBuilder.AppendLine("        /// </summary>");
                stringBuilder.AppendLine("        [XmlAttribute(\"" + names[i] + "\")]");//Z：为每个字段添加[XmlAttribute]或[XmlElementAttribute]特性，用于 XML 序列化时的映射； 让字段变成 XML 标签的 “属性

                string type = types[i];
                if (type.Contains("[]"))
                {
                    //type = type.Replace("[]", "");
                    //stringBuilder.AppendLine("        public List<" + type + "> " + names[i] + ";");

                    //可选代码：
                    //用_name字段去反序列化，name取_name.item的值,直接返回list<type>。
                    //因为xml每行可能有多个数组字段，这样就多了一层变量item，所以访问的时候需要.item才能取到list<type>
                    //因此用额外的一个变量直接返回List<type>      
                    type = type.Replace("[]", "");
                    stringBuilder.AppendLine("        public List<" + type + "> " + names[i] + "");
                    stringBuilder.AppendLine("        {");
                    stringBuilder.AppendLine("            get");
                    stringBuilder.AppendLine("            {");
                    stringBuilder.AppendLine("                if (_" + names[i] + " != null)");
                    stringBuilder.AppendLine("                {");
                    stringBuilder.AppendLine("                    return _" + names[i] + ".item;");
                    stringBuilder.AppendLine("                }");
                    stringBuilder.AppendLine("                return null;");
                    stringBuilder.AppendLine("            }");
                    stringBuilder.AppendLine("        }");
                    stringBuilder.AppendLine("        [XmlElementAttribute(\"" + names[i] + "\")]");   //Z：让字段变成 XML 的 “子标签”
                    stringBuilder.AppendLine("        public " + type + "Array _" + names[i] + ";");
                }
                else
                {
                    stringBuilder.AppendLine("        public " + type + " " + names[i] + ";");
                }

                stringBuilder.Append("\n");
            }

            //Z：写从二进制文件中 读取数据的方法（Z：其实最后得到的就是 List<Weapon>）
            stringBuilder.AppendLine("        public static List<" + className + "> LoadBytes()");
            stringBuilder.AppendLine("        {");
            stringBuilder.AppendLine("            string bytesPath = \"" + BytesDataPath + "/" + className + ".bytes\";");
            stringBuilder.AppendLine("            if (!File.Exists(bytesPath))");
            stringBuilder.AppendLine("                return null;");
            stringBuilder.AppendLine("            using (FileStream stream = new FileStream(bytesPath, FileMode.Open))");
            stringBuilder.AppendLine("            {");
            stringBuilder.AppendLine("                BinaryFormatter binaryFormatter = new BinaryFormatter();");
            stringBuilder.AppendLine("                all" + className + " table = binaryFormatter.Deserialize(stream) as all" + className + ";");
            stringBuilder.AppendLine("                return table." + className + "s;");
            stringBuilder.AppendLine("            }");
            stringBuilder.AppendLine("        }");
            stringBuilder.AppendLine("    }");
            stringBuilder.Append("\n");
            stringBuilder.AppendLine("    [Serializable]");//Z：定义了一个数组容器类，该类包含一个List<className>类型的字段（className + "s"），用于存储整个 Excel 表的所有数据行
            stringBuilder.AppendLine("    public class " + AllCsHead + className);    
            stringBuilder.AppendLine("    {");
            stringBuilder.AppendLine("        public List<" + className + "> " + className + "s;");
            stringBuilder.AppendLine("    }");
            stringBuilder.AppendLine("}");

            string csPath = CsClassPath + "/" + className + ".cs";
            if (File.Exists(csPath))//Z：判断如果有原 className.cs的文件，就先清除      
            {
                File.Delete(csPath);
            }

            //Z：把拼接好的 C# 代码写入到指定的.cs 文件中（其实就是创建了weapon.cs这个脚本）*******************************
            using (StreamWriter sw = new StreamWriter(csPath))// [StreamWriter 是字节流写入]
            {
                sw.Write(stringBuilder);
                Debug.Log("生成:" + csPath);
            }
        }
        catch (System.Exception e)
        {
            Debug.LogError("写入CS失败:" + e.Message);
            throw;
        }
    }


    /// <summary>
    /// Z：是将结构化数据（类名、字段名、字段类型、数据列表）转换为特定格式的 XML 文件；datasList：数据列表（每个元素是一个字符串数组，代表一行数据）
    /// </summary>
    /// <param name="className"></param>
    /// <param name="names"></param>
    /// <param name="types"></param>
    /// <param name="datasList"></param>
    static void WriteXml(string className, string[] names, string[] types, List<string[]> datasList)
    {
        try
        {
            StringBuilder stringBuilder = new StringBuilder();//Z：用于拼接 XML 字符串
            stringBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            stringBuilder.AppendLine("<" + AllCsHead + className + ">");//Z：写入最外层根节点，格式为<AllCsHead + className>
            stringBuilder.AppendLine("<" + className + "s>");  //Z：二级节点<className + s>（其实就是<Weapons>）
            for (int d = 0; d < datasList.Count; d++)//Z：其实这里的一次循环，写入的就是一行中的内容
            {
                stringBuilder.Append("\t<" + className + " ");
                //单行数据
                string[] datas = datasList[d];
                //填充属性节点（Z：其实就是 <weapon id="1" name="寒冰箭" prefabName="hanbingjian.prefab">  这一行）；
                for (int c = 0; c < datas.Length; c++)
                {
                    string type = types[c];
                    if (!type.Contains("[]"))
                    {
                        string name = names[c];
                        string value = datas[c];
                        stringBuilder.Append(name + "=\"" + value + "\"" + (c == datas.Length - 1 ? "" : " "));
                    }
                }
                stringBuilder.Append(">\n");
                //填充子元素节点(数组类型字段)（Z：其实填的就是数组类型 ）
                for (int c = 0; c < datas.Length; c++)
                {
                    string type = types[c];
                    if (type.Contains("[]"))
                    {
                        string name = names[c];
                        string value = datas[c];
                        string[] values = value.Split(ArrayTypeSplitChar);
                        stringBuilder.AppendLine("\t\t<" + name + ">");
                        for (int v = 0; v < values.Length; v++)
                        {
                            stringBuilder.AppendLine("\t\t\t<item>" + values[v] + "</item>");
                        }
                        stringBuilder.AppendLine("\t\t</" + name + ">");
                    }
                }
                stringBuilder.AppendLine("\t</" + className + ">");//Z：其实就是 <Weapon>
            }
            stringBuilder.AppendLine("</" + className + "s>");//Z：其实就是 <Weapons>
            stringBuilder.AppendLine("</" + AllCsHead + className + ">");  //Z：其实就是 </allweapon>

            string xmlPath = XmlDataPath + "/" + className + ".xml";
            if (File.Exists(xmlPath))
            {
                File.Delete(xmlPath);
            }
            using (StreamWriter sw = new StreamWriter(xmlPath))
            {
                sw.Write(stringBuilder);
                Debug.Log("生成文件:" + xmlPath);
            }
        }
        catch (System.Exception e)
        {
            Debug.LogError("写入Xml失败:" + e.Message);
        }
    }

    /// <summary>
    /// Z：将 Excel 文件转换为 C# 代码（.cs）或 XML 文件的工具方法，通过参数isCs控制转换目标类型
    /// 总之，这里①先将Excel文件中数据读取到DataSet数据结构中；
    /// ②通过两层循环遍历将字段名[第一行数据]存到string[] names
    /// ③将字段类型[第二行数据]存到string[] types
    /// ④将字段描述[第三行数据]存到string[] descs
    /// ⑤将真正的数据[第四行及以后的行] 存储到  List<string[]> datasList中
    /// ⑥根据选择将这些存储的数据写成CS/XML
    /// </summary>
    /// <param name="isCs">true代表转化CS文件、false代表转化为XML文件</param>
    static void Excel2CsOrXml(bool isCs)
    {
        string[] excelPaths = Directory.GetFiles(ExcelDataPath, "*.xlsx");//Z：Directory.GetFiles用于获取指定目录下的所有符合条件的文件路径
        for (int e = 0; e < excelPaths.Length; e++)
        {
            //0.读Excel
            string className;//类型名（Z：用于生成C#类名或XML节点名）
            string[] names;//字段名（Z：Excel第一行数据）
            string[] types;//字段类型（Z：Excel第二行数据）
            string[] descs;//字段描述（Z：Excel第三行数据）
            List<string[]> datasList;//数据（Z：其实这里就是要存储每一行的 数据 ）

            try
            {
                string excelPath = excelPaths[e];//excel路径  
                className = Path.GetFileNameWithoutExtension(excelPath).ToLower();//Z：获取类名，其实就是取文件名（不含扩展名）并转为小写
                //Z：总之，这里作用是读取Excel表   到 DataSet这个数据结构中
                FileStream fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read);//Z： 以只读模式打开文件流
                IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);//Z：创建Excel读取器（针对.xlsx格式）
                // 表格数据全部读取到result里*********************************************************************************************************
                DataSet result = excelDataReader.AsDataSet();//Z：将Excel数据读取到DataSet中（表格数据结构化存储）

                // 获取表格列数
                int columns = result.Tables[0].Columns.Count;
                // 获取表格行数
                int rows = result.Tables[0].Rows.Count;
                // 根据行列依次读取表格中的每个数据
                names = new string[columns];//Z：存储每个字段名
                types = new string[columns];//Z：存储每个字段类型
                descs = new string[columns];//Z：存储每个字段的描述
                datasList = new List<string[]>();//Z：初始化datasList用于存储多行数据（每行是一个字符串数组 [这里存的就是字段的值]，长度为列数）
                //Z：总之，这里两个循环的作用就是，将Excel表中的数据 分别存储到字段名数组string[] names、字段类型数组string[] types、字段描述数组string[] descs、行数据列表List<string[]> datasList中  ***************************************************************
                for (int r = 0; r < rows; r++)//Z：第一层循环，遍历每一行
                {
                    string[] curRowData = new string[columns];
                    for (int c = 0; c < columns; c++)//Z：第二层循环遍历每一列
                    {
                        //解析：获取第一个表格中指定行指定列的数据
                        string value = result.Tables[0].Rows[r][c].ToString();//Z：取第r行第c列的数据，转为字符串

                        //清除前两行的变量名、变量类型 首尾空格（Z：可以理解为用来防止错误）
                        if (r < 2)
                        {
                            value = value.TrimStart(' ').TrimEnd(' ');
                        }

                        curRowData[c] = value;//Z：存入当前行数据数组
                    }
                    //解析：第一行类变量名（Z：names数组=存储当前行数据的 临时数组）
                    if (r == 0)
                    {
                        names = curRowData;
                    }//解析：第二行类变量类型
                    else if (r == 1)
                    {
                        types = curRowData;
                    }//解析：第三行类变量描述
                    else if (r == 2)
                    {
                        descs = curRowData;
                    }//解析：第四行开始是数据（Z：从第四行开始就是真正的数据了）
                    else
                    {
                        datasList.Add(curRowData);
                    }
                }
            }

            //Z：将从Excel中读取的数据 写成Cs / Xml  ***********************************************************************************************
            catch (System.Exception exc)
            {
                Debug.LogError("请关闭Excel:" + exc.Message);
                return;
            }

            if (isCs)
            {
                //写Cs
                WriteCs(className, names, types, descs);
            }
            else
            {
                //写Xml
                WriteXml(className, names, types, datasList);
            }
        }

        AssetDatabase.Refresh();//Z：AssetDatabase是 Unity 的资源数据库类，Refresh()方法用于刷新资源，确保新生成的文件（.cs 或.xml）能被 Unity 编辑器立即识别
    }


    /// <summary>
    /// Z：将对应的 XML 文件反序列化为对象，再将对象序列化为二进制文件（.bytes），并可选择删除原始 XML 文件及目录  
    /// 核心逻辑是：加载指定 DLL → 筛选特定数据类 → 读取对应 XML 文件并反序列化为对象 → 将对象序列化为二进制文件 → 可选删除 XML 文件及目录
    /// 思路：
    /// ①根据反射找到特定数据类（包含all的类）
    /// ②根据类名(处理后)来找到对应的XML文件
    /// ③反序列化这个XML文件为这个特定数据类的对象
    /// ④将这个对象序列化为二进制文件并存储（后面真正用到的数据是这个二进制文件）
    /// </summary>
    static void WriteBytes()
    {
        string csAssemblyPath = Application.dataPath + "/../Library/ScriptAssemblies/Assembly-CSharp.dll";
        Assembly assembly = Assembly.LoadFile(csAssemblyPath);//Z：加载程序集，方便后续使用反射
        if (assembly != null)
        {
            Type[] types = assembly.GetTypes();//Z：从加载的 DLL 中获取所有定义的类型（类、结构体、接口等），存储在types数组中
            for (int i = 0; i < types.Length; i++)//Z：遍历所有获取到的类型，逐个处理
            {
                Type type = types[i];
                if (type.Namespace == "Table" && type.Name.Contains(AllCsHead))//Z：筛选命名空间为Table，且类名包含AllCsHead的类型；？？？？？？？？？？？？？？？？？？？？
                {
                    string className = type.Name.Replace(AllCsHead, "");//Z：移除AllCsHead前缀，得到基础类名

                    //读取xml数据
                    string xmlPath = XmlDataPath + "/" + className + ".xml";
                    if (!File.Exists(xmlPath))
                    {
                        Debug.LogError("Xml文件读取失败:" + xmlPath);
                        continue;
                    }
                    object table;//Z：声明一个object类型变量table，用于存储从 XML 反序列化得到的对象（Z：其实就是allweapon对象）
                    using (Stream reader = new FileStream(xmlPath, FileMode.Open))
                    {
                        //读取xml实例化table: all+classname
                        //object table = assembly.CreateInstance("Table." + type.Name);
                        XmlSerializer xmlSerializer = new XmlSerializer(type);//Z：创建XmlSerializer实例，指定反序列化的目标类型？？？？？？？？？？？？？？？介绍XmlSerializer 和  BinaryFormatter 
                        table = xmlSerializer.Deserialize(reader);//Z：从文件流中反序列化 XML 内容，得到table对象（Z：其实就是allweapon对象    ）***************************************************************       
                    }
                    //obj序列化二进制
                    string bytesPath = BytesDataPath + "/" + className + ".bytes";
                    if (File.Exists(bytesPath))
                    {
                        File.Delete(bytesPath);
                    }
                    using (FileStream fileStream = new FileStream(bytesPath, FileMode.Create))//Z：模式为Create（创建新文件，若已存在则覆盖）
                    {
                        //Z：创建BinaryFormatter实例，用于将对象序列化为二进制数据
                        BinaryFormatter binaryFormatter = new BinaryFormatter();
                        binaryFormatter.Serialize(fileStream, table);//Z：将table对象通过BinaryFormatter序列化到文件流中，生成二进制文件
                        Debug.Log("生成:" + bytesPath);
                    }

                    if (IsDeleteXmlInFinish)
                    {
                        File.Delete(xmlPath);
                        Debug.Log("删除:" + bytesPath);
                    }
                }
            }
        }

        if (IsDeleteXmlInFinish)
        {
            Directory.Delete(XmlDataPath);
            Debug.Log("删除:" + XmlDataPath);
        }
    }
}
