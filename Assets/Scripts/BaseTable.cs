using System;
using System.Collections.Generic;
using System.Xml.Serialization;

/// <summary>
/// Excel2CsBytesTool
/// Excel可以配置的数组类型：string[] int[] bool[] 
/// 可自行扩展
/// Z：这三个类是 XML 和代码之间处理数组的 “中间格式”，让程序能方便地把数组存成 XML，或者从 XML 里读出数组来用，主要用于数据的存储和读取转换
/// </summary>
namespace Table
{
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public class stringArray
    {
        [System.Xml.Serialization.XmlElementAttribute("item")]
        public List<string> item { get; set; }
    }

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public class intArray
    {
        [System.Xml.Serialization.XmlElementAttribute("item")]
        public List<int> item { get; set; }
    }

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public class boolArray
    {
        [System.Xml.Serialization.XmlElementAttribute("item")]
        public List<bool> item { get; set; }
    }
}
