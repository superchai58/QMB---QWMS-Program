using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace QWMS_SendData
{
    class XmlHelper
    {
        public static readonly XmlHelper Instance = new XmlHelper();

        #region XML序列化
        /// <summary>
        /// 序列化(不带声明头)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objEntity"></param>
        /// <returns></returns>
        public string Serializer<T>(object module)
        {
            StringBuilder strXml = new StringBuilder();
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;
            XmlSerializer xml = new XmlSerializer(typeof(T));
            try
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(strXml, settings))
                {
                    XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces();
                    namespaces.Add(string.Empty, string.Empty);
                    xml.Serialize(xmlWriter, module, namespaces);
                    xmlWriter.Close();
                }
                return strXml.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("序列化失败:" + ex.Message);
            }
        }

        #endregion

        #region XML反序列化
        /// <summary>
        /// 反序列化
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strXml"></param>
        /// <returns></returns>
        public object Deserialize<T>(string strXml)
        {
            try
            {
                using (StringReader reader = new StringReader(strXml))
                {
                    XmlSerializer xmldes = new XmlSerializer(typeof(T));
                    return xmldes.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("反序列化失败:" + ex.Message);
            }
        }
        /// <summary>
        /// 反序列化
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="Stream"></param>
        /// <returns></returns>
        public object Deserialize<T>(Stream Stream)
        {
            XmlSerializer xmldes = new XmlSerializer(typeof(T));
            return xmldes.Deserialize(Stream);
        }
        #endregion

        /// <summary>
        /// json字符串转换为Xml对象
        /// </summary>
        /// <param name="sJson"></param>
        /// <returns></returns>
        //public XmlDocument Json2Xml(string sJson)
        //{
        //    System.Web.Script.Serialization.JavaScriptSerializer oSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
        //    XmlDocument doc = new XmlDocument();
        //    XmlElement nRoot = doc.CreateElement("root");
        //    doc.AppendChild(nRoot);

        //    var oValue = oSerializer.DeserializeObject(sJson);
        //    if (oValue.GetType() == typeof(Dictionary<string, object>))
        //    {
        //        Dictionary<string, object> Dic = (Dictionary<string, object>)oValue;
        //        foreach (KeyValuePair<string, object> item in Dic)
        //        {
        //            //检查空
        //            if (item.Value.GetType() == typeof(object[]) && ((object[])item.Value).Length == 0)
        //                return null;
        //            XmlElement element = doc.CreateElement(item.Key);
        //            KeyValue2Xml(element, item);
        //            nRoot.AppendChild(element);
        //        }
        //    }
        //    else if (oValue.GetType() == typeof(object[]))
        //    {
        //        object[] o = oValue as object[];
        //        if (o.Length == 0)
        //            return null;
        //        for (int i = 0; i < o.Length; i++)
        //        {
        //            XmlElement element = doc.CreateElement("Item");
        //            KeyValuePair<string, object> item = new KeyValuePair<string, object>("Item", o[i]);
        //            KeyValue2Xml(element, item);
        //            nRoot.AppendChild(element);
        //        }
        //    }
        //    else
        //    {
        //        return null;
        //    }

        //    return doc;
        //}

        private void KeyValue2Xml(XmlElement node, KeyValuePair<string, object> Source)
        {
            object kValue = Source.Value;
            if (kValue.GetType() == typeof(Dictionary<string, object>))
            {
                foreach (KeyValuePair<string, object> item in kValue as Dictionary<string, object>)
                {
                    XmlElement element = node.OwnerDocument.CreateElement(item.Key);
                    KeyValue2Xml(element, item);
                    node.AppendChild(element);
                }
            }
            else if (kValue.GetType() == typeof(object[]))
            {
                object[] o = kValue as object[];
                for (int i = 0; i < o.Length; i++)
                {
                    XmlElement xitem = node.OwnerDocument.CreateElement("Item");
                    KeyValuePair<string, object> item = new KeyValuePair<string, object>("Item", o[i]);
                    KeyValue2Xml(xitem, item);
                    node.AppendChild(xitem);
                }

            }
            else
            {
                XmlText text = node.OwnerDocument.CreateTextNode(kValue.ToString());
                node.AppendChild(text);
            }
        }

        public XmlElement Dic2Xml(Dictionary<string, object> Dic, XmlDocument doc, XmlElement nRoot)
        {
            XmlElement nChild = doc.CreateElement("SAP");
            nRoot.AppendChild(nChild);
            foreach (KeyValuePair<string, object> item in Dic)
            {
                //检查空
                if (item.Value.GetType() == typeof(object[]) && ((object[])item.Value).Length == 0)
                    return null;
                XmlElement element = doc.CreateElement(item.Key);
                KeyValue2Xml(element, item);
                nChild.AppendChild(element);
            }
            return nRoot;
        }

        public string ConvertXmlToString(XmlDocument xmlDoc)
        {
            MemoryStream stream = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(stream, null);
            writer.Formatting = Formatting.Indented; xmlDoc.Save(writer);
            StreamReader sr = new StreamReader(stream, System.Text.Encoding.UTF8);
            stream.Position = 0;
            string xmlString = sr.ReadToEnd();
            sr.Close();
            stream.Close();
            return xmlString;
        }
    }
}
