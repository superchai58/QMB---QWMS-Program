using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace QWMS_SendData
{
    [Serializable]
    [XmlRoot("responses")]
    public class XMLData
    {
        [XmlArray("responseItems"), XmlArrayItem("response")]
        public List<XMLResult> responseItems;
    }

    [Serializable]
    [XmlRoot("response")]
    public class XMLResult
    {
        [XmlElement("success")]
        public string success;
        [XmlElement("reason")]
        public string reason;
        [XmlElement("reasonDesc")]
        public string reasonDesc;
    }
}
