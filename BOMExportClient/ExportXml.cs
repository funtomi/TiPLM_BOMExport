using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Thyt.TiPLM.Common;

namespace BOMExportClient {
    public class ExportXml {
        public ExportXml() {

        }

        /// <summary>
        /// 导出到xml
        /// </summary>
        /// <param name="list"></param>
        /// <param name="eventOid"></param>
        /// <param name="userOid"></param>
        public static void ExportToXml(List<DataTable> list,string rootName) {
            if (list==null||list.Count==0) {
                return;
            } 
            try {
                string xlsPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string fPath = Path.GetFullPath("..\\ExportFiles");
                string newFile = Path.Combine(fPath, BomExportClient.item.Id + "_" + rootName + ".xml");
                var doc = CreateXmlDocument(list, newFile, rootName);
                doc.Save(newFile);
                
            } catch (Exception ex) {
                PLMEventLog.WriteExceptionLog(ex);
            }
        }

        private static XmlDocument CreateXmlDocument(List<DataTable> list, string newFile, string rootName) {
            if (list==null||list.Count==0) {
                return null; 
            }
            var firstDs = list[0];
            if (firstDs==null) {
                return null;
            }
            firstDs.WriteXml(newFile);
            XmlDocument doc = new XmlDocument();
            doc.Load(newFile);
            doc = AddXmlSchema(doc, rootName);
            for (int i = 1; i < list.Count; i++) {
                var dt = list[i];
                dt.WriteXml(newFile);
                XmlDocument newDoc = new XmlDocument();
                newDoc.Load(newFile);
                var node = doc.ImportNode(newDoc.DocumentElement, true);
                string path = string.Format("ufinterface//{0}",rootName);
                doc.SelectSingleNode(path).AppendChild(node.FirstChild);
            }
            return doc;
        }

        private static XmlDocument CreateBomXmlSchema(string str, string tableName) {
            XmlDocument doc = new XmlDocument();
            doc.Load(str);
            doc = AddXmlSchema(doc, tableName);
            return doc;
        }

        /// <summary>
        /// 将datatable转为xml
        /// </summary>
        /// <param name="xmlDS"></param>
        /// <returns></returns>
        private static string ConvertDataTableToXML(DataTable xmlDS) {
            MemoryStream stream = null;
            XmlTextWriter writer = null;
            try {
                stream = new MemoryStream();
                writer = new XmlTextWriter(stream, Encoding.Default);
                xmlDS.WriteXml(writer);
                int count = (int)stream.Length;
                byte[] arr = new byte[count];
                stream.Seek(0, SeekOrigin.Begin);
                stream.Read(arr, 0, count);
                var str = Encoding.ASCII.GetString(arr).Trim();
                return str;
            } catch {
                return String.Empty;
            } finally {
                if (writer != null) writer.Close();
            }
        }

        /// <summary>
        /// 构建xml框架
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private static XmlDocument AddXmlSchema(XmlDocument doc, string tableName) {
            if (string.IsNullOrEmpty(tableName) || doc == null) {
                return null;
            }
            XmlDocument newDoc = new XmlDocument();
            var dec = newDoc.CreateXmlDeclaration("1.0", "utf-8", "yes");
            newDoc.AppendChild(dec);
            //var node = doc.SelectSingleNode(tableName);
            var root = newDoc.CreateElement("ufinterface");

            #region 默认，无需修改

            root.SetAttribute("receiver", "U8");//接收方 
            //档案或单据模版名，填档案或单据的唯一标识，如:客商档案：customer，客商分类：customerclass ，具体名称由总体确定，在数据交换中该名称要经常使用；
            root.SetAttribute("roottag", tableName);
            //proc：操作类型，分为“增删改查”，对应填Add / Delete /Edit /Query（必填），该字段导入操作，请填写Add / Delete /Edit，导出操作，请填写Query；
            root.SetAttribute("proc", "Edit");
            //编码是否已转换，该字段在导入的时候使用，如果已转换即已和U8基础数据编码一致填Y，将不会通过对照表的转换，如果没有转换即和U8基础数据编码不一致填N，将会自动通过对照表转换之后，进行相应的操作
            root.SetAttribute("codeexchanged", "N");
            root.SetAttribute("paginate", "0");
            root.SetAttribute("family", "生成制造");
            #endregion
            //发送方，填外部系统注册码（必填）
            root.SetAttribute("sender", "001");
            var tableNode = newDoc.CreateElement(tableName);
            
            var node = newDoc.ImportNode(doc.DocumentElement, true);
            tableNode.AppendChild(node.FirstChild);
            root.AppendChild(tableNode);
            newDoc.AppendChild(root);
            //doc.DocumentElement.InsertBefore(root, doc.DocumentElement.FirstChild);
            return newDoc;

        }

    }
}
