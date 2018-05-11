using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Thyt.TiPLM.DEL.Product;

namespace BOMExportClient {
    public abstract class BaseDal { 
        protected DEBusinessItem _dEBusinessItem;
        protected string _name;
        protected string _filePath;
        /// <summary>
        /// 构建保存路径
        /// </summary>
        /// <param name="item"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        protected string BuildFilePath(DEBusinessItem item,string name) {
            string fPath = Path.GetFullPath("..\\ExportFiles");
            string newFile = Path.Combine(fPath, item.Id + "_" + name + ".xml");
            return newFile;
        }
        private string GetFamilyStr(string tableName) {
            if (string.IsNullOrEmpty(tableName)) {
                return "";
            }
            switch (tableName) {
                default:
                case "Bom":
                    return "生产制造";
                case "Operation":
                    return "标准工序";
                case "Routing":
                    return "工艺路线";
                case "inventory":
                case "unitgroup":
                    return "基础档案";
            }
        }

        //private string GetRootTag(string tableName) {
        //    if (string.IsNullOrEmpty(tableName)) {
        //        return "";
        //    }
        //    switch (tableName) {
        //        default:
        //        case "Bom":
        //            return "bom";
        //        case "Operation":
        //            return "标准工序";
        //        case "Routing":
        //            return "工艺路线";
        //        case "inventory":
        //            return "inventory";

        //    }
        //}

        /// <summary>
        /// 创建xml框架(ufinterface节点和父节点)
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        protected XmlDocument CreateXmlSchema(string tableName, DEBusinessItem dEBusinessItem,string operatorStr) {
            if (string.IsNullOrEmpty(tableName) || dEBusinessItem == null) {
                return null;
            }
            XmlDocument doc = new XmlDocument();
            var dec = doc.CreateXmlDeclaration("1.0", "utf-8", "yes");
            doc.AppendChild(dec);
            //var node = doc.SelectSingleNode(tableName);
            var root = doc.CreateElement("ufinterface");

            #region 默认，无需修改
            root.SetAttribute("receiver", "U8");//接收方 
            //档案或单据模版名，填档案或单据的唯一标识，如:客商档案：customer，客商分类：customerclass ，具体名称由总体确定，在数据交换中该名称要经常使用；
            root.SetAttribute("roottag", tableName);
            //proc：操作类型，分为“增删改查”，对应填Add / Delete /Edit /Query（必填），该字段导入操作，请填写Add / Delete /Edit，导出操作，请填写Query；
            root.SetAttribute("proc", operatorStr);
            //编码是否已转换，该字段在导入的时候使用，如果已转换即已和U8基础数据编码一致填Y，将不会通过对照表的转换，如果没有转换即和U8基础数据编码不一致填N，将会自动通过对照表转换之后，进行相应的操作
            root.SetAttribute("codeexchanged", "N");
            root.SetAttribute("paginate", "0");
            root.SetAttribute("family", GetFamilyStr(tableName));
            #endregion
            //发送方，填外部系统注册码（必填）
            root.SetAttribute("sender", "001");
            var tableNode = doc.CreateElement(tableName);

            //var node = doc.ImportNode(doc.DocumentElement, true);
            //for (int j = 0; j < node.ChildNodes.Count; j++) {
            //    var childNode = node.ChildNodes[j];
            //    tableNode.AppendChild(childNode);
            //    j--;
            //}

            root.AppendChild(tableNode);
            doc.AppendChild(root);
            return doc;
        }

        //public void ExportToXml(List<DataTable> list, string rootName) {
        //    if (list == null || list.Count == 0) {
        //        return;
        //    }
        //    try {
        //        string xlsPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        //        string fPath = Path.GetFullPath("..\\ExportFiles");
        //        string newFile = Path.Combine(fPath, BomExportClient._item.Id + "_" + rootName + ".xml");
        //        var doc = CreateXmlDocument(list, newFile, rootName);
        //        doc.Save(newFile);

        //    } catch (Exception ex) {
        //        PLMEventLog.WriteExceptionLog(ex);
        //    }
        //}

        /// <summary>
        /// 创建xml文档
        /// </summary>
        /// <returns></returns>
        public abstract XmlDocument CreateXmlDocument(string operatorStr);
            //for (int i = 1; i < list.Count; i++) {
            //    var dt = list[i];
            //    dt.WriteXml(newFile);
            //    XmlDocument newDoc = new XmlDocument();
            //    newDoc.Load(newFile);
            //    var node = doc.ImportNode(newDoc.DocumentElement, true);
            //    string path = string.Format("ufinterface//{0}", rootName);
            //    for (int j = 0; j < node.ChildNodes.Count; j++) {
            //        var childNode = node.ChildNodes[j];
            //        doc.SelectSingleNode(path).AppendChild(childNode);
            //        j--;
            //    }
            //} 

        ///// <summary>
        ///// 获取操作字符，如果版本号小于1，则新增，否则修改
        ///// </summary>
        ///// <param name="dEBusinessItem"></param>
        ///// <returns></returns>
        //protected string GetOperatorStr(DEBusinessItem dEBusinessItem) {
        //    if (dEBusinessItem == null) {
        //        return "";
        //    }
        //    var vs = dEBusinessItem.LastRevision;
        //    if (vs <= 1) {
        //        return "Add";
        //    }
        //    return "Edit";
        //}
    }
}
