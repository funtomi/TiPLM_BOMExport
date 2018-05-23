using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Thyt.TiPLM.DEL.Product;

namespace ExportBOMToERP {
    public class UnitGroupDal : BaseDal {
        public UnitGroupDal(DEBusinessItem dItem) {
            _dEBusinessItem = dItem;
            this._name = "unitgroup";
            _filePath = BuildFilePath(dItem, _name);
        }

        protected override XmlDocument BuildXmlDocment(string operatorStr) {
            XmlDocument doc = CreateXmlSchema(_name, _dEBusinessItem, operatorStr);
            DataTable dt = GetDataTable(_dEBusinessItem);
            dt.WriteXml(_filePath);
            XmlDocument docTemp = new XmlDocument();
            docTemp.Load(_filePath);
            var node = doc.ImportNode(docTemp.DocumentElement, true);
            string path = string.Format("ufinterface//{0}", _name);
            var cNode = node.SelectNodes(_name);
            if (cNode == null || cNode.Count == 0) {
                return null;
            }
            for (int j = 0; j < cNode[0].ChildNodes.Count; j++) {
                var childNode = cNode[0].ChildNodes[j];
                doc.SelectSingleNode(path).AppendChild(childNode);
                j--;
            }
            return doc;
        }

        /// <summary>
        /// 获取数据
        /// </summary>
        /// <param name="dEBusinessItem"></param>
        /// <returns></returns>
        private DataTable GetDataTable(DEBusinessItem dEBusinessItem) {
            if (dEBusinessItem==null) {
                return null;
            }
            var dt = BuildElementDt();
            var row = dt.NewRow();
            foreach (DataColumn col in dt.Columns) {
                if (col.ColumnName=="code") {
                    row[col] = dEBusinessItem.Id;
                    continue;
                } 
                var val = dEBusinessItem.GetAttrValue(dEBusinessItem.ClassName, col.ColumnName.ToUpper());
                row[col] = val == null?DBNull.Value:val;
            }
            dt.Rows.Add(row);
            return dt;
        }

        /// <summary>
        /// 构建节点属性结构
        /// </summary>
        /// <returns></returns>
        private DataTable BuildElementDt() {
            DataTable dt = new DataTable(_name);
            dt.Columns.Add("code");//	计量单位组编码
            dt.Columns.Add("name");//	计量单位组名称
            dt.Columns.Add("type",typeof(int));//组类别
            dt.Columns.Add("cgrprelinvcode");//	对应存货编码
            dt.Columns.Add("bdefaultgroup", typeof(int));//	是否默认组
            return dt;
        }
    }
}
