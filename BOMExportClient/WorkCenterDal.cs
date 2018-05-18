using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Thyt.TiPLM.DEL.Product;

namespace BOMExportClient {
    public class WorkCenterDal : BaseDal {
        public WorkCenterDal(DEBusinessItem dItem) {
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
            if (dEBusinessItem == null) {
                return null;
            }
            var dt = BuildElementDt();
            var row = dt.NewRow();
            foreach (DataColumn col in dt.Columns) {
                if (col.ColumnName == "WcCode") {
                    row[col] = dEBusinessItem.Id;
                    continue;
                }
                if (col.ColumnName == "Description") {
                    row[col] = dEBusinessItem.Name;
                    continue;
                }
                if (col.ColumnName == "CalendarCode") {
                    row[col] = "SYSTEM";
                    continue;
                }
                var val = dEBusinessItem.GetAttrValue(dEBusinessItem.ClassName, col.ColumnName.ToUpper());
                row[col] = val == null ? DBNull.Value : val;
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
            //dt.Columns.Add("WcId	");//工作中心Id（自动编号）
            dt.Columns.Add("WcCode");//	工作中心代号  
            dt.Columns.Add("Description");//	名称  
            dt.Columns.Add("CalendarCode");//	行事历Id  （工作日历代号）
            dt.Columns.Add("ProductLineFlag");//	是否生产线  
            dt.Columns.Add("DeptCode	");//隶属部门
            return dt;
        }
    }
}
