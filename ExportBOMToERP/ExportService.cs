using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Xml;
using Thyt.TiPLM.Common;
using Thyt.TiPLM.DEL.Product;
using Thyt.TiPLM.UIL.Common;
using Thyt.TiPLM.UIL.Controls;

namespace ExportBOMToERP {
    public class ExportService {
        public ExportService(DEBusinessItem item) {
            _bItem = item;
        }
        private DEBusinessItem _bItem;
        private string _eaiAddress;

        protected string EaiAddress {
            get {
                if (string.IsNullOrEmpty(_eaiAddress)) {
                    _eaiAddress = GetEAIConfig();
                }
                return _eaiAddress;
            }
        }

        /// <summary>
        /// 获取可用ERP导出程序
        /// </summary>
        /// <param name="userOid"></param>
        /// <returns></returns>
        public string GetEAIConfig() {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            string[] cfgs = Directory.GetFiles(dir, "*config.xml");
            if (cfgs == null || cfgs.Length == 0) {
                MessageBoxPLM.Show("ERP导入配置没有配置在客户端！");
                return null;
            }
            return CheckEAIConfig(cfgs[0]);
        }

        /// <summary>
        /// 检查配置文件合法性
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string CheckEAIConfig(string fileName) {
            if (fileName == null) {
                MessageBoxPLM.Show("ERP导入配置文件没有内容！");
                return null;
            }
            if (!fileName.EndsWith(".xml")) {
                MessageBoxPLM.Show("ERP导入配置文件格式不正确，只支持xml格式的配置！");
                return null;
            }
            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);
            var eaiAddNode = doc.SelectSingleNode("ERPIntegratorConfig//config//EAIAddress");
            if (eaiAddNode == null) {
                MessageBoxPLM.Show("ERP导入配置文件没有EAI地址，请补充！");
                return null;
            }
            var add = eaiAddNode.InnerText;
            if (string.IsNullOrEmpty(add)) {
                MessageBoxPLM.Show("ERP导入配置文件没有EAI地址，请补充！");
                return null;
            }
            return add;
        }
        #region 导出导入
        public void AddOrEditItem() {
            if (_bItem == null) {
                MessageBoxPLM.Show("没有获取到当前对象，请检查！");
                return;
            }
            //string oprt = item.LastRevision > 1 ? "Edit" : "Add";
            bool succeed;
            string errText;

            succeed = DoExport(out errText, _bItem);
            if (!succeed && !string.IsNullOrEmpty(errText)) {
                MessageBoxPLM.Show(errText);
                return;
            }
            MessageBoxPLM.Show("ERP导入成功!");
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="oprt"></param>
        /// <param name="errText"></param>
        /// <param name="bItem"></param>
        private bool DoExport(out string errText, DEBusinessItem bItem) {
            var exportResult = DoExportImplement(bItem, bItem.ClassName, out errText);
            if (!exportResult) {
                return false;
            }
            if (_bItem.ClassName.ToLower() == "gygck") {
                return DoExportImplement(bItem, "Bom", out errText);
            }

            return true;
        }

        private bool DoExportImplement(DEBusinessItem bItem, string typeStr, out string errText) {
            XmlDocument doc = new XmlDocument();
            bool succeed = true;
            errText = "";
            string hasStr = "重复";
            string oprt = "Add";
            var dal = DalFactory.Instance.CreateDal(bItem, typeStr);
            doc = dal.CreateXmlDocument(oprt);
            if (doc==null) {
                return false;
            }
            succeed = ConnectEAI(doc.OuterXml, out errText);
            if (!succeed && errText.Contains(hasStr)) {
                oprt = "Edit";
                doc = dal.CreateXmlDocument(oprt);
                succeed = ConnectEAI(doc.OuterXml, out errText);
            }
            return succeed;
        }

        /// <summary>
        /// 导入到ERP
        /// </summary>
        /// <param name="xml"></param>
        private bool ConnectEAI(string xml, out string errText) {
            if (string.IsNullOrEmpty(xml)) {
                errText = "";
                return false;
            }
            errText = "";
            MSXML2.XMLHTTPClass xmlHttp = new MSXML2.XMLHTTPClass();
            xmlHttp.open("POST", EaiAddress, false, null, null);//TODO：地址需要改
            xmlHttp.send(xml);
            String responseXml = xmlHttp.responseText;
            //…… //处理返回结果 
            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml(responseXml);
            var itemNode = resultDoc.SelectSingleNode("ufinterface//item");
            var s = ConstCommon.CURRENT_PRODUCTNAME;

            if (itemNode == null) {
                errText = "没有收到ERP回执";

                PLMEventLog.WriteLog("没有收到ERP回执！", EventLogEntryType.Error);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp); //COM释放
                return false;
            }
            var succeed = Convert.ToInt32(itemNode.Attributes["succeed"].Value);//成功标识：0：成功；非0：失败；
            var dsc = itemNode.Attributes["dsc"].Value.ToString();
            //var u8key =itemNode.Attributes["u8key"].ToString();
            //var proc = itemNode.Attributes["proc"].ToString();
            if (succeed != 0) {
                errText = string.Format("ERP导入失败，原因：{0}", dsc);

                PLMEventLog.WriteLog(dsc, EventLogEntryType.Error);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp); //COM释放
                return false;

            }
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp); //COM释放
            return true;
        }
        #endregion
    }
}
