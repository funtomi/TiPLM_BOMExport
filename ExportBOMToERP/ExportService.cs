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
            var eaiAddNode = doc.SelectSingleNode("ERPIntegratorConfig//config");
            if (eaiAddNode == null) {
                MessageBoxPLM.Show("ERP导入配置文件没有EAI地址，请补充！");
                return null;
            }
            var add = eaiAddNode.Value;
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
            string oprt = "Add";
            XmlDocument doc = new XmlDocument();
            bool succeed = true;
            string hasStr = "重复";
            string errText = "";
            switch (_bItem.ClassName.ToLower()) {
                default:
                    break;
                case "unitgroup":
                    UnitGroupDal unitGroupDal = new UnitGroupDal(_bItem);
                    doc = unitGroupDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = unitGroupDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "unit":
                    UnitDal unitDal = new UnitDal(_bItem);
                    doc = unitDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = unitDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "inventoryclass":
                    InventoryClassDal ivtryClassDal = new InventoryClassDal(_bItem);
                    doc = ivtryClassDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = ivtryClassDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "tipart":
                case "tigz":
                    InventoryDal ivtryDal = new InventoryDal(_bItem);
                    doc = ivtryDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = ivtryDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "operation":
                case "gx":
                    OperationDal operationDal = new OperationDal(_bItem);
                    doc = operationDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = operationDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "routing":
                    RoutingDal routingDal = new RoutingDal(_bItem);
                    doc = routingDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = routingDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    if (!succeed) {
                        break;
                    }
                    BomDal bomDal = new BomDal(_bItem);
                    doc = bomDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = bomDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "resourcedoc":
                    ResourceDal resourceDal = new ResourceDal(_bItem);
                    doc = resourceDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = resourceDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;
                case "workcenters":
                    WorkCenterDal workCenterDal = new WorkCenterDal(_bItem);
                    doc = workCenterDal.CreateXmlDocument(oprt);
                    succeed = ConnectEAI(doc.OuterXml, out errText);
                    if (!succeed && errText.Contains(hasStr)) {
                        oprt = "Edit";
                        doc = workCenterDal.CreateXmlDocument(oprt);
                        succeed = ConnectEAI(doc.OuterXml, out errText);
                    }
                    break;

            }
            if (!succeed) {
                MessageBoxPLM.Show(errText);
                return;
            }
            MessageBoxPLM.Show("ERP导入成功!");
        }

        /// <summary>
        /// 导入到ERP
        /// </summary>
        /// <param name="xml"></param>
        private bool ConnectEAI(string xml, out string errText) {
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
