using BOMExportCommon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Thyt.TiPLM.BRL.Common;
using Thyt.TiPLM.Common;
using Thyt.TiPLM.Common.Interface.Addin;
using Thyt.TiPLM.DAL.Admin.NewResponsibility;
using Thyt.TiPLM.DAL.Common;
using Thyt.TiPLM.DEL.Admin.NewResponsibility;
using Thyt.TiPLM.DEL.Product;

namespace BOMExportServer {
    /// <summary>
    /// Class1 的摘要说明。
    /// </summary>
    public class BFEntrance : PLMBFRoot, IAddinServiceEntry,
                              Interface {
        #region IAddinServiceEntry 成员

        public WellKnownServiceTypeEntry[] RemoteTypes {
            get {
                WellKnownServiceTypeEntry[] entries = new WellKnownServiceTypeEntry[1];
                entries[0] = new WellKnownServiceTypeEntry(
                    typeof(BFEntrance),
                    ConstERP.RemotingURL,
                    WellKnownObjectMode.SingleCall);
                return entries;
            }
        }

        #endregion

        #region Interface Members

        /// <summary>
        /// 导出产品结构到ERP系统中
        /// </summary>
        /// <param name="eventOid">事件标识</param>
        /// <param name="userOid">用户标识</param>
        /// <param name="deErp">导出在配置</param>
        /// <param name="strErrInfo">出错信息</param>
        /// <param name="ds">获取数据（可能为空，用于客户端需要展现或保存数据）</param>
        /// <param name="InObjs">对于需要进行二次开发的企业，需要携带</param>
        /// <param name="outObjs">需要返回在扩展插件中进行二次开发处理的数据</param>
        /// <returns>如果成功，返回true；否则返回false</returns>
        /// <remarks>
        /// </remarks>
        public bool ExportPSToERP(DEExportEvent expEvent, DEErpExport deErp, out StringBuilder strErrInfo,
            out StringBuilder strWarnInfo, out DataSet ds, object InObjs, out object outObjs) {
            DBParameter dbParam = DBUtil.GetDbParameter(true);
            try {
                dbParam.Open();
                DAERPExport da = new DAERPExport(dbParam);
                bool succ = da.ExportPSToERP(expEvent, out strErrInfo, out strWarnInfo, deErp, out ds, InObjs, out outObjs);
                dbParam.Commit();
                return succ;
            } catch (Exception ex) {
                dbParam.Rollback();
                throw ex;
            } finally {
                dbParam.Close();
                if (AddinDelegate.Handler.PLMAddin_AfterProcessFinished != null)
                    AddinDelegate.Handler.PLMAddin_AfterProcessFinished(expEvent.Oid, expEvent.Exportor);
            }
        }

        /// <summary>
        /// 获取可用ERP导出程序
        /// </summary>
        /// <param name="userOid"></param>
        /// <returns></returns>
        public ArrayList GetExpProe(Guid userOid) {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            string[] cfgs = Directory.GetFiles(dir, "*.config");
            if (cfgs == null || cfgs.Length == 0)
                throw new Exception(string.Format("config配置文件，没有部署在服务器端,路径：{0}",Path.GetFullPath(dir)));
            XmlDocument doc = new XmlDocument();
            ArrayList lst = new ArrayList();
            DBParameter dbParam = DBUtil.GetDbParameter(false);
            try {
                dbParam.Open();
                DAUser daUser = new DAUser(dbParam);
                DEUser user = daUser.GetByOid(userOid);
                bool isFind;
                foreach (string cfg in cfgs) {
                    string cfgName = Path.GetFileNameWithoutExtension(cfg);
                    try {
                        doc.Load(cfg);
                    } catch (Exception ex) {

                        throw new Exception("配置文件《" + cfgName + "》存在错误：" + ex.Message, ex.InnerException);
                    }

                    XmlElement root = doc.DocumentElement;

                    //获取人员AuthorizedUsers
                    isFind = false;
                    if (user.LogId.ToLower() == "sysadmin") {
                        isFind = true;
                    } else {
                        XmlNode als = root.SelectSingleNode("AuthorizedUsers");
                        if (als != null) {
                            foreach (XmlElement xuser in als.ChildNodes) {
                                string uId = xuser.InnerText.Trim();
                                if (user.LogId.ToUpper() == uId.ToUpper()) {
                                    isFind = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (!isFind) continue;
                    XmlElement ErpInfo = (XmlElement)root.SelectSingleNode("ERPExportInfo");
                    DEErpExport deErp = new DEErpExport();

                    deErp.lstExportDataType = new ArrayList();
                    deErp.configName = Path.GetFileName(cfg);
                    if (ErpInfo == null) {
                        deErp.ExpProeName = Path.GetFileNameWithoutExtension(cfg);
                        deErp.op = ExpOption.eDataBase;
                        deErp.dec = "";
                        if (File.Exists(Path.Combine(dir, "ExtendERPExport.dll"))) {
                            deErp.ExtendSVRDllName = "ExtendERPExport.dll";
                        }
                        deErp.ExtendCLTDllName = "";
                    } else {
                        foreach (XmlNode nd in ErpInfo.ChildNodes) {
                            if (nd.Name.ToUpper() == "ERPExportName".ToUpper().Trim())
                                deErp.ExpProeName = nd.InnerText;
                            if (nd.Name.ToUpper() == "Desc".ToUpper().Trim())
                                deErp.dec = nd.InnerText;
                            if (nd.Name.ToUpper() == "ExportType".ToUpper().Trim()) {
                                string strOp = nd.InnerText;
                                if (strOp.ToUpper() == "Other".ToUpper())
                                    deErp.op = ExpOption.eOther;
                                else if (strOp.ToUpper() == "Excel".ToUpper())
                                    deErp.op = ExpOption.eXls;
                                else
                                    deErp.op = ExpOption.eDataBase;
                            }
                            if (nd.Name.ToUpper() == "ExtendSVRDllName".ToUpper().Trim())
                                deErp.ExtendSVRDllName = nd.InnerText;
                            if (nd.Name.ToUpper() == "ExtendCLTDllName".ToUpper().Trim())
                                deErp.ExtendCLTDllName = nd.InnerText;
                        }
                        if (String.IsNullOrEmpty(deErp.ExpProeName))
                            deErp.ExpProeName = Path.GetFileNameWithoutExtension(cfg);
                        if (String.IsNullOrEmpty(deErp.ExtendSVRDllName))
                            deErp.ExtendSVRDllName = "ExtendERPExport.dll";
                    }
                    XmlElement xmlData = (XmlElement)root.SelectSingleNode("DataMapping");

                    string c;
                    foreach (XmlNode nd in xmlData.ChildNodes) {
                        c = "";
                        if (nd.Name.ToUpper() == "ItemColumnMapping".ToUpper()) {
                            c = "导出物料";
                        }
                        if (nd.Name.ToUpper() == "BomColumnMapping".ToUpper()) {
                            c = "导出BOM";
                        }
                        if (nd.Name.ToUpper() == "RoutingMapping".ToUpper()) {
                            if (nd.Attributes["Lable"] == null)
                                c = "导出工艺";
                            else if (nd.Attributes["Lable"] != null) {
                                XmlAttribute attribute = nd.Attributes["Lable"];
                                c = attribute.Value;
                            }
                        }
                        if (!string.IsNullOrEmpty(c) && !deErp.lstExportDataType.Contains(c))
                            deErp.lstExportDataType.Add(c);
                    }
                    lst.Add(deErp);
                }
                return lst;
            } catch (Exception ex) {
                throw new PLMException("获取可用ERP导出配置文件出错", ex);
            } finally {
                dbParam.Close();
            }
        }

        #endregion
    }
}
