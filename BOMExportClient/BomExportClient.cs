using BOMExportCommon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Thyt.TiPLM.Common;
using Thyt.TiPLM.Common.Interface.Addin;
using Thyt.TiPLM.Common.Interface.Admin.DataModel;
using Thyt.TiPLM.DEL.Product;
using Thyt.TiPLM.PLL.Common;
using Thyt.TiPLM.PLL.Environment;
using Thyt.TiPLM.PLL.Product2;
using Thyt.TiPLM.UIL.Common;
using Thyt.TiPLM.UIL.Product.Common;

namespace BOMExportClient
{
    public class BomExportClient :IAddinClientEntry, IAutoExec
    {
        public Delegate d_AfterReleased { get; set; }
        public static DEBusinessItem item = null;
        public static DEExportEvent expEvent = null;
        XmlDocument doc = new XmlDocument(); 
        public void Init() {
            this.d_AfterReleased = new PLMBizItemDelegate(AfterItemReleased);
            BizItemHandlerEvent.Instance.D_AfterReleased = (PLMBizItemDelegate)Delegate.Combine(BizItemHandlerEvent.Instance.D_AfterReleased, this.d_AfterReleased);
        }

        public void UnInit() {
            BizItemHandlerEvent.Instance.D_AfterReleased = (PLMBizItemDelegate)Delegate.Remove(BizItemHandlerEvent.Instance.D_AfterReleased, this.d_AfterReleased);

        }

        /// <summary>
        /// 定版后启动
        /// </summary>
        /// <param name="bizItems"></param>
        private void AfterItemReleased(IBizItem[] bizItems) {
            if (bizItems != null) {
                ArrayList list = new ArrayList(bizItems);
                foreach (object obj2 in list) {
                    if (typeof(IBizItem).IsInstanceOfType(obj2)) {
                        IBizItem item = (IBizItem)obj2;
                        ExportExecute(item);
                        //DEPSOption option = PSStart.BuildLocalPSOptionByRev(item, ClientData.UserGlobalOption, ClientData.LogonUser.Oid);
                        //DEExportEvent expEvent = new DEExportEvent(Guid.NewGuid(), item.MasterOid, item.ClassName, item.Revision.Revision, ClientData.LogonUser.Oid, DateTime.Now, "", option);
                        //FrmBOMExport.ExportPS((DEBusinessItem)item, option);
                    }
                }
            }
        }

        private void ExportExecute(IBizItem Iitem) {
            IActionContext context = ActionContext2.GetInstance().GetCurrentContext();
            DEPSOption option = null;
            if (context !=null&&context.NavContext.Option!=null) {
                option = context.NavContext.Option;
            }
            if (option == null)
                option = ClientData.UserGlobalOption;  
            if (Iitem == null)
                return;
            DEBusinessItem bItem = Iitem as DEBusinessItem;
            if (bItem == null && Iitem is DESmartBizItem) {
                DESmartBizItem sb = Iitem as DESmartBizItem;
                bItem = (DEBusinessItem)PLItem.Agent.GetBizItem(sb.MasterOid, sb.RevOid, sb.IterOid,
                                                                 option.CurView, ClientData.LogonUser.Oid,
                                                                 BizItemMode.BizItem);
            }
            if (bItem == null) return;
            item = bItem;
            Guid eventOid = Guid.NewGuid();
            expEvent = new DEExportEvent(eventOid, item.MasterOid, item.ClassName, item.RevNum, ClientData.LogonUser.Oid, DateTime.Now, "产品结构导出", option);
            if (Execute()) {
                PLMEventLog.WriteLog("数据导出完成", EventLogEntryType.Information);
            }
        }

        private bool Execute() {
            Interface agent = RemoteProxy.GetObject(typeof(Interface), ConstERP.RemotingURL) as Interface;
            ArrayList lst;
            Exception et = null;
            try {
                lst = agent.GetExpProe(ClientData.LogonUser.Oid); 
                if (lst == null || lst.Count == 0) {
                    PLMEventLog.WriteLog("没有找到允许使用的ERP导出插件，可能用户没有权限，请与管理员联系", EventLogEntryType.Warning);
                    return false;
                }
                foreach (DEErpExport useDe in lst) {
                    useDe.IsExpOldItemAndBom = true;//不导出重复的物料信息 


                    expEvent.ExpOption.MaxLevel = PLSystemParam.ParameterPartMaxLevel;
                    string xmlPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    string DllPath = null;
                    #region todo 导出前功能，目前还没找到代码
                    //if (!String.IsNullOrEmpty(useDe.ExtendCLTDllName)) {
                    //    DllPath = Path.Combine(xmlPath, useDe.ExtendCLTDllName);
                    //    string ext = Path.GetExtension(DllPath);
                    //    if (ext.ToUpper() != ".DLL")
                    //        DllPath += ".dll";
                    //    if (File.Exists(DllPath)) {
                    //        inObj = ExtendBreforeExport(expEvent.Oid, ClientData.LogonUser.Oid, DllPath);
                    //    } else {
                    //        MessageBox.Show("配置文件中指定的客户端扩展处理文件不存在，无法继续");
                    //        return false;
                    //    }
                    //}
                    #endregion

                    StringBuilder strErr, strWarn;
                    DataSet ds = null;
                    Object inObj = null;
                    Object OutObj;
                    bool suc = agent.ExportPSToERP(expEvent, useDe, out strErr, out strWarn, out ds, inObj,
                                                   out OutObj);
                    if (ds != null && ds.Tables.Contains("ErrInfo")) {
                        ds.Tables.Remove(ds.Tables["ErrInfo"]);
                        ds.AcceptChanges();
                    }
                    if (strWarn != null && strWarn.Length > 0)
                        et = new Exception(strWarn.ToString());
                    if (suc) {
                        if (useDe.op == ExpOption.eXls) {
                            ExportExcel.ExportToExcel(ds, expEvent.Oid, ClientData.LogonUser.Oid);
                        } else if (useDe.op == ExpOption.eDataBase) {
                            string curPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                            string fileName = curPath + "\\DataPro.cmd";
                            if (File.Exists(fileName)) {
                                Process pr = new Process();
                                pr.StartInfo.FileName = fileName;
                                try {
                                    pr.Start();
                                } catch (Exception ex) {
                                    PLMEventLog.WriteLog("数据导入成功，数据已经导入到中间表，但调用客户端后处理程序: DataPro.cmd 失败", EventLogEntryType.FailureAudit);
                                }
                            }
                        }
                        //added by kexp in 20018/04/26
                    }
                    #region 导出后处理程序 todo
                    //    if (!String.IsNullOrEmpty(DllPath))
                    //        ExtendAfterExport(ds, OutObj, DllPath);
                    //    if (et != null) {
                    //        FrmErr frm = new FrmErr();
                    //        if (ClientData.mainForm != null)
                    //            frm.MdiParent = ClientData.mainForm;
                    //        // frm.ShowErrInfo(null, et,tbErr);
                    //        frm.ShowErrInfo(null, strWarn, tbErr);
                    //        frm.Show();
                    //    }
                    //} else {
                    //    if (strErr != null && strErr.Length > 0)
                    //        throw new Exception(strErr.ToString());
                    //    if (tbErr != null)
                    //        throw new Exception("发现数据错误");
                    //}
                    #endregion
                }
                return true;
            } catch (Exception ex) {
                PLMEventLog.WriteExceptionLog(ex);

                return false;
            }
        }
        
        private  void SetOption(DEErpExport useDE) {
            string opt = GetOptionString(useDE);
            ClientData.SetUserOption("ERPExportOption2", opt);
        }

        /// <summary>
        /// 获取配置信息
        /// </summary>
        /// <param name="de"></param>
        /// <returns></returns>
        private string GetOptionString(DEErpExport de) {
            XmlElement root = doc.CreateElement("ERPExportOption");
            root.SetAttribute("ExportName", de.ExpProeName);
            string strTypes = "";
            for (int i = 0; i < de.lstExportDataType.Count; i++) {
                string t = de.lstExportDataType[i].ToString();
                strTypes += t;
                if (i < de.lstExportDataType.Count - 1)
                    strTypes += ",";
            }
            root.SetAttribute("ExportDataType", strTypes);
            string isExpHistory = "Y" ;
            root.SetAttribute("NoExportHistory", isExpHistory);
            root.SetAttribute("MaxLevel", PLSystemParam.ParameterPartMaxLevel.ToString());//默认导出层数最大
            return root.OuterXml;
        }
    }
}
