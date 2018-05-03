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
using Thyt.TiPLM.DEL.Admin.DataModel;
using Thyt.TiPLM.DEL.Product;
using Thyt.TiPLM.PLL.Admin.DataModel;
using Thyt.TiPLM.PLL.Common;
using Thyt.TiPLM.PLL.Environment;
using Thyt.TiPLM.PLL.Product2;
using Thyt.TiPLM.UIL.Common;
using Thyt.TiPLM.UIL.Product.Common;

namespace BOMExportClient {
    public class BomExportClient : IAddinClientEntry, IAutoExec {
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

        /// <summary>
        /// 执行导出
        /// </summary>
        /// <param name="Iitem"></param>
        private void ExportExecute(IBizItem Iitem) {
            IActionContext context = ActionContext2.GetInstance().GetCurrentContext();
            DEPSOption option = null;
            if (context != null && context.NavContext.Option != null) {
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
            if (bItem.ClassName!="TIGZ"&&bItem.ClassName!="TIPART") {//只有物料才导出
                return;
            }
            item = bItem;
            var ds = BuildBomDataSet(item);
            //Guid eventOid = Guid.NewGuid();
            //expEvent = new DEExportEvent(eventOid, item.MasterOid, item.ClassName, item.RevNum, ClientData.LogonUser.Oid, DateTime.Now, "产品结构导出", option);
            ExportXml.ExportToXml(ds,"Bom");
            //if (Execute()) {
            //    PLMEventLog.WriteLog("数据导出完成", EventLogEntryType.Information);
            //}
        }

        /// <summary>
        /// 建立BOM节点
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private List<DataTable> BuildBomDataSet(DEBusinessItem item) {
            List<DataTable> ds = new List<DataTable>();
            if (item==null) {
                return null;
            }
            #region Version节点

            DataTable versionDt = GetVersionDataTable();
            var versionRow = versionDt.NewRow();
            versionRow["BomId"] = item.Id;
            versionRow["BomType"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
            versionRow["Version"] = item.LastRevision;
            versionRow["Status"] = item.GetAttrValue(item.ClassName, "INSSTATE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "INSSTATE");
            versionRow["ModifyTime"] = item.GetAttrValue(item.ClassName, "CHECKINTIME") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "CHECKINTIME");
            versionRow["VersionDesc"] = item.GetAttrValue(item.ClassName, "VERSIONDESC") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "VERSIONDESC");
            versionRow["VersionEffDate"] = item.GetAttrValue(item.ClassName, "VERSIONEFFDATE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "VERSIONEFFDATE");
            #region 待补充字段
            //versionRow["IdentCode"] = item.GetAttrValue(item.ClassName, "REVISIONOID");
            //versionRow["IdentDesc"] = item.GetAttrValue(item.ClassName, "REVISIONOID");
            //versionRow["ApplyCode"] = item.GetAttrValue(item.ClassName, "REVISIONOID");
            //versionRow["ApplySeq"] = item.GetAttrValue(item.ClassName, "REVISIONOID");
            #endregion
            versionDt.Rows.Add(versionRow);
            ds.Add(versionDt);
            #endregion

            #region Parent节点
            DataTable parenttDt = GetParentDataTable();
            var parentRow = parenttDt.NewRow();
            parentRow["BomId"] = item.Id;
            parentRow["InvCode"] = item.GetAttrValue(item.ClassName, "OID") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "OID");
            parentRow["InvName"] = item.GetAttrValue(item.ClassName, "NAME") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "NAME");
            parentRow["ParentScrap"] = item.GetAttrValue(item.ClassName, "PARENTSCRAP") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "PARENTSCRAP");
            #region 待补充字段
            //parentRow["InvAddCode"] = item.GetAttrValue(item.ClassName, "BOMTYPE");
            //parentRow["EngineerFigNo"] = item.GetAttrValue(item.ClassName, "BOMTYPE");
            //parentRow["ParentScrap"] = item.GetAttrValue(item.ClassName, "BOMTYPE");
            #endregion
            parenttDt.Rows.Add(parentRow);
            ds.Add(parenttDt);
            #endregion

            #region Component节点
            DataTable componentDt = GetComponentDataTable(); 
            ds = GetComponnetDataSet(item,ds); 
            #endregion
            return ds;
        }

        /// <summary>
        /// 获取BOM信息
        /// </summary>
        /// <param name="item"></param>
        /// <param name="parentId"></param>
        /// <returns></returns>
        private List<DataTable> GetComponnetDataSet(DEBusinessItem item, List<DataTable> ds) {
            DERelationBizItemList relationBizItemList = item.Iteration.LinkRelationSet.GetRelationBizItemList("PARTTOPART");
            if (relationBizItemList == null) {
                try {
                    relationBizItemList = PLItem.Agent.GetLinkRelationItems(item.Iteration.Oid, item.Master.ClassName, "PARTTOPART", ClientData.LogonUser.Oid, ClientData.UserGlobalOption);
                } catch {
                    return null;
                }
            }
            if (relationBizItemList==null||relationBizItemList.Count==0) {
                return null;
            }
            for (int i = 0; i < relationBizItemList.BizItems.Count; i++) {
                var iItem = relationBizItemList.BizItems[i] as DEBusinessItem;
                if (iItem==null) {
                    continue;
                }
                var dt = GetComponentDataTable();
                var row = dt.NewRow();
                row["BomId"]=item.Id;
                row["OpComponentId"]=iItem.Id;
                row["BomType"] = iItem.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : iItem.GetAttrValue(item.ClassName, "BOMTYPE");
                row["SortSeq"] = iItem.GetAttrValue(item.ClassName, "ORDER") == null ? DBNull.Value : iItem.GetAttrValue(item.ClassName, "ORDER");
                row["OpSeq"] = iItem.GetAttrValue(item.ClassName, "OPLINE") == null ? DBNull.Value : iItem.GetAttrValue(item.ClassName, "OPLINE");
                row["InvCode"] = iItem.Id;
                row["Version"] = iItem.LastRevision;
                row["Status"] = iItem.GetAttrValue(item.ClassName, "INSSTATE") == null ? DBNull.Value : iItem.GetAttrValue(item.ClassName, "INSSTATE");
                row["ModifyTime"] = iItem.GetAttrValue(item.ClassName, "CHECKINTIME") == null ? DBNull.Value : iItem.GetAttrValue(item.ClassName, "CHECKINTIME");
                row["EffEndDate"] = item.GetAttrValue(item.ClassName, "VERSIONENDDATE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "VERSIONENDDATE");
                row["EffBegDate"] = item.GetAttrValue(item.ClassName, "VERSIONEFFDATE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "VERSIONEFFDATE");
                row["CompScrap"] = item.GetAttrValue(item.ClassName, "PARENTSCRAP") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "PARENTSCRAP");
                
                #region 待填的属性

                //row["VersionDesc"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
                //row["VersionEffDate"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
                //row["IdentCode"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
                //row["IdentDesc"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
                //row["ApplyCode"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
                //row["ApplySeq"] = item.GetAttrValue(item.ClassName, "BOMTYPE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "BOMTYPE");
                #endregion
                dt.Rows.Add(row);
                ds.Add(dt);
            }
            return ds;
        }

        private DataTable GetParentDataTable() {
            DataTable dt = new DataTable("Parent");
            dt.Columns.Add("BomId");//BomId
            dt.Columns.Add("InvCode");//母件编码
            dt.Columns.Add("InvName");//母件名称
            dt.Columns.Add("InvAddCode");//
            dt.Columns.Add("EngineerFigNo");//
            dt.Columns.Add("ParentScrap", typeof(decimal));//母件损耗率 
            return dt;
        }

        private DataTable GetComponentDataTable() {
            DataTable dt = new DataTable("Component");
            dt.Columns.Add("OpComponentId");//BOMID
            dt.Columns.Add("BomId");//parentId
            dt.Columns.Add("SortSeq", typeof(int));//子件行号
            dt.Columns.Add("OpSeq", typeof(int));//工序行号
            dt.Columns.Add("InvCode");//子件编码
            dt.Columns.Add("CompScrap",typeof(decimal));//子件损耗率
            dt.Columns.Add("BomType");//BOM类型（1.主要，2.替代）
            dt.Columns.Add("Version", typeof(int));//版本号 
            dt.Columns.Add("VersionDesc");//版本说明 
            dt.Columns.Add("EffBegDate", typeof(DateTime));//版本生效日 
            dt.Columns.Add("EffEndDate", typeof(DateTime));//版本失效日 
            dt.Columns.Add("IdentCode");//替代标识 
            dt.Columns.Add("IdentDesc");//替代说明 
            dt.Columns.Add("ApplyCode");//变更单号
            dt.Columns.Add("ApplySeq");//变更单行号
            dt.Columns.Add("Status");//Status 状态(1:新建/3:审核/4:停用)
            dt.Columns.Add("ModifyTime", typeof(DateTime));//修改日期  
            return dt;
        }

        /// <summary>
        /// 构建Version节点结构
        /// </summary>
        /// <returns></returns>
        private DataTable GetVersionDataTable() {
            DataTable dt = new DataTable("Version");
            dt.Columns.Add("BomId");//BOMID
            dt.Columns.Add("BomType");//BOM类型（1.主要，2.替代）
            dt.Columns.Add("Version",typeof(int));//版本号 
            dt.Columns.Add("VersionDesc");//版本说明 
            dt.Columns.Add("VersionEffDate",typeof(DateTime));//版本失效日 
            dt.Columns.Add("IdentCode");//替代标识 
            dt.Columns.Add("IdentDesc");//替代说明 
            dt.Columns.Add("ApplyCode");//变更单号
            dt.Columns.Add("ApplySeq");//变更单行号
            dt.Columns.Add("Status");//Status 状态(1:新建/3:审核/4:停用)
            dt.Columns.Add("ModifyTime",typeof(DateTime));//修改日期 
            return dt;
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
                    List<DataSet> ds = null;
                    Object inObj = null;
                    Object OutObj;
                    bool suc = agent.ExportPSToERP(expEvent, useDe, out strErr, out strWarn, out ds, inObj,
                                                   out OutObj);

                    if (strWarn != null && strWarn.Length > 0)
                        et = new Exception(strWarn.ToString());
                    if (suc) {
                        //ExportXml.ExportToXml(ds, expEvent.Oid, ClientData.LogonUser.Oid);

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

        private void SetOption(DEErpExport useDE) {
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
            string isExpHistory = "Y";
            root.SetAttribute("NoExportHistory", isExpHistory);
            root.SetAttribute("MaxLevel", PLSystemParam.ParameterPartMaxLevel.ToString());//默认导出层数最大
            return root.OuterXml;
        }
    }
}
