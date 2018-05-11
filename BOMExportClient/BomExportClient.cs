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
        public Delegate d_AfterDeleted { get; set; }
        public static DEBusinessItem _item = null;
        private int _seq = 0;

        /// <summary>
        /// 获取当前对象
        /// </summary>
        /// <param name="iItem"></param>
        /// <returns></returns>
        private DEBusinessItem GetDEBusinessItem(IBizItem iItem) {
            IActionContext context = ActionContext2.GetInstance().GetCurrentContext();
            DEPSOption option = null;
            if (context != null && context.NavContext.Option != null) {
                option = context.NavContext.Option;
            }
            if (option == null)
                option = ClientData.UserGlobalOption;
            if (iItem == null)
                return null;
            DEBusinessItem bItem = iItem as DEBusinessItem;

            if (bItem == null && iItem is DESmartBizItem) {
                DESmartBizItem sb = iItem as DESmartBizItem;
                bItem = (DEBusinessItem)PLItem.Agent.GetBizItem(sb.MasterOid, sb.RevOid, sb.IterOid,
                                                                 option.CurView, ClientData.LogonUser.Oid,
                                                                 BizItemMode.BizItem);
            }
            return bItem;
        }

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
            if (bItem.ClassName != "TIGZ" && bItem.ClassName != "TIPART") {//只有物料才导出
                return;
            }
            _item = bItem;
            #region BOM

            #endregion
            var ds = BuildBomDataSet(_item);
            ExportXml.ExportToXml(ds, "Bom");
            ds = BuildOperationDataSet(_item);
            ExportXml.ExportToXml(ds, "Operation");
            ds = BuildRoutingDataSet(_item);
            ExportXml.ExportToXml(ds, "Routing");
        }
        #region 构建工艺路线结构

        private List<DataTable> BuildRoutingDataSet(DEBusinessItem rootItem) {
            List<DataTable> ds = new List<DataTable>();
            if (rootItem == null) {
                return null;
            }
            #region Version节点
            DataTable versionDt = GetVersionDataTable();
            var versionRow = versionDt.NewRow();
            versionRow["PRoutingId"] = rootItem.Id;
            versionRow["RountingType"] = 1;
            versionRow["Version"] = rootItem.LastRevision;
            versionRow["Status"] = rootItem.GetAttrValue(rootItem.ClassName, "INSSTATE") == null ? DBNull.Value : rootItem.GetAttrValue(rootItem.ClassName, "INSSTATE");
            versionRow["RunCardFlag"] = 3;
            versionRow["RunCardFlag"] = 0.000000;
            versionRow["VersionDesc"] = rootItem.GetAttrValue(rootItem.ClassName, "VERSIONDESC") == null ? DBNull.Value : rootItem.GetAttrValue(rootItem.ClassName, "VERSIONDESC");
            versionRow["VersionEffDate"] = rootItem.GetAttrValue(rootItem.ClassName, "VERSIONEFFDATE") == null ? DBNull.Value : rootItem.GetAttrValue(rootItem.ClassName, "VERSIONEFFDATE");
            versionRow["VersionEndDate"] = rootItem.GetAttrValue(rootItem.ClassName, "VERSIONENDDATE") == null ? DBNull.Value : rootItem.GetAttrValue(rootItem.ClassName, "VERSIONENDDATE");
            versionDt.Rows.Add(versionRow);
            ds.Add(versionDt);
            #endregion

            #region Part节点
            DataTable partDt = GetPartDataTable();
            var part = partDt.NewRow();
            part["PRoutingId"] = rootItem.Id;
            part["InvCode"] = rootItem.GetAttrValue(rootItem.ClassName, "OID") == null ? DBNull.Value : rootItem.GetAttrValue(rootItem.ClassName, "OID");
            partDt.Rows.Add(part);
            ds.Add(partDt);
            #endregion

            #region RoutingDetail节点
            ds = BuildRoutingDetailDataSet(rootItem, ds);
            #endregion

            return ds;
        }

        private List<DataTable> BuildRoutingDetailDataSet(DEBusinessItem rootItem, List<DataTable> ds) {
            if (rootItem == null) {
                return ds;
            }
            if (ds == null) {
                ds = new List<DataTable>();
            }

            DataTable detailDt = GetRoutingDetailDataTable();
            var items = GetLinks(rootItem, "PARTTOGX");
            if (items == null || items.Count == 0) {
                return ds;
            }
            for (int i = 0; i < items.BizItems.Count; i++) {
                #region OperationDetail节点
                var item = items.BizItems[i] as DEBusinessItem;
                var relation = items.RelationList[i] as DERelation2;
                if (item == null) {
                    continue;
                }
                var detailRow = detailDt.NewRow();
                detailRow["PRoutingDId"] = rootItem.Id;
                detailRow["PRoutingId"] = item.Id;
                detailRow["OpSeq"] = relation.GetAttrValue("OPLINE") == null ? DBNull.Value : relation.GetAttrValue("OPLINE");
                detailRow["OperationCode"] = item.GetAttrValue(item.ClassName, "OID") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "OID");
                detailRow["Description"] = item.GetAttrValue(item.ClassName, "GXSM") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "GXSM");
                detailRow["WcCode"] = item.GetAttrValue(item.ClassName, "WCCODE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "WCCODE");
                detailRow["EffBegDate"] = item.GetAttrValue(item.ClassName, "EFFBEGDATE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "EFFBEGDATE");
                detailRow["EffEndDate"] = item.GetAttrValue(item.ClassName, "EFFENDDATE") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "EFFENDDATE");
                detailRow["SubFlag"] = item.GetAttrValue(item.ClassName, "WWGX") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "WWGX");
                detailRow["OpSeq"] = relation.GetAttrValue("OPLINE") == null ? DBNull.Value : relation.GetAttrValue("OPLINE");
                detailRow["RltOptionFlag"] = relation.GetAttrValue("RLTOPTIONFLAG") == null ? false : relation.GetAttrValue("RLTOPTIONFLAG");
                detailRow["LtPercent"] = relation.GetAttrValue("LTPERCENT") == null ? 0 : relation.GetAttrValue("LTPERCENT");
                detailRow["PRoutinginspId"] = item.Id;
                detailRow["ReportFlag"] = relation.GetAttrValue("REPORTFLAG") == null ? true : relation.GetAttrValue("REPORTFLAG");
                detailRow["BFFlag"] = relation.GetAttrValue("BFFLAG") == null ? false : relation.GetAttrValue("BFFLAG");
                detailRow["FeeFlag"] = relation.GetAttrValue("FEEFLAG") == null ? true : relation.GetAttrValue("FEEFLAG");
                detailRow["PlanSubFlag"] = relation.GetAttrValue("PLANSUBFLAG") == null ? false : relation.GetAttrValue("PLANSUBFLAG");
                detailRow["DeliveryDays"] = relation.GetAttrValue("DELIVERYDAYS") == null ? 0 : relation.GetAttrValue("DELIVERYDAYS");
                detailRow["SplitFlag"] = false;
                detailDt.Rows.Add(detailRow);
            }
                #endregion
            ds.Add(detailDt);

            return ds;
        }

        private DataTable GetRoutingDetailDataTable() {
            DataTable dt = new DataTable("RoutingDetail");
            dt.Columns.Add("PRoutingDId");//工序ID
            dt.Columns.Add("PRoutingId");//工艺路线Id
            dt.Columns.Add("OpSeq");//工序行号
            dt.Columns.Add("OperationCode");//工序代号
            dt.Columns.Add("Description");//工序描述
            dt.Columns.Add("WcCode");//工作中心代号
            dt.Columns.Add("EffBegDate");//版本生效日期
            dt.Columns.Add("EffEndDate");//版本失效日期
            dt.Columns.Add("SubFlag", typeof(bool));//是否委外工序
            dt.Columns.Add("RltOptionFlag", typeof(bool));//是/否可选(1/0)
            dt.Columns.Add("LtPercent", typeof(int));//制造提前期百分比  
            dt.Columns.Add("PRoutinginspId");//客户BOM对应的工艺路线各工序的检验ID  
            dt.Columns.Add("ReportFlag", typeof(bool));//是否报告点
            dt.Columns.Add("BFFlag", typeof(bool));//是否倒冲工序
            dt.Columns.Add("FeeFlag", typeof(bool));//是否计费点
            dt.Columns.Add("PlanSubFlag", typeof(bool));//是否计划委外工序
            dt.Columns.Add("DeliveryDays", typeof(int));//交货日期
            dt.Columns.Add("SplitFlag", typeof(bool));
            return dt;
        }

        /// <summary>
        /// 构建Part节点结构
        /// </summary>
        /// <returns></returns>
        private DataTable GetPartDataTable() {
            DataTable dt = new DataTable("Part");
            dt.Columns.Add("PRoutingId");//工艺路线Id
            dt.Columns.Add("InvCode");//母件编码
            return dt;
        }
        #endregion

        #region 构建标准工序结构
        private List<DataTable> BuildOperationDataSet(DEBusinessItem rootItem) {
            List<DataTable> ds = new List<DataTable>();
            if (rootItem == null) {
                return null;
            }

            DataTable versionDt = GetOperationDetailDataTable();
            var items = GetLinks(rootItem, "PARTTOGX");
            if (items == null || items.Count == 0) {
                return ds;
            }
            for (int i = 0; i < items.BizItems.Count; i++) {
                #region OperationDetail节点
                var item = items.BizItems[i] as DEBusinessItem;
                var relation = items.RelationList[i] as DERelation2;
                if (item == null) {
                    continue;
                }
                var versionRow = versionDt.NewRow();
                versionRow["OperationId"] = item.GetAttrValue(item.ClassName, "OID") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "OID");
                versionRow["OpCode"] = item.Id;
                versionRow["Description"] = item.GetAttrValue(item.ClassName, "GXSM") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "GXSM");
                versionRow["SubFlag"] = item.GetAttrValue(item.ClassName, "WWGX") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "WWGX");
                versionRow["DeliverDays"] = item.GetAttrValue(item.ClassName, "DELIVERDAYS") == null ? DBNull.Value : item.GetAttrValue(item.ClassName, "DELIVERDAYS");
                versionRow["IsBF"] = relation.GetAttrValue("BFFLAG") == null ? false : relation.GetAttrValue("BFFLAG");
                versionRow["IsFee"] = item.GetAttrValue(item.ClassName, "ISFEE") == null ? true : item.GetAttrValue(item.ClassName, "ISFEE");
                versionRow["IsPlanSub"] = item.GetAttrValue(item.ClassName, "ISPLANSUB") == null ? false : item.GetAttrValue(item.ClassName, "ISPLANSUB");
                versionRow["IsReport"] = item.GetAttrValue(item.ClassName, "REPORT") == null ? true : item.GetAttrValue(item.ClassName, "REPORT");
                versionDt.Rows.Add(versionRow);
            }
                #endregion
            ds.Add(versionDt);
            return ds;
        }

        /// <summary>
        /// 创建标准工序结构
        /// </summary>
        /// <returns></returns>
        private DataTable GetOperationDetailDataTable() {
            DataTable dt = new DataTable("OperationDetail");
            dt.Columns.Add("OperationId");//工序Id
            dt.Columns.Add("OpCode");//工序代号
            dt.Columns.Add("Description");//描述 
            dt.Columns.Add("RltOptionFlag", typeof(bool));//是/否可选(1/0)   
            dt.Columns.Add("SubFlag", typeof(bool));//是/否委外工序(1/0)   
            dt.Columns.Add("DeliverDays", typeof(int));//交货天数 
            dt.Columns.Add("IsBF", typeof(bool));//是否倒冲工序 
            dt.Columns.Add("IsFee", typeof(bool));//是否计费点
            dt.Columns.Add("IsPlanSub", typeof(bool));//是否计划委外
            dt.Columns.Add("IsReport", typeof(bool));
            return dt;
        }
        #endregion

        #region 初始化

        public void Init() {
            this.d_AfterReleased = new PLMBizItemDelegate(AfterItemReleased);
            this.d_AfterDeleted = new PLMDelegate2(AfterItemDeleted);
            BizItemHandlerEvent.Instance.D_AfterDeleted = (PLMDelegate2)Delegate.Combine(BizItemHandlerEvent.Instance.D_AfterDeleted,this.d_AfterDeleted);
            BizItemHandlerEvent.Instance.D_AfterReleased = (PLMBizItemDelegate)Delegate.Combine(BizItemHandlerEvent.Instance.D_AfterReleased, this.d_AfterReleased);
        }

        private void AfterItemDeleted(object sender, PLMOperationArgs e) {

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
                        //ExportExecute(item);
                        AddOrEditItem(item);
                    }
                }
            }
        }
      
        #endregion
        #region 导出导入
        private void AddOrEditItem(IBizItem item) {
            var bItem = GetDEBusinessItem(item);
            if (bItem == null) {
                return;
            }
            string oprt = item.LastRevision > 1 ? "Edit" : "Add";
            switch (bItem.ClassName.ToLower()) {
                default:
                    break;
                case "unitgroup":
                    UnitGroupDal unitGroupDal = new UnitGroupDal(bItem);
                    var doc = unitGroupDal.CreateXmlDocument(oprt);
                    ConnectEAI(doc.OuterXml);
                    return;
            }
        }

        /// <summary>
        /// 导入到ERP
        /// </summary>
        /// <param name="xml"></param>
        private void ConnectEAI(string xml) { 
            MSXML2.XMLHTTPClass xmlHttp = new MSXML2.XMLHTTPClass();
            xmlHttp.open("POST", "http://kexp/u8eai/import.asp", false, null, null);//TODO：地址需要改
            xmlHttp.send(xml);
            String responseXml = xmlHttp.responseText;
            //…… //处理返回结果 
            XmlDocument resultDoc = new XmlDocument();
            resultDoc.LoadXml(responseXml);
            var itemNode = resultDoc.SelectSingleNode("ufinterface//item");
            if (itemNode==null) {
                PLMEventLog.WriteLog("没有收到ERP回执！", EventLogEntryType.Error);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp); //COM释放
                return;
            }
            var succeed = Convert.ToInt32(itemNode.Attributes["succeed"].Value);//成功标识：0：成功；非0：失败；
            var dsc =itemNode.Attributes["dsc"].ToString();
            //var u8key =itemNode.Attributes["u8key"].ToString();
            //var proc = itemNode.Attributes["proc"].ToString();
            if (succeed==0) { 
                PLMEventLog.WriteLog("导入成功!", EventLogEntryType.Information);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp); //COM释放
                return;
            }
            PLMEventLog.WriteLog(dsc, EventLogEntryType.Error);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xmlHttp); //COM释放
        }
        #endregion
        #region 构建BOM结构

        /// <summary>
        /// 建立BOM节点
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private List<DataTable> BuildBomDataSet(DEBusinessItem item) {
            List<DataTable> ds = new List<DataTable>();
            if (item == null) {
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
            ds = GetComponnetDataSet(item, ds);
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
            if (item == null) {
                return null;
            }
            //DERelationBizItemList relationBizItemList = item.Iteration.LinkRelationSet.GetRelationBizItemList("PARTTOPART");
            //if (relationBizItemList == null) {
            //    try {
            //        relationBizItemList = PLItem.Agent.GetLinkRelationItems(item.Iteration.Oid, item.Master.ClassName, "PARTTOPART", ClientData.LogonUser.Oid, ClientData.UserGlobalOption);
            //    } catch {
            //        return null;
            //    }
            //}
            //if (relationBizItemList == null || relationBizItemList.Count == 0) {
            //    return null;
            //}
            var dt = GetBoms(item);
            ds.Add(dt);
            return ds;
        }

        /// <summary>
        /// 获取物料信息
        /// </summary>
        /// <param name="rootItem"></param>
        /// <returns></returns>
        private DataTable GetBoms(DEBusinessItem rootItem) {
            _seq = 0;
            if (rootItem == null) {
                return null;
            }
            List<DEBusinessItem> list = new List<DEBusinessItem>();
            DERelationBizItemList relationBizItemList = GetLinks(rootItem, "PARTTOGX");
            if (relationBizItemList == null || relationBizItemList.Count == 0) {
                return null;
            }
            DataTable dt = GetComponentDataTable();
            for (int i = 0; i < relationBizItemList.BizItems.Count; i++) {
                var item = relationBizItemList.BizItems[i] as DEBusinessItem;
                var relationItem = relationBizItemList.RelationList[i] as DERelation2;
                if (item == null) {
                    continue;
                }
                var opLine = relationItem.GetAttrValue("OPLINE") == null ? "" : relationItem.GetAttrValue("OPLINE").ToString();

                dt = GetBomItems(item, dt, opLine);
            }
            return dt;
        }

        /// <summary>
        /// 获取工序下的子件列表
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private DataTable GetBomItems(DEBusinessItem item, DataTable dt, string opLine) {
            if (item == null) {
                return null;
            }
            if (dt == null) {
                dt = GetComponentDataTable();
            }
            var relations = GetLinks(item, "GXTOPART");
            if (relations == null || relations.Count == 0) {
                return null;
            }
            List<DEBusinessItem> list = new List<DEBusinessItem>();
            for (int i = 0; i < relations.BizItems.Count; i++) {
                var iItem = relations.BizItems[i] as DEBusinessItem;
                var relationItem = relations.RelationList[i] as DERelation2;

                if (iItem == null) {
                    continue;
                }
                if (dt.Select("OpComponentId=" + iItem.Id).Length > 0) {
                    continue;
                }
                _seq++;
                var row = dt.NewRow();
                #region 物料属性

                row["BomId"] = item.Id;
                row["OpComponentId"] = iItem.Id;
                row["BomType"] = iItem.GetAttrValue(iItem.ClassName, "BOMTYPE") == null ? DBNull.Value : iItem.GetAttrValue(iItem.ClassName, "BOMTYPE");
                row["SortSeq"] = 10 * _seq;
                row["OpSeq"] = opLine;
                row["InvCode"] = iItem.GetAttrValue(iItem.ClassName, "OID") == null ? DBNull.Value : iItem.GetAttrValue(iItem.ClassName, "OID");
                row["Version"] = iItem.LastRevision;
                row["Status"] = iItem.GetAttrValue(iItem.ClassName, "INSSTATE") == null ? DBNull.Value : iItem.GetAttrValue(iItem.ClassName, "INSSTATE");
                row["ModifyTime"] = iItem.GetAttrValue(iItem.ClassName, "CHECKINTIME") == null ? DBNull.Value : iItem.GetAttrValue(iItem.ClassName, "CHECKINTIME");
                row["EffEndDate"] = iItem.GetAttrValue(iItem.ClassName, "VERSIONENDDATE") == null ? DBNull.Value : iItem.GetAttrValue(iItem.ClassName, "VERSIONENDDATE");
                row["EffBegDate"] = iItem.GetAttrValue(iItem.ClassName, "VERSIONEFFDATE") == null ? DBNull.Value : iItem.GetAttrValue(iItem.ClassName, "VERSIONEFFDATE");
                #endregion
                #region 关系属性

                row["CompScrap"] = relationItem.GetAttrValue("PARENTSCRAP") == null ? DBNull.Value : relationItem.GetAttrValue("PARENTSCRAP");
                row["FVFlag"] = relationItem.GetAttrValue("FVFLAG") == null ? DBNull.Value : relationItem.GetAttrValue("FVFLAG");
                row["BaseQtyN"] = relationItem.GetAttrValue("BASEQTYN") == null ? DBNull.Value : relationItem.GetAttrValue("BASEQTYN");
                row["BaseQtyD"] = relationItem.GetAttrValue("BASEQTYD") == null ? DBNull.Value : relationItem.GetAttrValue("BASEQTYD");
                row["ByproductFlag"] = relationItem.GetAttrValue("BYPRODUCTFLAG") == null ? DBNull.Value : relationItem.GetAttrValue("BYPRODUCTFLAG");
                row["ProductType"] = relationItem.GetAttrValue("PRODUCTTYPE") == null ? DBNull.Value : relationItem.GetAttrValue("PRODUCTTYPE");
                #endregion

                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>
        /// 获取关联对象列表
        /// </summary>
        /// <param name="relation"></param>
        /// <returns></returns>
        private DERelationBizItemList GetLinks(DEBusinessItem item, string relation) {
            DERelationBizItemList relationBizItemList = item.Iteration.LinkRelationSet.GetRelationBizItemList(relation);
            if (relationBizItemList == null) {
                try {
                    relationBizItemList = PLItem.Agent.GetLinkRelationItems(item.Iteration.Oid, item.Master.ClassName, relation, ClientData.LogonUser.Oid, ClientData.UserGlobalOption);
                } catch {
                    return null;
                }
            }
            return relationBizItemList;
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
            dt.Columns.Add("CompScrap", typeof(decimal));//子件损耗率
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
            dt.Columns.Add("FVFlag", typeof(int));//固定/变动批量  
            dt.Columns.Add("BaseQtyN", typeof(decimal));//基础用量-分子
            dt.Columns.Add("BaseQtyD", typeof(decimal));//基础用量-分母
            dt.Columns.Add("ByproductFlag", typeof(bool));//是否产出品
            dt.Columns.Add("ProductType", typeof(int));//产出类型
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
            dt.Columns.Add("PRoutingId");//工艺路线Id
            dt.Columns.Add("RountingType", typeof(int));//工艺路线版本，默认默认“1-主工艺路线”
            dt.Columns.Add("Version", typeof(int));//版本号 
            dt.Columns.Add("VersionDesc");//版本说明 
            dt.Columns.Add("VersionEffDate", typeof(DateTime));//版本生效日 
            dt.Columns.Add("VersionEndDate", typeof(DateTime));//版本失效日 
            dt.Columns.Add("IdentCode");//替代标识 
            dt.Columns.Add("IdentDesc");//替代说明 
            dt.Columns.Add("ApplyCode");//变更单号
            dt.Columns.Add("ApplySeq");//变更单行号
            dt.Columns.Add("Status");//Status 状态(1:新建/3:审核/4:停用)
            dt.Columns.Add("ModifyTime", typeof(DateTime));//修改日期 
            dt.Columns.Add("RunCardFlag", typeof(bool));
            dt.Columns.Add("PfBatchQty", typeof(decimal));
            return dt;
        }
        #endregion
    }
}
