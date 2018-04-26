using BOMExportCommon;
using Oracle.DataAccess.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Thyt.TiPLM.BRL.Product;
using Thyt.TiPLM.Common;
using Thyt.TiPLM.DAL.Admin.NewResponsibility;
using Thyt.TiPLM.DAL.Common;
using Thyt.TiPLM.DAL.Product;
using Thyt.TiPLM.DEL.Admin.DataModel;
using Thyt.TiPLM.DEL.Admin.NewResponsibility;
using Thyt.TiPLM.DEL.Product;
using Thyt.TiPLM.PLL.Admin.DataModel;

namespace BOMExportServer {
    /// <summary>
    /// Class1 的摘要说明。
    /// </summary>
    public class DAERPExport : DABase {
        #region 自定义变量

        private static OracleCommand Errcmd;
        private readonly DEDataConvert deItem;
        private int CurSeq;
        private bool IsFindSeq;
        private string RootID = "";
        private string colItemTime = "";
        private DBDataType colItemTimeType;
        private string configName; //配置名称
        private DEExportEvent expEvent;

        /// <summary>
        /// 导入时间
        /// </summary>
        private DateTime now;

        private Guid sOid;
        private StringBuilder strEr;
        private StringBuilder strWarn;
        public StringBuilder strsuc;
        private DEUser user;

        #endregion

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="dbParam"></param>
        public DAERPExport(DBParameter dbParam)
            : base(dbParam) {
            deItem = new DEDataConvert(dbParam);
        }

        #region 导出序号获取与设置

        /// <summary>
        /// 获取当前的导出序列号
        /// </summary>
        /// <returns>序列号</returns>
        private int GetCurrentSequence() {
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = (OracleConnection)dbParam.Connection;
            cmd.CommandText = "SELECT PLM_PARAMVALUE FROM PLM_SYS_PARAMETER WHERE PLM_PARAMNAME = 'ERP_EXPORT_SEQUENCE'";
            OracleDataReader dr = null;
            try {
                dr = cmd.ExecuteReader();
                if (dr.Read()) {
                    return Convert.ToInt32(dr.GetString(0));
                }
            } finally {
                if (dr != null)
                    dr.Close();
            }
            return 0;
        }

        /// <summary>
        /// 保存当前导出序列号
        /// </summary>
        /// <param name="seq">导出序列号</param>
        private void SetCurrentSequence(int seq) {
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = (OracleConnection)dbParam.Connection;
            try {
                cmd.CommandText =
                    "Update PLM_SYS_PARAMETER Set PLM_PARAMVALUE =:ExportSeq WHERE PLM_PARAMNAME='ERP_EXPORT_SEQUENCE'";
                cmd.Parameters.Add(":ExportSeq", OracleDbType.Int32).Value = seq;
                cmd.ExecuteNonQuery();
            } catch {
                try {
                    cmd.CommandText = "DELETE FROM PLM_PARAMVALUE WHERE PLM_PARAMNAME = 'ERP_EXPORT_SEQUENCE'";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "INSERT INTO PLM_PARAMVALUE VALUES('+ERP_EXPORT_SEQUENCE+',:ExportSeq)";
                    cmd.ExecuteNonQuery();
                } catch {
                }
            }
        }

        #endregion

        /// <summary>
        /// 获取ERP数据库连接
        /// </summary>
        /// <returns></returns>
        private IDbConnection GetERPConnection() {
            return deItem.DS.GetConnection();
        }

        #endregion

        /// <summary>
        /// 导出产品结构到ERP系统中
        /// </summary>
        /// <param name="expEvent"></param>
        /// <param name="strErrinfo">出错信息</param>
        /// <param name="deErp"></param>
        /// <param name="InObjs"></param>
        /// <returns>如果成功，返回true；否则返回false</returns>
        public bool ExportPSToERP(DEExportEvent exp, out StringBuilder strErrinfo,
                                  out StringBuilder strWarnInfo, DEErpExport deErp,
                                  out DataSet dsResult, object InObjs, out object outObjs) {
            string ERPConfig = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + deErp.configName;

            if (!File.Exists(ERPConfig))
                throw new Exception("配置文件 " + deErp.configName + " 没有部署在服务器端");
            deItem.GetConfiguration(ERPConfig, deErp);

            strEr = new StringBuilder();
            strWarn = new StringBuilder();
            strsuc = new StringBuilder();
            dsResult = null;
            outObjs = null;
            DataSet ds = new DataSet();

            strErrinfo = strEr;
            strWarnInfo = strWarn;
            // 权限校验
            DAUser daUser = new DAUser(dbParam);
            expEvent = exp;
            user = daUser.GetByOid(expEvent.Exportor);
            sOid = expEvent.Oid;


            Errcmd = new OracleCommand();
            Errcmd.Connection = (OracleConnection)dbParam.Connection;

            DABomExport daExport = new DABomExport(dbParam);
            //expEvent = daExport.GetExportEvent(expEvent.Oid);
            //if (expEvent == null)
            //    throw new PLMException("GetExportEvent无法获取到指定的导出事件信息，导出失败！");
            /*
                        if( deItem.ItemTable != null && deItem.ItemTable.IsExport && deItem.ItemTable.ID_Item == null )
                            throw new Exception("在配置文件中没有设置物料导出所必须的“M.PLM_M_ID”字段,数据无法更新，导出中止");
                        if( deItem.BomTable != null && deItem.BomTable.IsExport && deItem.BomTable.ID_Item == null )
                            throw new Exception("在配置文件中没有设置Bom导出所必须的“PID”字段,数据无法更新，导出中止");
                        if( deItem.BomHeadTable != null && deItem.BomHeadTable.IsExport && deItem.BomHeadTable.ID_Item == null )
                            throw new Exception("在配置文件中没有设置BomHead导出所必须的“PID”字段,数据无法更新，导出中止");
             * */
            // 获取根对象
            QRItem qrItem = new QRItem(dbParam);
            RevOccurrenceType revOccurrenceType;
            DEBusinessItem item = qrItem.GetBizItem(
                expEvent.ItemOid,
                expEvent.ItemRevNum,
                0,
                expEvent.Exportor,
                expEvent.PSOption,
                out revOccurrenceType,
                BizItemMode.BizItem) as DEBusinessItem;
            if (item == null)
                throw new PLMException(ExceptionManager.E_PDT_CANNOT_GET_VALIDDATA);

            configName = "";
            if (expEvent.PSOption.ConfigStatus != Guid.Empty) {
                DAVStatus daVar = new DAVStatus(dbParam);
                DEVStatus configStatus = daVar.GetStatusByOidAndIter(expEvent.PSOption.ConfigStatus,
                                                                     expEvent.PSOption.ConfigStatusRev, Guid.Empty, 0);
                if (configStatus != null) configName = configStatus.Name;
            }
            /*
                        if( deItem.ItemTable != null && deItem.ItemTable.IsExport && !CheakItemCanExport(item.Master.ClassName) )
                            throw new PLMException("不能导出数据类型为" + item.Master.ClassName +
                                                   "的数据,只允许导出配置文件的<ItemColumnMapping :SourceColumn 中定义的类型或其子类型>");
             * */
            RootID = item.Master.Id;
            now = DateTime.Now;
            CurSeq = GetCurrentSequence() + 1;

            Type type = null;
            IDbConnection conn = null;
            IDbTransaction trans = null;
            bool IsRight = true;
            try {
                #region 生成完整的产品结构数据

                GetPS(item);
                ClearUpItem();

                #endregion

                //数据预处理与校验
                //if( !string.IsNullOrEmpty(deItem.DP_FUN_NAME) )
                //    DPERPData(deItem.DP_FUN_NAME);
                if (!string.IsNullOrEmpty(deItem.DP_BEFORE_NAME) && !deItem.IsExeDPAfterInterface)
                    DpBeforeExport(deItem.DP_BEFORE_NAME, item, expEvent.PSOption, expEvent.ExpOption.MaxLevel);
                if (!String.IsNullOrEmpty(deErp.ExtendSVRDllName)) {
                    IsRight = ExtendBeforeExport(deErp.ExtendSVRDllName, InObjs, out type);
                    if (!IsRight)
                        return false;
                }
                if (!string.IsNullOrEmpty(deItem.DP_BEFORE_NAME) && deItem.IsExeDPAfterInterface)
                    DpBeforeExport(deItem.DP_BEFORE_NAME, item, expEvent.PSOption, expEvent.ExpOption.MaxLevel);
                //发现校验错误
                GetErrInfo();
                if (strEr.Length > 0) {
                    strEr.Insert(0, "数据经验和中间表处理过程中发现错误，数据导出停止\r\n");
                    return false;
                }
                // 获取要导出的ITEM信息
                Hashtable hsItemTables = new Hashtable();
                if (deItem.IsExportItem)
                    for (int i = 0; i < deItem.lstItemTables.Count; i++) {
                        ItemColumnMapping ItemMapping = deItem.lstItemTables[i] as ItemColumnMapping;
                        DataTable tb = BuildItemTable(deErp.IsExpOldItemAndBom, ItemMapping);
                        if (tb != null && tb.Rows.Count > 0)
                            hsItemTables.Add(deItem.lstItemTables[i], tb);
                    }
                //  DataTable itemTable = BuildItemTable(deErp.IsExpOldItemAndBom);

                // 获取要导出的BOM信息
                Hashtable hsBomTables = new Hashtable();
                //   DataTable bomTable = BuildBomTable(deItem.BomTable, false, deErp.IsExpOldItemAndBom);
                if (deItem.IsExportBom)
                    for (int i = 0; i < deItem.lstBomTables.Count; i++) {
                        ItemColumnMapping BomMapping = deItem.lstBomTables[i] as ItemColumnMapping;
                        DataTable tb = BuildBomTable(BomMapping, deErp.IsExpOldItemAndBom);
                        if (tb != null && tb.Rows.Count > 0)
                            hsBomTables.Add(deItem.lstBomTables[i], tb);
                    }
                // 获取要导出的工艺路线信息
                Hashtable hsRouteTables = new Hashtable();
                for (int i = 0; i < deItem.lstRouteTables.Count; i++) {
                    RelItemColumnMapping routeMapping = deItem.lstRouteTables[i] as RelItemColumnMapping;
                    if (routeMapping == null || !routeMapping.IsExport) continue;
                    DataTable tb = BuildRouteTable(routeMapping);
                    if (tb != null && tb.Rows.Count > 0)
                        hsRouteTables.Add(deItem.lstRouteTables[i], tb);
                }
                GetErrInfo();
                if (strEr.Length > 0) {
                    // GetErrInfo(eventOid);
                    strEr.Insert(0, "获取数据出错，数据导出停止\r\n");
                    return false;
                }
                try {
                    if (deErp.op == ExpOption.eDataBase) {
                        conn = GetERPConnection();
                        conn.Open();
                        trans = conn.BeginTransaction();
                    }
                } catch (Exception ex) {
                    strEr.Insert(0, "建立到中间库的数据连接出错:" + ex.Message + "，数据导出停止\r\n连接字符串：" + deItem.DS.ConnStr);
                    return false;
                }
                if (type != null) {
                    IsRight = ExtendMidExport(type, hsItemTables, hsBomTables, hsRouteTables, conn, trans);
                }
                GetErrInfo();
                if (!IsRight)
                    return false;
                if (strEr.Length > 0) {
                    strEr.Insert(0, "获取数据出错，数据导出停止\r\n");
                    return false;
                }

                #region 导出到ERP临时表

                IDictionaryEnumerator IE;
                // 更新ITEM到ERP临时表
                if (deItem.lstItemTables != null) {
                    for (int i = 0; i < deItem.lstItemTables.Count; i++) {
                        ItemColumnMapping mapping = deItem.lstItemTables[i] as ItemColumnMapping;
                        if (mapping == null) continue;
                        DataTable tb = hsItemTables.Contains(mapping) ? (DataTable)hsItemTables[mapping] : null;
                        if (tb == null || tb.Rows.Count == 0) continue;
                        if (mapping.IsExport) {
                            if (deErp.op == ExpOption.eDataBase)
                                UpdateItemTable(mapping, tb, conn, trans);
                            ds.Tables.Add(tb.Copy());
                        }
                    }
                }
                // 更新BOM到ERP临时表
                if (deItem.lstBomTables != null) {
                    for (int i = 0; i < deItem.lstBomTables.Count; i++) {
                        ItemColumnMapping mapping = deItem.lstBomTables[i] as ItemColumnMapping;
                        if (mapping == null) continue;
                        DataTable tb = hsBomTables.Contains(mapping) ? (DataTable)hsBomTables[mapping] : null;
                        if (tb == null || tb.Rows.Count == 0) continue;
                        if (mapping.IsExport) {
                            if (deErp.op == ExpOption.eDataBase)
                                UpdateTable(mapping, tb, conn, trans);
                            ds.Tables.Add(tb.Copy());
                        }
                    }
                }

                if (hsRouteTables.Count > 0) {
                    IE = hsRouteTables.GetEnumerator();
                    IE.Reset();
                    while (IE.MoveNext()) {
                        DataTable tb = IE.Value as DataTable;
                        ItemColumnMapping mapping = IE.Key as ItemColumnMapping;
                        if (mapping.IsExport) {
                            if (deErp.op == ExpOption.eDataBase)
                                UpdateTable(IE.Key, IE.Value, conn, trans);
                            ds.Tables.Add(tb.Copy());

                        }
                    }
                }

                if (!deItem.IsExeDPAllEnd && deItem.DP_END_NAME != "") {
                    AfterExport();
                }

                if (type != null) {
                    IsRight = ExtendAfterExport(type, out dsResult, out outObjs);
                }
                if (IsRight && deErp.op == ExpOption.eDataBase) {
                    if (conn != null && !conn.State.Equals(ConnectionState.Closed))
                        trans.Commit();
                }

                #endregion

                if (dsResult == null && deErp.op != ExpOption.eDataBase)
                    dsResult = ds;

                if (IsFindSeq) SetCurrentSequence(CurSeq);

                if (deItem.IsExeDPAllEnd && !String.IsNullOrEmpty(deItem.DP_END_NAME)) {
                    AfterExport();
                    GetErrInfo();
                }
                if (strEr.Length == 0 && deErp.op == ExpOption.eDataBase) {
                    try {
                        RecordExportHistory(deErp.IsExpOldItemAndBom);
                    } catch (Exception et) {
                        strWarn.Append("\r\n数据导出成功，但记录导出历史失败:" + et.Message + "\r\n");
                    }
                }
            } catch (Exception ex) {
                if (deErp.op == ExpOption.eDataBase && (trans != null && !conn.State.Equals(ConnectionState.Closed)))
                    trans.Rollback();
                RecordErr(sOid, ex.ToString());
            } finally {
                if (deErp.op == ExpOption.eDataBase && (conn != null && !conn.State.Equals(ConnectionState.Closed)))
                    conn.Close();
                daExport.ClearPS(sOid);
                GetErrInfo();
            }
            return strEr.Length == 0;
        }

        private void RecordExportHistory(bool isExpOldItemAndBom) {
            OracleCommand cmd = new OracleCommand();
            try {
                cmd.Connection = (OracleConnection)dbParam.Connection;
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(":expuser", OracleDbType.Varchar2).Value = user.LogId;
                cmd.Parameters.Add(":expdate", OracleDbType.Date).Value = now;
                cmd.Parameters.Add(":PLM_OPER", OracleDbType.Varchar2);
                cmd.Parameters.Add(":soid", OracleDbType.Raw).Value = sOid.ToByteArray();
                cmd.Parameters.Add(":db_name", OracleDbType.Varchar2).Value = deItem.DS.GetDBName;

                StringBuilder str = new StringBuilder();
                string strEx;
                if (deItem.lstItemTables.Count > 0 && deItem.IsExportItem) {
                    str.Append(
                        "insert into plm_exp_item_history(plm_moid,plm_roid,plm_iroid,plm_class,plm_expuser,plm_expdate,PLM_DATABASE,PLM_OPER,PLM_TOOL)");
                    str.Append(
                        " select distinct ps.plm_cmasteroid,ps.plm_crevoid,ps.plm_citeroid,ps.plm_childclass,:expuser,:expDate,'" +
                        deItem.DS.GetDBName + "',:PLM_OPER,'BOM集成' ");
                    str.Append(" from plm_psm_ps ps where ps.plm_sessionoid =:ssoid");

                    if (isExpOldItemAndBom) {
                        strEx =
                            " and  exists ( select 1 from plm_exp_item_history h	where h.plm_iroid = ps.plm_citeroid   and h.plm_database = :db_name )";
                        cmd.Parameters[":PLM_OPER"].Value = "更新";
                        cmd.CommandText = str + strEx;
                        cmd.ExecuteNonQuery();
                    }
                    strEx =
                        " and not  exists ( select 1 from plm_exp_item_history h	where h.plm_iroid = ps.plm_citeroid   and h.plm_database = :db_name )";
                    cmd.Parameters[":PLM_OPER"].Value = "新增";
                    cmd.CommandText = str + strEx;
                    cmd.ExecuteNonQuery();
                }
                if (deItem.lstBomTables.Count > 0 && deItem.IsExportBom) {
                    str.Remove(0, str.Length);
                    str.Append(
                        "insert into plm_exp_bom_history(plm_p_moid,plm_p_roid,plm_p_iroid,plm_p_class,plm_c_moid,plm_c_roid,plm_c_iroid,plm_c_class,plm_rel_oid,plm_number,plm_expuser,plm_expdate, plm_database,plm_oper)");
                    str.Append(
                        " select distinct ps0.plm_cmasteroid,ps0.plm_crevoid,ps0.plm_citeroid,ps0.plm_childclass,ps1.plm_cmasteroid,ps1.plm_crevoid,ps1.plm_citeroid,ps1.plm_childclass,ps1.plm_reloid,ps1.plm_number,:expuser,:expDate,'" +
                        deItem.DS.GetDBName + "',:PLM_OPER ");
                    str.Append(" from plm_psm_ps ps0 ,plm_psm_ps ps1 ");
                    str.Append(
                        " where ps1.plm_sessionoid = :ssoid  and ps0.plm_sessionoid = ps1.plm_sessionoid  and ps0.plm_citeroid = ps1.plm_piteroid");

                    strEx =
                        " and Exists (select 1 from plm_exp_bom_history h2  where h2.plm_p_iroid  = ps0.plm_citeroid  and h2.plm_c_iroid = ps1.plm_citeroid";
                    strEx +=
                        " and h2.plm_p_moid = ps0.plm_cmasteroid  and h2.plm_c_moid = ps1.plm_cmasteroid   and h2.plm_database = :db_Name)";
                    cmd.Parameters[":PLM_OPER"].Value = "更新";
                    cmd.CommandText = str + strEx;
                    cmd.ExecuteNonQuery();

                    strEx =
                        " and not Exists (select 1 from plm_exp_bom_history h2  where h2.plm_p_iroid  = ps0.plm_citeroid  and h2.plm_c_iroid = ps1.plm_citeroid";
                    strEx +=
                        " and h2.plm_p_moid = ps0.plm_cmasteroid  and h2.plm_c_moid = ps1.plm_cmasteroid  and h2.plm_database = :db_Name)";
                    cmd.Parameters[":PLM_OPER"].Value = "新增";
                    cmd.CommandText = str + strEx;
                    cmd.ExecuteNonQuery();
                }
            } finally {
                try {
                    cmd.Dispose();
                } catch {
                }
            }
        }

        #region 获取导出数据

        /// <summary>
        /// 获取产品结构
        /// </summary>
        /// <param name="expEvent"></param>
        /// <param name="item"></param>
        /// <param name="opt"></param>
        private void GetPS(DEBusinessItem item) {
            try {
                DAItem daItem = new DAItem(dbParam);
                daItem.GetRelationStructure(expEvent.Oid, item.MasterOid, item.RevOid, item.IterOid, item.ClassName,
                                            "PARTTOPART", expEvent.PSOption, expEvent.ExpOption.MaxLevel, user.Oid);
            } catch (Exception ex) {
                throw new Exception("获取产品结构数据失败", ex);
            }
        }


        /// <summary>
        /// 清理PSM_PS 中的不需要导出的数据
        /// </summary>
        /// <param name="soid"></param>
        private void ClearUpItem() {
            OracleCommand cmd = new OracleCommand();
            if (!deItem.IsExportItem || deItem.lstItemTables.Count == 0) return;
            ArrayList lstCls = new ArrayList();
            for (int i = 0; i < deItem.lstItemTables.Count; i++) {
                ItemColumnMapping mapping = deItem.lstItemTables[i] as ItemColumnMapping;
                if (mapping == null) continue;
                for (int j = 0; j < mapping.lstItemClass.Count; j++) {
                    string c = mapping.lstItemClass[j].ToString();
                    if (!lstCls.Contains(c)) lstCls.Add(c);
                }
            }
            cmd.Connection = (OracleConnection)dbParam.Connection;
            StringBuilder str = new StringBuilder();
            str.Append("Delete From plm_psm_ps ps Where ps.plm_sessionoid = :soid And ps.plm_citeroid Not In(");
            string cls;
            for (int i = 0; i < lstCls.Count; i++) {
                cls = lstCls[i].ToString();
                if (i != 0)
                    str.Append(" Union All ");
                str.Append(" select ps" + i + ".plm_citeroid  From plm_psm_ps ps" + i);
                str.Append(" , plm_cusv_" + cls + " t" + i);
                str.Append(" where ps" + i + ".plm_citeroid  = t" + i + ".PLM_OID ");
            }
            str.Append(")");
            cmd.Parameters.Add(":soid", OracleDbType.Raw);
            cmd.Parameters[":soid"].Value = expEvent.Oid.ToByteArray();
            cmd.CommandText = str.ToString();
            try {
                cmd.ExecuteNonQuery();
            } catch (Exception ex) {
                throw new PLMException(ex.Message + " 清空不能导出数据： " + str);
            }
        }

        private void SetOrderSql(string colName, OrderType tp, StringBuilder strOrder) {
            if (tp == OrderType.None) return;
            if (strOrder.Length == 0)
                strOrder.Append(" Order by ");
            else
                strOrder.Append(",");
            if (tp == OrderType.Asc) {
                strOrder.Append(colName);
            } else {
                strOrder.Append(colName + " DESC");
            }
        }


        /// <summary>
        /// 从系统中获取需要导出的ITEM信息，并构造成ITEM表准备导出到ERP中
        /// </summary>
        /// <param name="sessionOid">会话标识</param>
        /// <returns>表</returns>
        private DataTable BuildItemTable(bool isExpHistoryItemAndBom, ItemColumnMapping ItemMapping) {
            if (ItemMapping == null || String.IsNullOrEmpty(ItemMapping.TableName) || !ItemMapping.IsExport)
                return null;

            string cls;
            StringBuilder sql = new StringBuilder();
            StringBuilder sqlOrder = new StringBuilder();
            SourceItem srcItem;
            StringBuilder er1 = new StringBuilder();
            StringBuilder er2 = new StringBuilder();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = (OracleConnection)dbParam.Connection;
            cmd.CommandType = CommandType.Text;
            OracleDataReader dr = null;

            DataTable tbItem = new DataTable(ItemMapping.TableName);
            try {
                foreach (ColumnMappingItem item in ItemMapping.MappingItemList) {
                    DataColumn col = new DataColumn(item.DestCol, deItem.ConvertSystemType(item.DestDataType));
                    tbItem.Columns.Add(col);
                }
            } catch (Exception et) {
                throw new Exception("构建物料表结构出错，请检查物料表的配置", et);
            }

            //按照lstItemClass 中存在的数据类型分别获取数据
            //这些定义的类型之间不能存在父子关系


            //存储当前获取数据目标表与原表的映射
            Hashtable hsIndex = new Hashtable();
            for (int i = 0; i < ItemMapping.lstItemClass.Count; i++) {
                cls = ItemMapping.lstItemClass[i].ToString();
                sql.Remove(0, sql.Length);
                sql.Append("SELECT DISTINCT ");
                hsIndex.Clear();
                int index = 0;
                sqlOrder.Remove(0, sqlOrder.Length);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(":SessionOid1", OracleDbType.Raw).Value = sOid.ToByteArray();
                foreach (ColumnMappingItem item in ItemMapping.MappingItemList) {
                    srcItem = item.hsSrcItems[cls] as SourceItem;
                    if (srcItem == null) {
                        srcItem = item.hsSrcItems[""] as SourceItem;
                        if (srcItem != null)
                            item.curSrcItem = srcItem;
                        else
                            item.curSrcItem = null;
                        continue;
                    }
                    item.curSrcItem = srcItem;
                    sql.Append(srcItem.SrcCol);
                    if (item.ordTy != OrderType.None) {
                        SetOrderSql(srcItem.SrcCol, item.ordTy, sqlOrder);
                    }
                    sql.Append(",");
                    hsIndex.Add(item.DestCol, index++);
                }
                sql.Remove(sql.Length - 1, 1);
                sql.Append(" FROM PLM_PSM_ITEMMASTER_REVISION M,PLM_CUSV_" + cls + " I,PLM_PSM_PS PS ");
                sql.Append(" WHERE M.PLM_R_OID=PS.PLM_CREVOID AND I.PLM_OID=PS.PLM_CITEROID ");
                sql.Append(" AND PS.PLM_SESSIONOID = :SessionOid1 ");
                if (!isExpHistoryItemAndBom) {
                    sql.Append(
                        "	and not exists ( select 1 from plm_exp_item_history h	where h.plm_iroid = ps.plm_citeroid ");
                    sql.Append(" and h.plm_database = :db_name )");
                    cmd.Parameters.Add(":db_name", OracleDbType.Varchar2).Value = deItem.DS.GetDBName;
                }
                if (sqlOrder.Length > 0)
                    sql.Append(sqlOrder.ToString());
                cmd.CommandText = sql.ToString();
                try {
                    try {
                        dr = cmd.ExecuteReader();
                    } catch (Exception ex1) {
                        RecordErr(sOid, ex1.Message + "出错sql语句：" + sql);
                        return null;
                    }
                    while (dr.Read()) {
                        er1.Remove(0, er1.Length);
                        er2.Remove(0, er2.Length);
                        DataRow row = tbItem.NewRow();
                        GetFillData(row, hsIndex, ItemMapping, dr, er1, er2);
                        string id = row[ItemMapping.ID_Item.DestCol] == DBNull.Value
                                        ? "物料代号没有填写"
                                        : row[ItemMapping.ID_Item.DestCol].ToString();
                        string clsName2 = ModelContext.MetaModel.GetClassLabel(cls);
                        if (er1.Length == 0) {
                            tbItem.Rows.Add(row);
                        } else {
                            RecordErr(sOid, null, id, clsName2, null, er1.ToString(), null);
                        }
                        if (er2.Length > 0) {
                            RecordErr(sOid, null, id, clsName2, null, null, er2.ToString());
                        }
                    }
                } catch (Exception et) {
                    if (et.InnerException == null)
                        strEr.Append(et.Message);
                    else
                        strEr.Append(et.Message + ":" + et.InnerException);
                } finally {
                    if (dr != null)
                        dr.Close();
                }
            }
            CheckDataUnique(tbItem, ItemMapping);

            return tbItem;
        }


        /// <summary>
        /// 从系统中获取需要导出的BOM信息，并构造成BOM表准备导出到ERP中
        /// </summary>
        /// <param name="sessionOid">会话标识</param>
        /// <returns>表</returns>
        private DataTable BuildBomTable(ItemColumnMapping mapping, bool isExpHistoryItemAndBom) {
            if (mapping == null || String.IsNullOrEmpty(mapping.TableName) || !mapping.IsExport)
                return null;
            StringBuilder srcCols = new StringBuilder();
            StringBuilder strOrder = new StringBuilder();
            DataTable bomTable = new DataTable(mapping.TableName);
            string colName;
            int idx;
            int index = 0;
            Hashtable hsIndex = new Hashtable();
            StringBuilder Ercol = new StringBuilder(); //记录每行出错数据
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!string.IsNullOrEmpty(item.DestCol))
                    bomTable.Columns.Add(new DataColumn(item.DestCol, deItem.ConvertSystemType(item.DestDataType)));

                switch (item.curSrcItem.Value) {
                    case "NUMBER":
                        if (srcCols.Length > 0)
                            srcCols.Append(",");
                        srcCols.Append("PS.PLM_NUMBER");
                        hsIndex.Add(item.DestCol, index++);
                        SetOrderSql("PS.PLM_NUMBER", item.ordTy, strOrder);
                        break;
                    case "LEVEL":
                        if (srcCols.Length > 0)
                            srcCols.Append(",");
                        srcCols.Append("PS.PLM_LEVEL");
                        hsIndex.Add(item.DestCol, index++);
                        SetOrderSql("PS.PLM_LEVEL", item.ordTy, strOrder);
                        break;
                    case "PID":
                        if (srcCols.Length > 0)
                            srcCols.Append(",");
                        srcCols.Append("M1.PLM_M_ID");
                        hsIndex.Add(item.DestCol, index++);
                        SetOrderSql("M1.PLM_M_ID", item.ordTy, strOrder);
                        break;
                    case "CID":
                        if (srcCols.Length > 0)
                            srcCols.Append(",");
                        srcCols.Append("M2.PLM_M_ID");
                        hsIndex.Add(item.DestCol, index++);
                        SetOrderSql("M2.PLM_M_ID", item.ordTy, strOrder);
                        break;
                    case "CONFIGSTATUS":
                    case "ROOTID":
                    case "EXPSEQUENCE":
                    case "USER.ID":
                    case "USER.NAME":
                    case "NOW":
                    case "SSOID":
                        break;
                    default:
                        //关系属性 
                        if (item.curSrcItem.Value.StartsWith("R.") || item.curSrcItem.Value.StartsWith("PS."))
                        {
                            //item.Value.Substring(2);
                            if (srcCols.Length > 0)
                                srcCols.Append(",");
                            srcCols.Append(item.curSrcItem.Value);
                            SetOrderSql(item.curSrcItem.Value, item.ordTy, strOrder);
                        } else if (item.curSrcItem.Value.StartsWith("P.")) {
                            idx = item.curSrcItem.Value.IndexOf(".M.");
                            if (idx != -1) {
                                colName = item.curSrcItem.Value.Substring(idx + 3);
                                if (srcCols.Length > 0)
                                    srcCols.Append(",");
                                srcCols.Append("M1." + colName);
                                SetOrderSql("M1." + colName, item.ordTy, strOrder);
                            } else {
                                idx = item.curSrcItem.Value.IndexOf(".I.");
                                if (idx != -1) {
                                    colName = item.curSrcItem.Value.Substring(idx + 3);
                                    if (srcCols.Length > 0)
                                        srcCols.Append(",");
                                    srcCols.Append("I1." + colName);
                                    SetOrderSql("I1." + colName, item.ordTy, strOrder);
                                }
                            }
                        } else if (item.curSrcItem.Value.StartsWith("C.")) {
                            idx = item.curSrcItem.Value.IndexOf(".M.");
                            if (idx != -1) {
                                colName = item.curSrcItem.Value.Substring(idx + 3);
                                if (srcCols.Length > 0)
                                    srcCols.Append(",");
                                srcCols.Append("M2." + colName);
                                SetOrderSql("M2." + colName, item.ordTy, strOrder);
                            } else {
                                idx = item.curSrcItem.Value.IndexOf(".I.");
                                if (idx != -1) {
                                    colName = item.curSrcItem.Value.Substring(idx + 3);
                                    if (srcCols.Length > 0)
                                        srcCols.Append(",");
                                    srcCols.Append("I2." + colName);
                                    SetOrderSql("I2." + colName, item.ordTy, strOrder);
                                }
                            }
                        } else
                            break;
                        hsIndex.Add(item.DestCol, index++);
                        break;
                }
            }
            // 执行存储过程进行本组数量的计算
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = (OracleConnection)dbParam.Connection;
            srcCols.Insert(0, "SELECT DISTINCT ");
            srcCols.Append(" FROM PLM_PSM_ITEMMASTER_REVISION M1,PLM_PSM_ITEMMASTER_REVISION M2,");
            srcCols.Append("  PLM_CUSV_PART I1,PLM_CUSV_PART I2 ,");
            srcCols.Append("  PLM_PSM_PS PS Left Join PLM_CUS_R_PARTTOPART R on PS.PLM_RELOID = R.PLM_OID ");
            srcCols.Append(" WHERE PS.PLM_SESSIONOID=:SessionOid  ");
            srcCols.Append("   And ps.plm_piteroid = I1.PLM_OID And I1.PLM_REVISIONOID = M1.PLM_R_OID  ");
            srcCols.Append(" And ps.plm_citeroid = I2.PLM_OID And ps.plm_crevoid = M2.PLM_R_OID ");
            //if( !isExpHistoryItemAndBom )
            //{
            //    srcCols.Append("	and not Exists (select 1 from plm_exp_bom_history h2 ");
            //    srcCols.Append(" where h2.plm_p_iroid  = ps.plm_piteroid and h2.plm_c_iroid = ps.plm_citeroid )");
            //}
            if (strOrder.Length > 0)
                srcCols.Append(strOrder.ToString());
            cmd.CommandText = srcCols.ToString();
            cmd.Parameters.Add(":SessionOid", OracleDbType.Raw).Value = sOid.ToByteArray();
            OracleDataReader dr = null;
            DEMetaAttribute deAttr;
            FixedAttribute fixAttr;
            try {
                try {
                    dr = cmd.ExecuteReader();
                } catch (Exception ex1) {
                    throw new Exception(ex1.Message + ":\r\n\t出错sql语句：" + srcCols);
                }
                while (dr.Read()) {
                    string pId = "", cId = "";

                    Ercol.Remove(0, Ercol.Length);
                    DataRow row = bomTable.NewRow();
                    foreach (ColumnMappingItem item in mapping.MappingItemList) {
                        string attrName = "";
                        string attr = "";
                        switch (item.curSrcItem.Value) {
                            case "EXPSEQUENCE":
                            case "USER.ID":
                            case "USER.NAME":
                            case "ROOTID":
                            case "NOW":
                            case "EXPDATE":
                            case "CONFIGSTATUS":
                            case "SSOID":
                                row[item.DestCol] = GetDefaultValue(item.DestDataType, item.curSrcItem.Value, item);
                                break;
                            default:
                                if (hsIndex[item.DestCol] == null || String.IsNullOrEmpty(item.curSrcItem.Value)) {
                                    if (!string.IsNullOrEmpty(item.curSrcItem.DefaultValue.ToString())) {
                                        if (item.DestDataType == DBDataType.DateTime) {
                                            try {
                                                row[item.DestCol] = Convert.ToDateTime(item.curSrcItem.DefaultValue);
                                            } catch {
                                                Ercol.Append("BOM配置文件中" + item.curSrcItem.DefaultValue + "不是有效时间格式");
                                            }
                                        } else {
                                            try {
                                                row[item.DestCol] = GetDefaultValue(item.DestDataType,
                                                                                    item.curSrcItem.DefaultValue, item);
                                            } catch (Exception ett) {
                                                string strValues = "配置文件中对应中间库表'" + mapping.TableName + "." +
                                                                   item.DestCol +
                                                                   "'对应的属性默认值 " + item.curSrcItem.DefaultValue +
                                                                   "的设置有问题";
                                                strValues += "：" + ett.Message;
                                                throw new PLMException("\r\n" + strValues + "\r\n导出异常中止！");
                                            }
                                        }
                                    }
                                    break;
                                }
                                index = Convert.ToInt32(hsIndex[item.DestCol]);
                                if (item.curSrcItem.Value == "PID")
                                    pId = deItem.ConvertOracle(item.DestDataType, dr, index).ToString();
                                if (item.curSrcItem.Value == "CID")
                                    cId = deItem.ConvertOracle(item.DestDataType, dr, index).ToString();
                                if (item.curSrcItem.Value.IndexOf('_') != -1)
                                    attr = item.curSrcItem.Value.Substring(item.curSrcItem.Value.LastIndexOf('_') + 1);
                                if (item.curSrcItem.Value == "NUMBER")
                                    attrName = "装配数量";
                                else if (attr != "") {
                                    if (item.curSrcItem.Value.StartsWith("R.")) {
                                        attrName = "装配关系";
                                        deAttr = ModelContext.MetaModel.GetRelationAttribute("PARTTOPART", attr);
                                        if (deAttr != null)
                                            attrName += ": " + deAttr.Label;
                                        else
                                            attrName += item.curSrcItem.Value;
                                    } else if (item.curSrcItem.Value.StartsWith("PS.")) {
                                        attrName = item.curSrcItem.Value;
                                    } else {
                                        if (item.curSrcItem.Value.StartsWith("C."))
                                            attrName = "子件:";
                                        else if (item.curSrcItem.Value.StartsWith("P."))
                                            attrName = "父件:";

                                        if (item.curSrcItem.Value.IndexOf("_R_") != -1 ||
                                            item.curSrcItem.Value.IndexOf("_M_") != -1) {
                                            fixAttr = FixedAttribute.GetFixedAttr(attr);
                                            if (fixAttr != null)
                                                attrName += fixAttr.AttrLabel;
                                            else
                                                attrName += item.curSrcItem.Value;
                                        } else {
                                            deAttr = ModelContext.MetaModel.GetAttribute("PART", attr);
                                            if (deAttr != null)
                                                attrName += deAttr.Label;
                                            else
                                                attrName += item.curSrcItem.Value;
                                        }
                                    }
                                }
                                if (attrName == "") attrName = item.curSrcItem.Value;
                                object src = null;
                                if (!dr.IsDBNull(index)) {
                                    try {
                                        src = deItem.ConvertOracle(item.curSrcItem.SrcDataType, dr, index);
                                        if (item.curSrcItem.SrcDataType == DBDataType.Varchar) {
                                            string s = src.ToString().Trim();
                                            if (string.IsNullOrEmpty(s) && !item.IsAllowNull) {
                                                Ercol.Append("  " + attrName + "为空; ");
                                                continue;
                                            }
                                        }
                                        if (item.curSrcItem.SrcDataType == DBDataType.DateTime) {
                                            if (((DateTime)src) == DateTime.MinValue &&
                                                !string.IsNullOrEmpty(item.curSrcItem.DefaultValue.ToString())) {
                                                src = Convert.ToDateTime(item.curSrcItem.DefaultValue);
                                            }
                                        }
                                    } catch (Exception cvt) {
                                        Ercol.Append("\r\n配置文件中对应中间库表'" + mapping.TableName + "." + item.DestCol +
                                                     "'对应的PLM属性的数据类型设置有问题");
                                        Ercol.Append("：" + cvt.Message);
                                    }
                                    try {
                                        if (!String.IsNullOrEmpty(item.DestCol) && src != null)
                                            row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType, src);
                                    } catch (Exception cvt) {
                                        Ercol.Append("\r\n" + attrName + "数据定义有问题" + cvt.Message + dr.GetValue(index));
                                    }
                                } else {
                                    if (!item.IsAllowNull) {
                                        if (item.curSrcItem.DefaultValue != null) {
                                            try {
                                                row[item.DestCol] = GetDefaultValue(item.DestDataType,
                                                                                    item.curSrcItem.DefaultValue, item);
                                            } catch {
                                                throw new Exception("Bom中" + mapping.TableName + "'" + item.DestCol +
                                                                    "'的缺省值" +
                                                                    item.curSrcItem.DefaultValue + "无法转换成所需要的数据类型，请检查文件");
                                            }
                                        } else {
                                            Ercol.Append("\r\n\t" + attrName + "数据为空");
                                        }
                                    }
                                }
                                break;
                        }

                        #region 处理数据转换

                        if (item.curSrcItem != null && item.curSrcItem.DataConvertTable.Count > 0) {
                            object unit = item.curSrcItem.DataConvertTable[row[item.DestCol]];
                            try {
                                if (unit != null)
                                    row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType, unit);
                            } catch {
                                Ercol.Append(" " + item.curSrcItem.SrcCol + "的值'" + row[item.DestCol] + "'的转换值" + unit +
                                             "无法转化为" + item.DestCol + "数据类型");
                            }
                        }

                        #endregion

                        if (item.DestDataType == DBDataType.Varchar && row[item.DestCol] != DBNull.Value) {
                            string cut_obj = row[item.DestCol].ToString();
                            if (item.i_Size > 0) {
                                if (cut_obj.Length > item.i_Size) {
                                    cut_obj = GetStringforSize(cut_obj, item.i_Size);
                                    if (item.IsAllowCut)
                                        row[item.DestCol] = cut_obj;
                                    else {
                                        Ercol.Append("\r\n" + attrName + "的取值“" + row[item.DestCol] + "”字符串超过允许范围");
                                    }
                                }
                            }
                        }
                    }
                    if (Ercol.Length > 0) {
                        strEr.Append("\r\n BOM导出数据");
                        strEr.Append(mapping.TableName + " 父件<" + pId + ">");
                        if (!string.IsNullOrEmpty(cId))
                            strEr.Append("子件<" + cId + ">错误：");
                        strEr.Append(Ercol.ToString());
                    }
                    bomTable.Rows.Add(row);
                }
            } catch (Exception ex) {
                strEr.Append("\n" + ex.Message);
            } finally {
                if (dr != null)
                    dr.Close();
            }
            CheckDataUnique(bomTable, mapping);
            return bomTable;
        }

        private object GetDefaultValue(DBDataType dType, object objValue, ColumnMappingItem item) {
            string strObj = objValue.ToString().ToUpper();
            object obj = null;
            switch (strObj) {
                case "EXPSEQUENCE":
                    obj = CurSeq;
                    IsFindSeq = true;
                    break;
                case "USER.ID":
                    obj = user.LogId;
                    break;
                case "USER.NAME":
                    obj = user.Name;
                    break;
                case "ROOTID":
                    obj = RootID;
                    break;
                case "NOW":
                case "EXPDATE":
                    if (dType == DBDataType.DateTime)
                        obj = now;
                    else {
                        if (item.curSrcItem.Format != null)
                            obj = now.ToString(item.curSrcItem.Format);
                        else
                            obj = now.ToString();
                    }
                    break;
                case "CONFIGSTATUS":
                    if (configName != "")
                        obj = configName;
                    break;
                case "SSOID":
                    obj = expEvent.Oid.ToString();
                    break;
                default:

                    obj = deItem.ConvertToSystem(dType, objValue);
                    break;
            }
            return obj;
        }

        /*
                /// <summary>
                /// 从系统中获取需要导出的工艺路线信息，并构造成工艺路线表准备导出到ERP中
                /// </summary>
                /// <returns>表</returns>
                private DataTable BuildRouteTable(object obj)
                {
                    RelItemColumnMapping routeMapping = obj as RelItemColumnMapping;
                    if( string.IsNullOrEmpty(routeMapping.TableName) || !routeMapping.IsExport ) return null;
                    DataTable routeTable = new DataTable(routeMapping.TableName);
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = (OracleConnection) dbParam.Connection;
                    StringBuilder sql = new StringBuilder();
                    StringBuilder er1 = new StringBuilder();
                    StringBuilder er2 = new StringBuilder();
                    string strPLMCol;
                    Hashtable hsIndex, hsIndexCol = new Hashtable();
                    OracleDataReader dr = null;
                    string right_M_Name, right_I_Name, relName, left_I_Name;
                    StringBuilder strOrder = new StringBuilder();
                    foreach( ColumnMappingItem item in routeMapping.MappingItemList )
                    {
                        if( String.IsNullOrEmpty(item.DestCol) ) continue;
                        DataColumn col = new DataColumn(item.DestCol, deItem.ConvertSystemType(item.DestDataType));
                        routeTable.Columns.Add(col);
                        if( item.curSrcItem == null )
                        {
                            if( item.hsSrcItems.Count > 0 )
                                item.curSrcItem = item.hsSrcItems[item.lstClassName[0]] as SourceItem;
                        }
                    }
                    cmd.Parameters.Add(":SessionOid", OracleDbType.Raw).Value = sOid.ToByteArray();

                    DataColumn colRel = new DataColumn("RELNAME");
                    routeTable.Columns.Add(colRel);
                    int idx;
                    Hashtable hslinks = new Hashtable(); //存放定位信息
                    foreach( ArrayList lst in routeMapping.lstRelationLink )
                    {
                        idx = 0;
                        sql.Remove(0, sql.Length);
                        hslinks.Clear();
                        //根据关联获取数据
                        hsIndex = routeMapping.hsClass[lst] as Hashtable;
                        hsIndexCol.Clear();
                        sql.Append("SELECT DISTINCT  ");
                        strOrder.Remove(0, strOrder.Length);
                        foreach( ColumnMappingItem item in routeMapping.MappingItemList )
                        {
                            strPLMCol = routeMapping.GetPLMColName(item, hsIndex);
                            if( item.curSrcItem == null || item.curSrcItem.SrcType == 0 )
                                continue;
                            if( String.IsNullOrEmpty(strPLMCol) )
                                continue;
                            sql.Append(strPLMCol + " " + item.curSrcItem.PLMCol);
                            sql.Append(",");
                            hsIndexCol.Add(item.DestCol, idx++);
                            SetOrderSql(strPLMCol, item.ordTy, strOrder);
                        }
                        sql.Remove(sql.Length - 1, 1);
                        sql.Append(" FROM PLM_PSM_PS T,");
                        sql.Append(" PLM_PSM_ITEMMASTER_REVISION M0,PLM_CUSV_" + routeMapping.SrcStartName + " I0");
                        string linkListName = ModelContext.MetaModel.GetClassLabel(routeMapping.SrcStartName);
                        int index;
                        foreach( RelationLink link in lst )
                        {
                            sql.Append(",");
                            index = Convert.ToInt32(hsIndex[link.RelationName]);
                            sql.Append(" PLM_CUS_R_" + link.RelationName + " R" + index + ",");
                            linkListName += "-->" + ModelContext.MetaModel.GetRelation(link.RelationName).Label;
                            sql.Append("PLM_PSM_ITEMMASTER_REVISION M" + index + ",");
                            sql.Append("PLM_CUSV_" + link.RightClassName + " I" + index);
                            linkListName += "-->" + ModelContext.MetaModel.GetClassLabel(link.RightClassName);
                        }
                        sql.Append(
                            " WHERE T.PLM_CITEROID = I0.PLM_OID AND  T.PLM_SESSIONOID =:SessionOid AND T.PLM_CMASTEROID = M0.PLM_M_OID ");

                        foreach( RelationLink link in lst )
                        {
                            sql.Append(" AND ");
                            index = Convert.ToInt32(hsIndex[link.RelationName]);
                            relName = "R" + index;
                            index = Convert.ToInt32(hsIndex[link.RightClassName]);
                            right_M_Name = "M" + index;
                            right_I_Name = "I" + index;
                            index = index - 1;
                            left_I_Name = "I" + index;

                            sql.Append(left_I_Name + ".PLM_OID  = " + relName + ".PLM_LEFTOBJ AND ");
                            sql.Append(relName + ".PLM_RIGHTOBJ = " + right_M_Name + ".PLM_M_OID AND ");

                            sql.Append(right_M_Name + ".PLM_M_LASTREVISION=" + right_M_Name + ".PLM_R_REVISION AND ");
                            sql.Append(right_M_Name + ".PLM_R_OID=" + right_I_Name + ".PLM_REVISIONOID AND ");
                            sql.Append(right_M_Name + ".PLM_R_LASTITERATION=" + right_I_Name + ".PLM_ITERATION AND ");
                            sql.Append(right_M_Name + ".PLM_M_STATE <> 'O' "); // AND ");
                        }
                        //		sql.Append(")");
                        if( strOrder.Length > 0 )
                            sql.Append(strOrder.ToString());
                        cmd.CommandText = sql.ToString();
                        Hashtable hsId = new Hashtable();
                        Hashtable hsId2 = new Hashtable();

                        #region 查询数据

                        try
                        {
                            try
                            {
                                dr = cmd.ExecuteReader();
                            }
                            catch( Exception ex1 )
                            {
                                throw new Exception("获取工艺路线的sql语句有问题 " + ex1.Message + "：\r\n" + sql);
                            }
                            while( dr.Read() )
                            {
                                DataRow row = routeTable.NewRow();
                                er1.Remove(0, er1.Length);
                                er2.Remove(0, er2.Length);
                                GetFillData(row, hsIndexCol, routeMapping, dr, er1, er2);
                                row[colRel] = linkListName;
                                string id = row[routeMapping.ID_Item.DestCol].ToString();
                                if( er1.Length == 0 )
                                {
                                    routeTable.Rows.Add(row);
                                }
                                else
                                {
                                    if( hsId.Contains(id) )
                                    {
                                        er1.Insert(0, hsId[id].ToString());
                                    }
                                    hsId[id] = er1.ToString();
                                }
                                if( er2.Length > 0 )
                                {
                                    if( hsId2.Contains(id) )
                                    {
                                        er2.Insert(0, hsId2[id].ToString());
                                    }
                                    hsId2[id] = er2.ToString();
                                }
                            }
                            IDictionaryEnumerator IE;
                            if( hsId.Count > 0 )
                            {
                                strEr.Append("\r\n通过关联" + linkListName + "获取工艺数据，发现错误:\r\n");
                                IE = hsId.GetEnumerator();
                                IE.Reset();
                                while( IE.MoveNext() )
                                {
                                    strEr.Append("   物料代号为：<" + IE.Key + ">所关联的数据:\r\n");
                                    strEr.Append(IE.Value + "\r\n");
                                }
                            }
                            if( hsId2.Count > 0 )
                            {
                                strWarn.Append("\r\n通过关联" + linkListName + "获取工艺数据，发现非致命性错误:\r\n");
                                IE = hsId2.GetEnumerator();
                                IE.Reset();
                                while( IE.MoveNext() )
                                {
                                    strWarn.Append("   物料代号为：<" + IE.Key + ">所关联的数据:\r\n");
                                    strWarn.Append(IE.Value + "\r\n");
                                }
                            }
                        }
                        catch( Exception ex )
                        {
                            strEr.Append("\r\n " + ex.Message);
                        }
                        finally
                        {
                            if( dr != null )
                                dr.Close();
                        }

                        #endregion
                    }
                    CheckDataUnique(routeTable, routeMapping);

                    return routeTable;
                }
    
        */

        /// <summary>
        /// 从系统中获取需要导出的工艺路线信息，并构造成工艺路线表准备导出到ERP中
        /// </summary>
        /// <returns>表</returns>
        private DataTable BuildRouteTable(object obj) {
            RelItemColumnMapping routeMapping = obj as RelItemColumnMapping;
            if (string.IsNullOrEmpty(routeMapping.TableName) || !routeMapping.IsExport) return null;
            DataTable routeTable = new DataTable(routeMapping.TableName);
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = (OracleConnection)dbParam.Connection;
            StringBuilder sql = new StringBuilder();
            StringBuilder er1 = new StringBuilder();
            StringBuilder er2 = new StringBuilder();
            string strPLMCol;
            Hashtable hsIndex, hsIndexCol = new Hashtable();
            OracleDataReader dr = null;
            string right_M_Name, right_I_Name, relName, left_I_Name;
            StringBuilder strOrder = new StringBuilder();
            foreach (ColumnMappingItem item in routeMapping.MappingItemList) {
                if (String.IsNullOrEmpty(item.DestCol)) continue;
                DataColumn col = new DataColumn(item.DestCol, deItem.ConvertSystemType(item.DestDataType));
                routeTable.Columns.Add(col);
                if (item.curSrcItem == null) {
                    if (item.hsSrcItems.Count > 0)
                        item.curSrcItem = item.hsSrcItems[item.lstClassName[0]] as SourceItem;
                }
            }
            cmd.Parameters.Add(":SessionOid", OracleDbType.Raw).Value = sOid.ToByteArray();

            DataColumn colRel = new DataColumn("RELNAME");
            routeTable.Columns.Add(colRel);
            int idx;
            Hashtable hslinks = new Hashtable(); //存放定位信息

            IDictionaryEnumerator ic = routeMapping.hsLinks.GetEnumerator();
            ic.Reset();
            while (ic.MoveNext()) {
                ArrayList lst = ic.Value as ArrayList;
                string linkNum = ic.Key.ToString();

                idx = 0;
                sql.Remove(0, sql.Length);
                hslinks.Clear();
                //根据关联获取数据
                hsIndex = routeMapping.hsClass[linkNum] as Hashtable;
                hsIndexCol.Clear();
                sql.Append("SELECT DISTINCT  ");
                strOrder.Remove(0, strOrder.Length);
                foreach (ColumnMappingItem item in routeMapping.MappingItemList) {
                    strPLMCol = routeMapping.GetPLMColName(item, linkNum);
                    if (item.curSrcItem == null || item.curSrcItem.SrcType == 0)
                        continue;
                    if (String.IsNullOrEmpty(strPLMCol))
                        continue;
                    sql.Append(strPLMCol + " " + item.curSrcItem.PLMCol);
                    sql.Append(",");
                    hsIndexCol.Add(item.DestCol, idx++);
                    SetOrderSql(strPLMCol, item.ordTy, strOrder);
                }

                sql.Remove(sql.Length - 1, 1);
                sql.Append(" FROM PLM_PSM_PS T,");
                sql.Append(" PLM_PSM_ITEMMASTER_REVISION M0,PLM_CUSV_" + routeMapping.SrcStartName + " I0");
                string linkListName = ModelContext.MetaModel.GetClassLabel(routeMapping.SrcStartName);
                int index;
                foreach (RelationLink link in lst) {
                    sql.Append(",");
                    index = Convert.ToInt32(hsIndex[link.RelationName]);
                    sql.Append(" PLM_CUS_R_" + link.RelationName + " R" + index + ",");
                    linkListName += "-->" + ModelContext.MetaModel.GetRelation(link.RelationName).Label;
                    sql.Append("PLM_PSM_ITEMMASTER_REVISION M" + index + ",");
                    sql.Append("PLM_CUSV_" + link.RightClassName + " I" + index);
                    linkListName += "-->" + ModelContext.MetaModel.GetClassLabel(link.RightClassName);
                }
                sql.Append(
                    " WHERE T.PLM_CITEROID = I0.PLM_OID AND  T.PLM_SESSIONOID =:SessionOid AND T.PLM_CREVOID = M0.PLM_R_OID ");

                foreach (RelationLink link in lst) {
                    sql.Append(" AND ");
                    index = Convert.ToInt32(hsIndex[link.RelationName]);
                    relName = "R" + index;
                    index = Convert.ToInt32(hsIndex[link.RightClassName]);
                    right_M_Name = "M" + index;
                    right_I_Name = "I" + index;
                    index = index - 1;
                    left_I_Name = "I" + index;

                    sql.Append(left_I_Name + ".PLM_OID  = " + relName + ".PLM_LEFTOBJ AND ");
                    sql.Append(relName + ".PLM_RIGHTOBJ = " + right_M_Name + ".PLM_M_OID AND ");

                    sql.Append(right_M_Name + ".PLM_M_LASTREVISION=" + right_M_Name + ".PLM_R_REVISION AND ");
                    sql.Append(right_M_Name + ".PLM_R_OID=" + right_I_Name + ".PLM_REVISIONOID AND ");
                    sql.Append(right_M_Name + ".PLM_R_LASTITERATION=" + right_I_Name + ".PLM_ITERATION  ");
                    // sql.Append(" AND "+right_M_Name + ".PLM_M_STATE <> 'O' "); // AND ");
                }
                //		sql.Append(")");
                if (strOrder.Length > 0)
                    sql.Append(strOrder.ToString());
                cmd.CommandText = sql.ToString();
                Hashtable hsId = new Hashtable();
                Hashtable hsId2 = new Hashtable();

                #region 查询数据

                try {
                    try {
                        dr = cmd.ExecuteReader();
                    } catch (Exception ex1) {
                        throw new Exception("获取工艺路线的sql语句有问题 " + ex1.Message + "：\r\n" + sql);
                    }
                    while (dr.Read()) {
                        DataRow row = routeTable.NewRow();
                        er1.Remove(0, er1.Length);
                        er2.Remove(0, er2.Length);
                        GetFillData(row, hsIndexCol, routeMapping, dr, er1, er2);
                        row[colRel] = linkListName;
                        string id = row[routeMapping.PART_ID_Item.DestCol].ToString();
                        if (er1.Length == 0) {
                            routeTable.Rows.Add(row);
                        } else {
                            if (hsId.Contains(id)) {
                                er1.Insert(0, hsId[id].ToString());
                            }
                            hsId[id] = er1.ToString();
                        }
                        if (er2.Length > 0) {
                            if (hsId2.Contains(id)) {
                                er2.Insert(0, hsId2[id].ToString());
                            }
                            hsId2[id] = er2.ToString();
                        }
                    }
                    IDictionaryEnumerator IE;
                    if (hsId.Count > 0) {
                        strEr.Append("\r\n通过关联" + linkListName + "获取工艺数据，发现错误:\r\n");
                        IE = hsId.GetEnumerator();
                        IE.Reset();
                        while (IE.MoveNext()) {
                            strEr.Append("   物料代号为：<" + IE.Key + ">所关联的数据:\r\n");
                            strEr.Append(IE.Value + "\r\n");
                        }
                    }
                    if (hsId2.Count > 0) {
                        strWarn.Append("\r\n通过关联" + linkListName + "获取工艺数据，发现非致命性错误:\r\n");
                        IE = hsId2.GetEnumerator();
                        IE.Reset();
                        while (IE.MoveNext()) {
                            strWarn.Append("   物料代号为：<" + IE.Key + ">所关联的数据:\r\n");
                            strWarn.Append(IE.Value + "\r\n");
                        }
                    }
                } catch (Exception ex) {
                    strEr.Append("\r\n " + ex.Message);
                } finally {
                    if (dr != null)
                        dr.Close();
                }

                #endregion
            }
            CheckDataUnique(routeTable, routeMapping);

            return routeTable;
        }


        private void GetFillData(DataRow row, Hashtable hsIndex, ItemColumnMapping mapping, OracleDataReader dr,
                                 StringBuilder er1, StringBuilder er2) {
            string MapType = mapping as RelItemColumnMapping == null ? "物料导出" : "工艺路线导出";
            bool IsItem = mapping as RelItemColumnMapping == null ? true : false;
            string strValues = "";
            DEMetaAttribute attr;
            string clsName;
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                int index = -1;
                if (hsIndex[item.DestCol] != null) {
                    index = Convert.ToInt32(hsIndex[item.DestCol]);
                }

                #region 判断那些指定为不允许为空的列

                if (!item.IsAllowNull) {
                    if (index == -1) {
                        if (item.curSrcItem == null)
                            throw new PLMException("请检查BOM导出的配置文件中关于" + MapType + item.DestCol +
                                                   "的设置是否正确，是否应当允许为空，导出异常中止！");
                        if (item.curSrcItem.Value == null || item.curSrcItem.Value.Trim() == "")
                            throw new PLMException("\r\n 配置文件中目标属性“" + item.DestCol +
                                                   "”定义为不为空，但未给定取值(Value)或取值为空，导出异常中止！");
                    } else if (dr.IsDBNull(index) && item.curSrcItem.Value == null) {
                        attr =
                            ModelContext.MetaModel.GetAttribute(item.curSrcItem.SrcClass,
                                                                item.curSrcItem.SrcCol.Substring(6));
                        if (attr == null) {
                            attr = ModelContext.MetaModel.GetRelationAttribute(item.curSrcItem.SrcClass,
                                                                               item.curSrcItem.SrcCol.Substring(6));
                        }
                        clsName = ModelContext.MetaModel.GetClassLabel(item.curSrcItem.SrcClass);
                        if (string.IsNullOrEmpty(clsName))
                            clsName = ModelContext.MetaModel.GetRelation(item.curSrcItem.SrcClass).Label;
                        if (attr != null)
                            er1.Append("  " + clsName + "." + attr.Label + "为空; ");
                        else
                            er1.Append(" " + item.curSrcItem.SrcCol + "为空！");
                        continue;
                    }
                }

                #endregion

                #region 给属性赋值

                if (item.curSrcItem == null) {
                    continue;
                }
                switch (item.curSrcItem.SrcType) {
                    case 0:
                        switch (item.curSrcItem.Value) {
                            case "USER.ID":
                            case "EXPSEQUENCE":
                            case "ROOTID":
                            case "SSOID":
                            case "CONFIGSTATUS":
                                row[item.DestCol] = GetDefaultValue(item.DestDataType, item.curSrcItem.Value, item);
                                break;

                            case "EXPDATE":
                            case "NOW":
                                row[item.DestCol] = GetDefaultValue(item.DestDataType, item.curSrcItem.Value, item);
                                if (IsItem)
                                    colItemTime = item.DestCol;
                                colItemTimeType = item.DestDataType;
                                break;

                            default:
                                if (!string.IsNullOrEmpty(item.curSrcItem.Value)) {
                                    if (item.DestDataType == DBDataType.DateTime) {
                                        try {
                                            row[item.DestCol] = Convert.ToDateTime(item.curSrcItem.Value);
                                        } catch {
                                            if (IsItem)
                                                er1.Append("物料配置文件中" + item.curSrcItem.Value + "不是有效时间格式");
                                            else
                                                er1.Append("工艺有关配置文件中" + item.curSrcItem.Value + "不是有效时间格式");
                                        }
                                    } else {
                                        try {
                                            row[item.DestCol] = GetDefaultValue(item.DestDataType, item.curSrcItem.Value,
                                                                                item);
                                        } catch (Exception ett) {
                                            strValues = "配置文件中对应中间库表'" + mapping.TableName + "." + item.DestCol +
                                                        "'对应的属性默认值 " + item.curSrcItem.Value + "的设置有问题";
                                            strValues += "：" + ett.Message;
                                            throw new PLMException("\r\n" + strValues + "\r\n导出异常中止！");
                                        }
                                    }
                                }
                                //   throw new PLMException("配置文件中目标属性“" + item.DestCol + "”的取值“" + item.curSrcItem.Value +
                                //                         "”无法识别，导出异常中止！");
                                break;
                        }
                        break;
                    case 1:
                        if (index != -1 && !dr.IsDBNull(index)) {
                            object src = null;
                            try {
                                src = deItem.ConvertOracle(item.curSrcItem.SrcDataType, dr, index);
                            } catch (Exception cvt) {
                                strValues = "配置文件中对应中间库表'" + mapping.TableName + "." + item.DestCol +
                                            "'对应的PLM属性的数据类型设置有问题";
                                strValues += "：" + cvt.Message;
                                if (item.IsAllowConvertErr) {
                                    if (!item.IsAllowNull) {
                                        if (item.curSrcItem.Value != null)
                                            src = item.curSrcItem.Value;
                                        else {
                                            er1.Append(strValues + "并且没有给定默认值");
                                            continue;
                                        }
                                    }
                                    er2.Append(strValues);
                                } else
                                    er1.Append(strValues);
                            }
                            attr =
                                ModelContext.MetaModel.GetAttribute(item.curSrcItem.SrcClass,
                                                                    item.curSrcItem.SrcCol.Substring(6));
                            clsName = ModelContext.MetaModel.GetClassLabel(item.curSrcItem.SrcClass);
                            try {
                                if (item.curSrcItem.SrcDataType == DBDataType.Varchar) {
                                    string s = src.ToString().Trim();
                                    if (string.IsNullOrEmpty(s) && !item.IsAllowNull) {
                                        if (attr != null)
                                            er1.Append("  " + clsName + "." + attr.Label + "为空; ");
                                        else
                                            er1.Append(" " + item.curSrcItem.SrcCol + "为空！");
                                        continue;
                                    }
                                }
                                if (src != null)
                                    row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType, src);
                            } catch (Exception cvt2) {
                                if (item.IsAllowConvertErr) {
                                    if (attr != null)
                                        er2.Append("  " + clsName + "." + attr.Label + "的值'" + src + "'无法转化为" +
                                                   item.DestCol +
                                                   "所要求的数据类型;");
                                    else
                                        er2.Append(" " + item.curSrcItem.SrcCol + "的值'" + src + "'无法转化为" + item.DestCol +
                                                   "所要求的数据类型;");
                                    try {
                                        if (item.curSrcItem.Value != null)
                                            row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType,
                                                                                       item.curSrcItem.Value);
                                    } catch (Exception cvt) {
                                        strValues = "配置文件中" + mapping.TableName + "给定的默认值" + item.curSrcItem.Value +
                                                    "数据类型转换出错";
                                        strValues += "：" + cvt.Message;

                                        throw new PLMException(strValues, cvt, true);
                                    }
                                } else if (attr != null)
                                    er1.Append("  " + clsName + "." + attr.Label + "的值'" + src + "'无法转化为" + item.DestCol +
                                               "所要求的数据类型;");
                                else
                                    er1.Append(" " + item.curSrcItem.SrcCol + "的值'" + src + "'无法转化为" + item.DestCol +
                                               "所要求的数据类型;");

                                continue;
                            }
                            //增加对字符串型数据的判断
                        } else {
                            try {
                                if (item.curSrcItem.Value != null)
                                    row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType, item.curSrcItem.Value);
                            } catch (Exception cvt) {
                                strValues = "配置文件中" + mapping.TableName + "给定的默认值" + item.curSrcItem.Value + "数据类型转换出错";
                                strValues += "：" + cvt.Message;
                                if (item.IsAllowNull && item.IsAllowConvertErr) {
                                    er2.Append(" " + strValues);
                                    continue;
                                } else {
                                    throw new PLMException(strValues, cvt, true);
                                }
                            }
                        }

                        break;
                    default:
                        break;
                }

                #endregion

                #region 处理数据转换

                if (item.curSrcItem != null && item.curSrcItem.DataConvertTable.Count > 0) {
                    object unit = item.curSrcItem.DataConvertTable[row[item.DestCol]];
                    try {
                        if (unit != null)
                            row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType, unit);
                    } catch {
                        if (item.IsAllowConvertErr) {
                            er2.Append(" " + item.curSrcItem.SrcCol + "的值'" + row[item.DestCol] + "'的转换值" + unit +
                                       "无法转化为" + item.DestCol + "数据类型");
                            try {
                                if (item.curSrcItem.Value != null)
                                    row[item.DestCol] = deItem.ConvertToSystem(item.DestDataType, item.curSrcItem.Value);
                            } catch (Exception cvt) {
                                strValues = "配置文件中" + mapping.TableName + "给定的默认值" + item.curSrcItem.Value + "数据类型转换出错";
                                strValues += "：" + cvt.Message;

                                throw new PLMException(strValues, cvt, true);
                            }
                        } else
                            er1.Append(" " + item.curSrcItem.SrcCol + "的值'" + row[item.DestCol] + "'的转换值" + unit +
                                       "无法转化为" + item.DestCol + "数据类型");
                    }
                }

                #endregion

                #region 处理数据截断

                if (item.DestDataType == DBDataType.Varchar && row[item.DestCol] != DBNull.Value) {
                    string cut_obj = row[item.DestCol].ToString();
                    if (item.i_Size > 0) {
                        if (cut_obj.Length > item.i_Size) {
                            cut_obj = GetStringforSize(cut_obj, item.i_Size);
                            if (item.IsAllowCut)
                                row[item.DestCol] = cut_obj;
                            else
                                er1.Append("\r\n" + item.curSrcItem.SrcClass + "的属性" + item.curSrcItem.SrcCol + "”的取值“" +
                                           row[item.DestCol] + "”字符串超过允许范围");
                        }
                    }
                }

                #endregion
            }
        }

        private string GetStringforSize(string obj, int size) {
            byte[] slg = Encoding.GetEncoding("GB2312").GetBytes(obj);
            if (slg.Length <= size)
                return obj;
            else {
                string str = Encoding.GetEncoding("GB2312").GetString(slg, 0, size);
                if (str.EndsWith("?"))
                    str = str.Substring(0, str.Length - 1);
                return str;
            }
        }


        /// <summary>
        /// 校验数据是否合法
        /// </summary>
        /// <param name="tb"></param>
        /// <param name="mapping"></param>
        private void CheckDataUnique(DataTable tb, ItemColumnMapping mapping) {
            try {
                if (tb == null || tb.Rows.Count == 0) return;
                ArrayList lstKey = mapping.GetKeyItem;
                if (lstKey == null || lstKey.Count == 0) return;
                StringBuilder er2 = new StringBuilder();
                Hashtable hsKeys = new Hashtable();
                string strRelName = "RELNAME";
                //			bool isCAPP = tb.Columns.Contains(strRelName);
                string strValue;

                DataRow row;
                for (int i = 0; i < tb.Rows.Count; i++) {
                    strValue = "";
                    row = tb.Rows[i];
                    for (int j = 0; j < lstKey.Count; j++) {
                        ColumnMappingItem item = lstKey[j] as ColumnMappingItem;
                        DEMetaAttribute attr = null;
                        if (item.curSrcItem.SrcClass != "" && item.curSrcItem.SrcCol != "")
                            attr =
                                ModelContext.MetaModel.GetAttribute(item.curSrcItem.SrcClass,
                                                                    item.curSrcItem.SrcCol.Substring(6,
                                                                                                     item.curSrcItem.
                                                                                                         SrcCol.
                                                                                                         Length - 6));
                        string clsName = ModelContext.MetaModel.GetClassLabel(item.curSrcItem.SrcClass);
                        string attrLable;
                        if (string.IsNullOrEmpty(clsName))
                            clsName = "";
                        if (item.curSrcItem.SrcCol == "M.PLM_M_ID")
                            attrLable = "代号";
                        else if (attr == null)
                            attrLable = item.DestCol;
                        else
                            attrLable = attr.Label;

                        string strColValue = "";
                        if (tb.Columns.Contains(item.DestCol)) {
                            strColValue = row[item.DestCol] == DBNull.Value ? "空" : row[item.DestCol].ToString();
                        }
                        strValue += clsName + attrLable + ":" + strColValue + " ";
                    }
                    ArrayList lstKeys;
                    if (hsKeys.Contains(strValue)) {
                        lstKeys = hsKeys[strValue] as ArrayList;
                    } else {
                        lstKeys = new ArrayList();
                    }
                    lstKeys.Add(row[strRelName].ToString());
                    hsKeys[strValue] = lstKeys;
                }

                IDictionaryEnumerator IE = hsKeys.GetEnumerator();
                IE.Reset();
                while (IE.MoveNext()) {
                    er2.Remove(0, er2.Length);
                    ArrayList lstKey2 = IE.Value as ArrayList;
                    if (lstKey2.Count <= 1) continue;
                    er2.Append("导出数据表" + mapping.TableName + "发现重复关键字\r\n");
                    er2.Append(IE.Key.ToString());
                    er2.Append("\r\n");
                    for (int k = 0; k < lstKey2.Count; k++) {
                        er2.Append(lstKey2[k] + "\r\n");
                    }
                    RecordErr(sOid, er2.ToString());
                }
            } catch (Exception ex) {
                throw new Exception("进行数据完整性检查出错", ex);
            }
        }

        #endregion

        #region 数据整理

        /*
        /// <summary>
        /// 检验根结点数据是否允许导出
        /// </summary>
        /// <param name="cls"></param>
        /// <returns></returns>
        private bool CheakItemCanExport(string cls)
        {
            ArrayList lst;
            if( deItem.lstItemClass.Contains(cls) )
                return true;
            else
            {
                foreach( string clsName in deItem.lstItemClass )
                {
                    lst = deItem.GetSubClass(clsName);
                    if( lst == null || lst.Count == 0 || !lst.Contains(cls) )
                        continue;
                    else
                        return true;
                }
                return false;
            }
        }


        /// <summary>
        /// 使用存储过程进行数据预处理
        /// </summary>
        /// <param name="funName"></param>
        /// <param name="soid"></param>
        private void DPERPData(string funName)
        {
            try
            {
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = (OracleConnection) dbParam.Connection;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = funName;
                cmd.Parameters.Add(":SOid", OracleDbType.Raw).Value = sOid.ToByteArray();
                cmd.ExecuteNonQuery();
            }
            catch( Exception ex )
            {
                throw new Exception("数据预处理过程中存储过程" + deItem.DP_FUN_NAME + "内部出错" + ex.Message, ex);
            }
        }*/

        private void DpBeforeExport(string funName, DEBusinessItem item, DEPSOption Opt, int maxlvl) {
            try {
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = (OracleConnection)dbParam.Connection;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = funName;
                cmd.Parameters.Add(":SOid", OracleDbType.Raw).Value = sOid.ToByteArray();
                cmd.Parameters.Add(":UserOid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":UserOid"].Value = user.Oid.ToByteArray();
                cmd.Parameters.Add(":MOid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":MOid"].Value = item.MasterOid.ToByteArray();
                cmd.Parameters.Add(":RevOid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":RevOid"].Value = item.RevOid.ToByteArray();
                cmd.Parameters.Add(":IrOid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":IrOid"].Value = item.IterOid.ToByteArray();
                cmd.Parameters.Add(":ClsName", OracleDbType.NVarchar2, ParameterDirection.Input);
                cmd.Parameters[":ClsName"].Value = item.ClassName;
                cmd.Parameters.Add(":ViewOid", OracleDbType.NVarchar2, ParameterDirection.Input);
                cmd.Parameters[":ViewOid"].Value = Opt.CurView.ToString();
                cmd.Parameters.Add(":ViewGuid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":ViewGuid"].Value = Opt.CurView.ToByteArray();
                cmd.Parameters.Add(":EffWay", OracleDbType.NChar, ParameterDirection.Input);
                cmd.Parameters[":EffWay"].Value = Convert.ToChar(DEPSOption.ConvertTo(Opt.ViewWay));
                cmd.Parameters.Add(":DeadLine", OracleDbType.Date, ParameterDirection.Input);
                cmd.Parameters[":DeadLine"].Value = Opt.DeadLine;
                cmd.Parameters.Add(":EffSetOid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":EffSetOid"].Value = Opt.EffectivitySet.ToByteArray();
                cmd.Parameters.Add(":EffTime", OracleDbType.Date, ParameterDirection.Input);
                cmd.Parameters[":EffTime"].Value = Opt.EffectiveTime;
                cmd.Parameters.Add(":EffSerial", OracleDbType.NVarchar2, ParameterDirection.Input);
                cmd.Parameters[":EffSerial"].Value = Opt.EffectiveSerial;
                cmd.Parameters.Add(":CfgStatus", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":CfgStatus"].Value = Opt.ConfigStatus.ToByteArray();
                cmd.Parameters.Add(":CfgStatusRev", OracleDbType.Int32, ParameterDirection.Input);
                cmd.Parameters[":CfgStatusRev"].Value = Opt.ConfigStatusRev;
                cmd.Parameters.Add(":SnapShotOid", OracleDbType.Raw, ParameterDirection.Input);
                cmd.Parameters[":SnapShotOid"].Value = Opt.ProductSnapshot.ToByteArray();
                cmd.Parameters.Add(":ContextOid", OracleDbType.Raw).Value = Opt.Context.ContextOid.ToByteArray();
                cmd.Parameters.Add(":ContextKeys", OracleDbType.Varchar2).Value = Opt.Context.KeyParams;
                cmd.Parameters.Add(":OccurWhenInvalid", OracleDbType.NVarchar2, ParameterDirection.Input);
                cmd.Parameters[":OccurWhenInvalid"].Value = Opt.OccurWhenInvalid ? "Y" : "N";
                cmd.Parameters.Add(":MaxLevel", OracleDbType.Int32, ParameterDirection.Input);
                cmd.Parameters[":MaxLevel"].Value = maxlvl;

                cmd.ExecuteNonQuery();
            } catch (Exception ex) {
                throw new Exception("数据预处理过程中存储过程" + deItem.DP_BEFORE_NAME + "内部出错" + ex.Message, ex);
            }
        }


        /// <summary>
        /// 使用存储过程进行导出后数据整理
        /// </summary>
        private void AfterExport() {
            OracleCommand cmd2 = new OracleCommand();
            cmd2.Connection = (OracleConnection)dbParam.Connection;
            cmd2.CommandType = CommandType.StoredProcedure;
            cmd2.CommandText = deItem.DP_END_NAME;
            OracleParameter p_userId = cmd2.Parameters.Add("USER_ID", OracleDbType.Varchar2);
            p_userId.Value = user.LogId;
            OracleParameter p_expTime = cmd2.Parameters.Add("EXPDATE", OracleDbType.Date);
            p_expTime.Value = now;
            try {
                cmd2.ExecuteNonQuery();
            } catch (Exception ex) {
                RecordErr(sOid, "数据导出成功，但对数据库的后续处理失败：" + ex);

                //	throw new PLMException("数据导出成功，但对数据库的后续处理失败：" + ex.Message, ex, true);
            }
        }


        /// <summary>
        /// 导出数据扩展预处理
        /// </summary>
        /// <param name="expEvent"></param>
        /// <returns></returns>
        private bool ExtendBeforeExport(string dllName, Object inObjs, out Type type) {
            Assembly ass;
            bool IsRight = true;
            type = null;
            try {
                //ExtendERPExport.dll
                string extendDllpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                extendDllpath += "\\" + dllName;
                if (File.Exists(extendDllpath)) {
                    ass = Assembly.LoadFrom(extendDllpath);
                    type = ass.GetType("ExtendERPExport.ExtendERPExport");
                    IsRight = (bool)type.InvokeMember("BeforeExport", BindingFlags.InvokeMethod, null, null,
                                                       new Object[] { user, 0, RootID, now, dbParam, expEvent, inObjs },
                                                       null);
                    if (!IsRight) {
                        StringBuilder strExtend =
                            (StringBuilder)
                            type.InvokeMember("ErrInfo",
                                              BindingFlags.GetField | BindingFlags.Public | BindingFlags.Static, null,
                                              null, null);
                        if (strExtend.Length > 0) {
                            strEr.Append(strExtend + "\r\n\r\n");
                        }
                    }
                }
                return IsRight;
            } catch (TargetException refTargEx) {
                strEr.Append("\r\n BeforeExport 方法 调用目标为空：" + refTargEx.Message);
                throw refTargEx.InnerException;
            } catch (TargetParameterCountException refparm) {
                strEr.Append("\r\n BeforeExport 方法 参数出错：" + refparm.Message);
                throw refparm.InnerException;
            } catch (TargetInvocationException refex) {
                strEr.Append("\r\n BeforeExport 方法 内部出错：" + refex.Message);
                throw refex.InnerException;
            } catch (Exception eb) {
                strEr.Append("数据导出预处理程序ExtendERPExport.BeforeExport出现错误" + eb.Message);
                throw eb.InnerException;
            }
        }

        /// <summary>
        /// 导出前对内存中数据进行处理
        /// </summary>
        /// <param name="type"></param>
        /// <param name="itemTable"></param>
        /// <param name="bomTable"></param>
        /// <param name="hsRoutetables"></param>
        /// <returns></returns>
        private bool ExtendMidExport(Type type, Hashtable hsItemTables, Hashtable hsBomTables, Hashtable hsRoutetables,
                                     IDbConnection dconn, IDbTransaction dtrans) {
            bool IsRight = true;
            try {
                if (type != null) {
                    IsRight =
                        (bool)
                        type.InvokeMember("MidExport",
                                          BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Static, null,
                                          null, new Object[] { hsItemTables, hsBomTables, hsRoutetables, dconn, dtrans },
                                          null);
                    if (!IsRight) {
                        StringBuilder strExtend =
                            (StringBuilder)
                            type.InvokeMember("ErrInfo",
                                              BindingFlags.GetField | BindingFlags.Public | BindingFlags.Static, null,
                                              null,
                                              null);
                        if (strExtend.Length > 0) {
                            strEr.Append(strExtend + "\r\n\r\n");
                        }
                    }
                }
                return IsRight;
            } catch (TargetException refTargEx) {
                strEr.Append("\r\n MidExport 方法 调用目标为空：" + refTargEx.Message);
                throw refTargEx.InnerException;
            } catch (TargetParameterCountException refparm) {
                strEr.Append("\r\n MidExport 方法 参数出错：" + refparm.Message);
                throw refparm.InnerException;
            } catch (TargetInvocationException refex) {
                strEr.Append("\r\n MidExport 方法 内部出错：" + refex.Message);
                throw refex.InnerException;
            } catch (Exception eb) {
                strEr.Append("数据导出后处理程序ExtendERPExport.AfterExport出现错误" + eb.Message);
                throw eb.InnerException;
            }
        }


        /// <summary>
        /// 导出数据扩展后处理
        /// </summary>
        /// <param name="type"></param>
        private bool ExtendAfterExport(Type type, out DataSet dsResult, out Object outObjs) {
            bool IsRight = true;
            dsResult = null;
            outObjs = null;
            try {
                if (type != null) {
                    IsRight =
                        (bool)
                        type.InvokeMember("AfterExport",
                                          BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Static, null,
                                          null, new Object[] { }, null);
                    if (!IsRight) {
                        StringBuilder strExtend =
                            (StringBuilder)
                            type.InvokeMember("ErrInfo",
                                              BindingFlags.GetField | BindingFlags.Public | BindingFlags.Static, null,
                                              null,
                                              null);
                        if (strExtend.Length > 0) {
                            strEr.Append(strExtend + "\r\n\r\n");
                        }
                    } else {
                        dsResult =
                            (DataSet)
                            type.InvokeMember("dsResult",
                                              BindingFlags.GetField | BindingFlags.Public | BindingFlags.Static, null,
                                              null,
                                              null);
                        outObjs =
                            type.InvokeMember("outObjs",
                                              BindingFlags.GetField | BindingFlags.Public | BindingFlags.Static, null,
                                              null,
                                              null);
                    }
                }
                return IsRight;
            } catch (TargetException refTargEx) {
                strEr.Append("\r\n AfterExport 方法 调用目标为空：" + refTargEx.Message);
                throw refTargEx.InnerException;
            } catch (TargetParameterCountException refparm) {
                strEr.Append("\r\n AfterExport 方法 参数出错：" + refparm.Message);
                throw refparm.InnerException;
            } catch (TargetInvocationException refex) {
                strEr.Append("\r\n AfterExport 方法 内部出错：" + refex.Message);
                throw refex.InnerException;
            } catch (Exception eb) {
                strEr.Append("数据导出后处理程序ExtendERPExport.AfterExport出现错误" + eb.Message);
                throw eb.InnerException;
            }
        }

        #endregion

        #region 更新中间表

        /// <summary>
        /// 更新ERP临时表
        /// </summary>
        /// <param name="conn">连接</param>
        /// /// <param name="trans">事务对象</param>
        private void UpdateTable(object objKey, object objtb, IDbConnection conn, IDbTransaction trans) {
            ItemColumnMapping mapping = objKey as ItemColumnMapping;
            DataTable tb = objtb as DataTable;
            if (tb == null || tb.Rows.Count == 0)
                return;
            ArrayList lstIds = new ArrayList();
            StringBuilder sql = new StringBuilder();
            OleDbCommand scmd = (OleDbCommand)conn.CreateCommand();
            scmd.Transaction = trans as OleDbTransaction;

            #region 删除目标表中相关数据

            if (mapping.IsUpdate) {
                for (int i = 0; i < tb.Rows.Count; i++) {
                    string id = tb.Rows[i][mapping.ID_Item.DestCol].ToString();
                    if (!lstIds.Contains(id))
                        lstIds.Add(id);
                }
                try {
                    if (lstIds.Count > 0) {
                        sql.Append(" DELETE FROM ");
                        sql.Append(mapping.TableName);
                        sql.Append(" where ");
                        sql.Append(mapping.ID_Item.DestCol + " = ? ");
                        scmd.CommandText = sql.ToString();
                        //		if (DBType=="OLEDB")		
                        OleDbParameter p = scmd.Parameters.Add("@PLM_" + mapping.ID_Item.DestCol,
                                                               deItem.ConvertToOleDBType(mapping.ID_Item.DestDataType));
                        for (int i = 0; i < lstIds.Count; i++) {
                            p.Value = lstIds[i].ToString();
                            scmd.ExecuteNonQuery();
                        }
                    }
                } catch (Exception ex) {
                    throw new Exception("删除中间表" + mapping.TableName + "中已有数据出错" + ex.Message + "\r\n" + ex + "\r\n出错语句：" +
                                        scmd.CommandText);
                }
            }

            #endregion

            #region 数据更新

            // 创建新的
            sql.Remove(0, sql.Length);
            sql.Append("INSERT INTO ");
            sql.Append(mapping.TableName);
            sql.Append("(");
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!tb.Columns.Contains(item.DestCol) || !item.IsExport)
                    continue;
                sql.Append(item.DestCol);
                sql.Append(",");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append(") values (");
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!tb.Columns.Contains(item.DestCol) || !item.IsExport)
                    continue;
                sql.Append("?");
                sql.Append(",");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append(")");
            scmd.CommandText = sql.ToString();
            scmd.Parameters.Clear();
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!tb.Columns.Contains(item.DestCol) || !item.IsExport)
                    continue;
                OleDbParameter p = new OleDbParameter("@" + item.DestCol, deItem.ConvertToOleDBType(item.DestDataType));
                scmd.Parameters.Add(p);
                p.SourceColumn = item.DestCol;
            }
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.InsertCommand = scmd;
            try {
                da.Update(tb);
            } catch (Exception ex) {
                PLMEventLog.WriteExceptionLog(ex);
                string str2 = "更新中间表" + mapping.TableName + " 失败:" + ex.Message + ":" + ex.InnerException;
                str2 += "\r\n 链接中间表 的链接字符串" + conn.ConnectionString;
                str2 += "\r\n插入语句：" + da.InsertCommand.CommandText;
                RecordErr(sOid, str2);
            } finally {
                if (scmd != null) scmd.Dispose();
                da.Dispose();
            }

            #endregion
        }

        /// <summary>
        /// 更新物料表
        /// </summary>
        /// <param name="mapping"></param>
        /// <param name="conn"></param>
        /// <param name="trans"></param>
        private void UpdateItemTable(ItemColumnMapping mapping, DataTable tb, IDbConnection conn, IDbTransaction trans) {
            if (tb == null || tb.Rows.Count == 0)
                return;
            if (!mapping.IsUpdate) {
                UpdateTable(mapping, tb, conn, trans);
                return;
            }
            //			ArrayList lstIds = new ArrayList();
            StringBuilder sql = new StringBuilder();
            OleDbCommand icmd = (OleDbCommand)conn.CreateCommand();
            OleDbCommand ucmd = (OleDbCommand)conn.CreateCommand();
            OleDbCommand scmd = (OleDbCommand)conn.CreateCommand();
            OleDbCommand dcmd = (OleDbCommand)conn.CreateCommand();
            scmd.Transaction = trans as OleDbTransaction;
            ucmd.Transaction = trans as OleDbTransaction;
            icmd.Transaction = trans as OleDbTransaction;
            //			dcmd.Transaction = trans as OleDbTransaction;

            #region 获取已有数据

            sql.Remove(0, sql.Length);
            sql.Append(" select  * ");
            //sql.Append(mapping.ID_Item.DestCol);
            sql.Append(" from ");
            sql.Append(mapping.TableName);

            scmd.CommandText = sql.ToString();
            //				OleDbDataReader rd = null;
            //			try
            //			{
            //				rd = scmd.ExecuteReader();
            //				while (rd.Read())
            //				{
            //					string id = rd.GetString(0);
            //					if (!lstIds.Contains(id))
            //						lstIds.Add(id);
            //				}
            //			}
            //			finally
            //			{
            //				if (rd != null)
            //					rd.Close();
            //			}
            //				//人为改变数据表中数据状态
            //				for (int i = 0; i < tb.Rows.Count; i++)
            //				{
            //					string id = tb.Rows[i][mapping.ID_Item.DestCol].ToString();
            //					if (lstIds.Contains(id))
            //					{
            //						tb.Rows[i].AcceptChanges();
            //						tb.Rows[i][mapping.ID_Item.DestCol] = id;
            //					}
            //				}

            #endregion

            #region 更新已有数据

            sql.Remove(0, sql.Length);
            sql.Append("update " + mapping.TableName + " set ");
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (item.DestCol == mapping.ID_Item.DestCol)
                    continue;
                if (!tb.Columns.Contains(item.DestCol))
                    continue;
                if (!item.IsExport)
                    continue;
                sql.Append(item.DestCol);
                sql.Append(" = ?,");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append("  where ");
            sql.Append(mapping.ID_Item.DestCol);
            sql.Append(" = ?  ");

            ucmd.CommandText = sql.ToString();
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (item.DestCol == mapping.ID_Item.DestCol)
                    continue;
                if (!tb.Columns.Contains(item.DestCol))
                    continue;
                if (!item.IsExport)
                    continue;
                OleDbParameter p = ucmd.Parameters.Add("@" + item.DestCol, deItem.ConvertToOleDBType(item.DestDataType));
                p.SourceColumn = item.DestCol;
            }
            OleDbParameter pId =
                ucmd.Parameters.Add("@" + mapping.ID_Item.DestCol,
                                    deItem.ConvertToOleDBType(mapping.ID_Item.DestDataType));
            pId.SourceColumn = mapping.ID_Item.DestCol;

            #endregion

            #region 删除数据

            //			sql.Remove(0, sql.Length);
            //			sql.Append("delete from " + mapping.TableName );
            //			foreach (ColumnMappingItem item in mapping.MappingItemList)
            //			{
            //				if (item.DestCol == mapping.ID_Item.DestCol)
            //					continue;
            //				if(!tb.Columns.Contains(item.DestCol ))
            //					continue;
            //				if(!item.IsExport)
            //					continue;
            //				sql.Append(item.DestCol);
            //				sql.Append(" = ?,");
            //			}
            //			sql.Remove(sql.Length - 1, 1);
            //			sql.Append("  where ");
            //			sql.Append(mapping.ID_Item.DestCol);
            //			sql.Append(" = ?  ");
            //			OleDbParameter dId =
            //				dcmd.Parameters.Add("@" + mapping.ID_Item.DestCol, deItem.ConvertToOleDBType(mapping.ID_Item.DestDataType));
            //			dId.SourceColumn = mapping.ID_Item.DestCol;
            //			if(colItemTime !="")
            //			{
            //				sql.Append(" and ");
            //				sql.Append(colItemTime);
            //				sql.Append(" = ?  ");
            //				OleDbParameter dTime =
            //					dcmd.Parameters.Add("@" + colItemTime, deItem.ConvertToOleDBType(colItemTimeType));
            //				dTime.SourceColumn = colItemTime;
            //			}
            //			dcmd.CommandText = sql.ToString();

            #endregion

            #region  插入新数据

            // 创建新的
            sql.Remove(0, sql.Length);
            sql.Append("INSERT INTO ");
            sql.Append(mapping.TableName);
            sql.Append("(");
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!tb.Columns.Contains(item.DestCol))
                    continue;
                if (!item.IsExport)
                    continue;
                sql.Append(item.DestCol);
                sql.Append(",");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append(") values (");

            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!tb.Columns.Contains(item.DestCol))
                    continue;
                if (!item.IsExport)
                    continue;
                sql.Append("?");
                sql.Append(",");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append(")");
            icmd.CommandText = sql.ToString();
            icmd.Parameters.Clear();
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                if (!tb.Columns.Contains(item.DestCol))
                    continue;
                if (!item.IsExport)
                    continue;
                OleDbParameter p = new OleDbParameter("@" + item.DestCol, deItem.ConvertToOleDBType(item.DestDataType));
                icmd.Parameters.Add(p);
                p.SourceColumn = item.DestCol;
            }
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.InsertCommand = icmd;
            da.UpdateCommand = ucmd;
            da.SelectCommand = scmd;
            //da.DeleteCommand = dcmd;

            #endregion

            DataTable tbI = new DataTable(mapping.TableName);
            try {
                da.Fill(tbI);
                FillItemTableOfMidTable(tbI, tb, mapping);
            } catch (Exception ex) {
                PLMEventLog.WriteExceptionLog(ex);
                throw new Exception("在更新中间表过程中，获取中间表" + mapping.TableName + " 数据失败:" + ex.Message + ":" + ex);
            }

            try {
                da.Update(tbI);
            } catch (Exception ex) {
                PLMEventLog.WriteExceptionLog(ex);
                string str2 = "更新中间表" + mapping.TableName + " 失败:" + ex.Message + ":" + ex.InnerException;
                str2 += "\r\n 链接中间表 的链接字符串" + conn.ConnectionString;
                str2 += "\r\n插入语句：" + da.InsertCommand.CommandText;
                str2 += "\r\n更新语句：" + da.UpdateCommand.CommandText;
                str2 += "\r\n获取语句：" + da.SelectCommand.CommandText;
                RecordErr(sOid, str2);
            } finally {
                if (scmd != null) scmd.Dispose();
                if (icmd != null) icmd.Dispose();
                if (ucmd != null) ucmd.Dispose();
                //	if (dcmd != null) dcmd.Dispose();
                da.Dispose();
            }
        }

        private void FillItemTableOfMidTable(DataTable tbOrg, DataTable tbItem, ItemColumnMapping mapping) {
            tbItem.AcceptChanges();
            //DataSet dt = new DataSet();
            //DataTable tbNew = tbItem.Copy();
            //dt.Tables.Add(tbNew);
            //dt.AcceptChanges();
            //// dt.WriteXml(@"C:\Items.xml", XmlWriteMode.WriteSchema);
            //dt.Tables.Remove(tbNew);
            //dt.AcceptChanges();
            tbOrg.AcceptChanges();
            ArrayList lst = new ArrayList();
            for (int i = 0; i < tbItem.Rows.Count; i++) {
                DataRow row = tbItem.Rows[i];
                string id = row[mapping.ID_Item.DestCol].ToString();
                if (!lst.Contains(id))
                    lst.Add(id);
                DataRow rCur = GetRowById(tbOrg, id, mapping.ID_Item.DestCol);
                if (rCur == null) {
                    rCur = tbOrg.NewRow();
                    rCur = RowFillValue(tbOrg, rCur, tbItem, row);
                    tbOrg.Rows.Add(rCur);
                } else {
                    rCur = RowFillValue(tbOrg, rCur, tbItem, row);
                }
            }
            //tbNew = tbOrg.Copy();
            //tbNew.AcceptChanges();
            //dt.Tables.Add(tbNew);
        }

        private DataRow RowFillValue(DataTable tbOrg, DataRow rNew, DataTable tbItem, DataRow row) {
            for (int j = 0; j < tbItem.Columns.Count; j++) {
                string colName = tbItem.Columns[j].ColumnName;
                if (tbOrg.Columns.Contains(colName)) {
                    rNew[colName] = row[colName];
                }
            }
            return rNew;
        }

        private DataRow GetRowById(DataTable tbOrg, string Id, string IdColName) {
            for (int i = 0; i < tbOrg.Rows.Count; i++) {
                DataRow row = tbOrg.Rows[i];
                string id = row[IdColName].ToString();
                if (Id == id)
                    return row;
            }
            return null;
        }

        public static void RecordErr(Guid ssoid, string pId, string Id, string cls, string rel, string ErrInfo,
                                     string WarningInfo) {
            string strs = "  Values(:ssOid,";
            Errcmd.Parameters.Clear();
            Errcmd.Parameters.Add(":ssOid", OracleDbType.Raw).Value = ssoid.ToByteArray();
            Errcmd.CommandText = "insert into PLM_PSM_ERP_EXCEPTION(PLM_SESSION,";
            if (pId != null && !pId.Equals(String.Empty)) {
                Errcmd.Parameters.Add(":pId", OracleDbType.Varchar2).Value = pId;
                Errcmd.CommandText += "PLM_PID,";
                strs += ":pId,";
            }
            if (!string.IsNullOrEmpty(Id)) {
                Errcmd.Parameters.Add(":tId", OracleDbType.Varchar2).Value = Id;
                Errcmd.CommandText += "PLM_ID,";
                strs += ":tId,";
            }
            if (!string.IsNullOrEmpty(cls)) {
                Errcmd.Parameters.Add(":tCls", OracleDbType.Varchar2).Value = cls;
                Errcmd.CommandText += "PLM_CLASS,";
                strs += ":tCls,";
            }

            if (!string.IsNullOrEmpty(rel)) {
                Errcmd.Parameters.Add(":rel", OracleDbType.Varchar2).Value = rel;
                Errcmd.CommandText += "PLM_RELATION ,";
                strs += ":rel,";
            }
            if (!String.IsNullOrEmpty(ErrInfo)) {
                Errcmd.Parameters.Add(":ErrInfo", OracleDbType.Varchar2).Value = ErrInfo;
                Errcmd.CommandText += "PLM_ERROR)";
                strs += ":ErrInfo)";
            } else if (!String.IsNullOrEmpty(WarningInfo)) {
                Errcmd.Parameters.Add(":Warning", OracleDbType.Varchar2).Value = WarningInfo;
                Errcmd.CommandText += "PLM_WARNING)";
                strs += ":Warning)";
            } else {
                return;
            }
            Errcmd.CommandText += strs;
            try {
                Errcmd.ExecuteNonQuery();
            } catch (Exception ex) {
                throw new Exception("向例外表PLM_PSM_ERP_EXCEPTION 添加数据出错。请检查该表是否存在" + ex.Message + "\r\n" + ex);
            }
        }

        public static void RecordErr(Guid ssoid, string ErrInfo) {
            RecordErr(ssoid, null, null, null, null, ErrInfo, null);
        }

        private void GetErrInfo() {
            try {
                StringBuilder strs = new StringBuilder();
                Errcmd.Parameters.Clear();
                Errcmd.Parameters.Add(":ssoid", OracleDbType.Raw).Value = sOid.ToByteArray();
                //更新例外表，补齐数据

                strs.Append("update PLM_PSM_ERP_EXCEPTION t ");
                strs.Append("set t.plm_class = ");
                strs.Append(" (select cls.plm_label  from plm_sys_metaclass cls where  cls.plm_name = t.plm_class ) ");
                strs.Append(" where  t.plm_class is not null ");

                Errcmd.CommandText = strs.ToString();
                Errcmd.ExecuteNonQuery();

                Errcmd.Parameters.Clear();
                Errcmd.Parameters.Add(":ssoid", OracleDbType.Raw).Value = sOid.ToByteArray();

                Errcmd.CommandText =
                    " select PLM_PID,PLM_ID,PLM_RELATION,PLM_CLASS,PLM_ERROR,PLM_WARNING   from PLM_PSM_ERP_EXCEPTION where PLM_SESSION =:ssoid order by PLM_PID,PLM_ID,PLM_RELATION,PLM_ERROR";

                OracleDataReader dr = Errcmd.ExecuteReader();

                //Hashtable hs = new Hashtable();
                //ArrayList lst, lst2;


                while (dr.Read()) {
                    DEErrinfo info = new DEErrinfo();
                    info.pid = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    info.id = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    info.rel = dr.IsDBNull(2) ? "" : dr.GetString(2);
                    info.cls = dr.IsDBNull(3) ? "" : dr.GetString(3);
                    info.errInfo = dr.IsDBNull(4) ? "" : dr.GetString(4);
                    info.warningInfo = dr.IsDBNull(5) ? "" : dr.GetString(5);
                    RecordErrOrWarn(info);
                }
                if (strEr.Length == 0 && strsuc.Length > 0) {
                    strWarn.Insert(0, strsuc.ToString());
                }

                Errcmd.CommandText = " delete from PLM_PSM_ERP_EXCEPTION where PLM_SESSION =:ssoid";
                Errcmd.ExecuteNonQuery();
            } catch (Exception ex) {
                throw new Exception("从例外表PLM_PSM_ERP_EXCEPTION 获取数据数据出错。请检查该表是否存在" + ex.Message + "\r\n" + ex);
            }
        }

        private void RecordErrOrWarn(DEErrinfo info) {
            StringBuilder strtmp;
            string strRd;
            string strErrType;
            if (!String.IsNullOrEmpty(info.errInfo)) {
                strtmp = strEr;
                strRd = info.errInfo;
                strErrType = "错误：";
            } else {
                strRd = info.warningInfo;
                if (info.warningInfo == "BOM成功导出" || info.warningInfo == "物料成功导出") {
                    strErrType = "成功：";
                    strtmp = strsuc;
                } else {
                    strErrType = "警告";
                    ;
                    strtmp = strWarn;
                }
            }
            if (info.pid == string.Empty && info.id != string.Empty) {
                strtmp.Append("\r\n ");
                if (info.cls != "") {
                    strEr.Append("  [" + info.cls + "]:");
                }
                strtmp.Append("<" + info.id + ">");
                strtmp.Append(strRd);
            } else if (info.pid != string.Empty && info.id != string.Empty) {
                strtmp.Append("\r\n ");
                strtmp.Append("父件<" + info.pid + " 子件" + info.id + ">:");
                strtmp.Append(strRd);
            } else if (info.rel != string.Empty) {
                strtmp.Append("\r\n ");
                strtmp.Append("工艺数据：" + info.rel + ":");
                strtmp.Append(strRd);
            } else {
                strtmp.Append("\r\n ");
                strtmp.Append(strErrType);
                strtmp.Append(strRd);
            }
        }

        #endregion
    }
}
