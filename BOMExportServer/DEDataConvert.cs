using BOMExportCommon;
using Oracle.DataAccess.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Thyt.TiPLM.Common;
using Thyt.TiPLM.DAL.Common;

namespace BOMExportServer {
    public class DEErrinfo {
        public string cls;
        public string errInfo;
        public string id;
        public string pid;
        public string rel;
        public string warningInfo;
    }

    public class DEDataSource {
        private OleDbConnection conn;

        /// <summary>
        /// 数据库类型ORACLE9/SQLSEV/DB2
        /// </summary>
        public string ConnDBType = "ORACLE";

        public string ConnStr = "";

        /// <summary>
        /// 数据服务名或数据库实例名
        /// </summary>
        public string DataBase = "";

        public string PWD = "";

        /// <summary>
        /// 数据库所在服务器IP或机器名
        /// </summary>
        public string Server = "";

        public string User = "";
        /// <summary>
        /// 用于判断导出历史所用数据库
        /// </summary>
        public string GetDBName {
            get {
                if (String.IsNullOrEmpty(DataBase)) return "";
                return DataBase + ";" + User;
            }
        }

        /// <summary>
        /// 获取ERP数据库连接
        /// </summary>
        /// <returns></returns>
        public OleDbConnection GetConnection() {
            if (conn != null) return conn;
            if (ConnStr == "") {
                if (ConnDBType.StartsWith("ORACLE")) {
                    ConnStr = "Provider=Oracle Provider for OLE DB;Data Source=" + DataBase + ";User ID=" + User +
                              ";Password=" + PWD +
                              ";";
                } else if (ConnDBType.StartsWith("SQLSERVER")) {
                    ConnStr =
                        "Provider=Microsoft OLE DB Provider for SQL Server;Persist Security Info=False;Initial Catalog=" +
                        DataBase + ";Data Source=" +
                        Server + ";User Id=" + User + ";Password=" + PWD + ";";
                }
                //				else if (ConnDBType.StartsWith("DB2"))
                //				{
                //				//	Provider=IBMDA400.DataSource.1;Persist Security Info=False;User ID=randy;Data Source=218.69.102.170;Protection Level=None;Extended Properties="";Initial Catalog=S653E20F;Transport Product=Client Access;SSL=DEFAULT;Force Translate=65535;Default Collection=RANDY;Convert Date Time To Char=TRUE;Catalog Library List="";Cursor Sensitivity=3;Use SQL Packages=False;SQL Package Library Name="";SQL Package Name="";Add Statements To SQL Package=True;Unusable SQL Package Action=1;Block Fetch=True;Data Compression=True;Sort Sequence=0;Sort Table Name="";Sort Language ID="";Query Options File Library="";Trace=0;Hex Parser Option=0;Maximum Decimal Precision=31;Maximum Decimal Scale=31;Minimum Divide Scale=0
                //					ConnStr = "DataSource=" + Server + ";userid=" + User + ";password=" + PWD + ";DefaultCollection=" + DataBase +";";
                //				}
            }

            conn = new OleDbConnection(ConnStr);
            return conn;
        }

        public void CloseConnection() {
            if (conn != null && conn.State == ConnectionState.Open)
                conn.Close();
        }
    }


    /// <summary>
    /// 数据库数据类型
    /// </summary>
    public enum DBDataType {
        Unknown = 0,
        Char = 1,
        Varchar = 2,
        DateTime = 3,
        SmallInt = 4,
        Integer = 5,
        BigInt = 6,
        Float = 7,
        Guid = 8,
    }

    /// <summary>
    /// 排序方式
    /// </summary>
    public enum OrderType {
        None = 0,
        Asc = 1,
        Desc = 2,
    }

    /// <summary>
    /// 源属性类型
    /// </summary>
    public enum SourceType {
        /// <summary>
        /// 对象属性
        /// </summary>
        Attribute = 1,

        /// <summary>
        /// 特殊值
        /// </summary>
        SpecialValue = 2,

        /// <summary>
        /// 默认值
        /// </summary>
        DefaultValue = 3,
    }

    public class SourceItem {
        public static string DefaultLinkNum = "-1";
        public Hashtable DataConvertTable = new Hashtable();
        public object DefaultValue;
        public string Format;
        public string PLMCol;
        public string SrcClass;
        public string SrcCol;
        public DBDataType SrcDataType = DBDataType.Unknown;
        public int SrcType = 1; //数据类型（1 PLM中属性，其它 特殊属性）
        public string Value;
        public string LinkNum = DefaultLinkNum; //用于工艺数据导出记录关系链
    }

    /// <summary>
    /// 数据映射项
    /// </summary>
    public class ColumnMappingItem {
        public SourceItem curSrcItem;
        public string DestCol;
        public DBDataType DestDataType = DBDataType.Unknown;
        public Hashtable hsSrcItems = new Hashtable();
        public int i_Size;
        public bool IsAllowConvertErr;
        public bool IsAllowCut = true;
        public bool IsAllowNull = true;
        public bool IsExport = true; //字段是否导出
        public bool IsKey;
        public OrderType ordTy = OrderType.None;
        public string lb;

        public ArrayList lstClassName {
            get {
                ArrayList lst = new ArrayList();
                lst.AddRange(hsSrcItems.Keys);
                return lst;
            }
        }
    }

    public class RelationLink {
        public string LeftClassName;
        public string RelationName;
        public string RightClassName;
        public string SiteInfo;
        public DBDataType SiteInfoDataType = DBDataType.Varchar;

        public RelationLink() {
        }

        public RelationLink(string leftClass, string relName, string rightClass) {
            LeftClassName = leftClass;
            RelationName = relName;
            RightClassName = rightClass;
        }
    }

    /// <summary>
    /// 数据映射
    /// </summary>
    public class ItemColumnMapping {
        /// <summary>
        /// 记录关键列的列映射
        /// </summary>
        public ColumnMappingItem ID_Item;
         
        /// <summary>
        /// 数据是否导出(新增加内容)
        /// </summary>
        public bool IsExport = true;

        /// <summary>
        /// 更新方式true 者 Item，采取更新方式，BOM和CAPP采用先删除后添加中间表中数据
        /// false 者，采用直接添加新纪录的方式更新中间表
        /// </summary>
        public bool IsUpdate = true;

        /// <summary>
        /// 数据映射列
        /// </summary>
        public ArrayList MappingItemList = new ArrayList();

        /// <summary>
        /// 源类起始名称
        /// </summary>
        public string SrcStartName;

        /// <summary>
        /// 中间库表名称
        /// </summary>
        public string TableName;

        /// <summary>
        /// 能够处理的Part及其子类
        /// </summary>
        public ArrayList lstItemClass = new ArrayList();


        /// <summary>
        /// 获取主键列
        /// </summary>
        public ArrayList GetKeyItem {
            get {
                ArrayList lst = new ArrayList();
                for (int i = 0; i < MappingItemList.Count; i++) {
                    ColumnMappingItem item = MappingItemList[i] as ColumnMappingItem;
                    if (!item.IsKey) continue;
                    if (!lst.Contains(item))
                        lst.Add(item);
                }
                return lst;
            }
        }
    }

    /// <summary>
    /// 用于关联的映射建立
    /// </summary>
    public class RelItemColumnMapping : ItemColumnMapping {
        /// <summary>
        /// 存储类名与表简称映射关系(适用于关联数据）(可能为空）
        /// </summary>
        public Hashtable hsClass = new Hashtable();
        /// <summary>
        /// 记录关键列的列映射(专用于工业数据导出）
        /// </summary>
        public ColumnMappingItem PART_ID_Item;

        /// <summary>
        /// 关联列表
        /// </summary>
        public ArrayList lstRelationLink = new ArrayList();

        public Hashtable hsLinks = new Hashtable();
        private Hashtable hsLst = new Hashtable();

        public ArrayList GetLinkList(string key) {
            return hsLst[key] as ArrayList;
        }

        public void SetClassName() {
            if (SrcStartName == null)
                SrcStartName = "PART";
            Hashtable hs;
            hsClass = new Hashtable();
            IDictionaryEnumerator ic = hsLinks.GetEnumerator();
            ic.Reset();

            while (ic.MoveNext()) {
                string key = ic.Key.ToString();
                ArrayList lst = ic.Value as ArrayList;
                hs = new Hashtable();
                int i = 0;
                hs.Add(SrcStartName, i++);

                foreach (RelationLink link in lst) {
                    hs[link.RelationName] = i;
                    hs[link.RightClassName] = i++;
                }
                hsLst[key] = lst;
                hsClass[key] = hs;
            }
        }

        /*
                public void SetClassName()
                {
                    Hashtable hs;
                    hsClass = new Hashtable();
                    foreach (ArrayList lst in lstRelationLink)
                    {
                        hs = new Hashtable();
                     int   i = 0;
                        if (SrcStartName == null)
                            SrcStartName = "PART";
                        hs.Add(SrcStartName, i++);
                        if (lst.Count != 0)
                        {
                            foreach (RelationLink link in lst)
                            {
                                hs[link.RelationName] = i;
                                hs[link.RightClassName] = i++;
                            }
                        }
                        hsClass.Add(lst, hs);
                    }
                }


                /// <summary>
                /// 获取列的别名
                /// </summary>
                /// <param name="item"></param>
                /// <returns></returns>
                public string GetPLMColName(ColumnMappingItem item, Hashtable hs)
                {
                    string strPLMColName = "";
                    int i;
                    if (hs.Count == 0)
                    {
                        return "";
                    }
                    ArrayList lst = new ArrayList();
                    lst.AddRange(item.hsSrcItems.Keys);
                    for (int j = 0; j < lst.Count; j++)
                    {
                        string cls = lst[j].ToString();
                        if (hs[cls] == null)
                            continue;
                        i = Convert.ToInt32(hs[cls]);
                        item.curSrcItem = item.hsSrcItems[cls] as SourceItem;
                        if (item.curSrcItem == null) return strPLMColName;
                        if (item.curSrcItem.SrcCol.IndexOf("M.") != -1
                            || item.curSrcItem.SrcCol.IndexOf("I.") != -1
                            || item.curSrcItem.SrcCol.IndexOf("R.") != -1)
                        {
                            strPLMColName = item.curSrcItem.SrcCol.Substring(0, 1) + i + item.curSrcItem.SrcCol.Substring(1);
                            item.curSrcItem.PLMCol = item.curSrcItem.SrcCol.Substring(0, 1) + i + "_" +
                                                     item.curSrcItem.SrcCol.Substring(2);
                        }
                        break;
                    }
                    return strPLMColName;
                }

        */
        /// <summary>
        /// 获取列的别名
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public string GetPLMColName(ColumnMappingItem item, string key) {
            string strPlmColName = "";
            int i;
            Hashtable hs = hsClass[key] as Hashtable;
            if (hs == null) return "";
            IDictionaryEnumerator ic = item.hsSrcItems.GetEnumerator();
            ic.Reset();
            while (ic.MoveNext()) {
                string cls = ic.Key.ToString();
                SourceItem srcItem = ic.Value as SourceItem;

                if (srcItem.LinkNum == key || srcItem.LinkNum == SourceItem.DefaultLinkNum) {
                    item.curSrcItem = srcItem;
                    i = Convert.ToInt32(hs[cls]);
                    if (item.curSrcItem.SrcCol.IndexOf("M.") != -1
                    || item.curSrcItem.SrcCol.IndexOf("I.") != -1
                    || item.curSrcItem.SrcCol.IndexOf("R.") != -1) {
                        strPlmColName = item.curSrcItem.SrcCol.Substring(0, 1) + i + item.curSrcItem.SrcCol.Substring(1);
                        item.curSrcItem.PLMCol = item.curSrcItem.SrcCol.Substring(0, 1) + i + "_" +
                                                 item.curSrcItem.SrcCol.Substring(2);
                    }
                    break;
                }
            }
            return strPlmColName;
        }
    }


    /// <summary>
    /// DEDataConvert 的摘要说明。
    /// </summary>
    public class DEDataConvert : DABase {
        /// <summary>
        /// 有权限执行该插件的用户标识列表
        /// </summary>
        public ArrayList AuthorizedUsers;

        ///// <summary>
        ///// BOM字段映射表
        ///// </summary>
        //public ItemColumnMapping BomTable;


        ///// <summary>
        ///// BOMHead字段映射表
        ///// </summary>
        //public ItemColumnMapping BomHeadTable = null;

        public ArrayList lstBomTables = new ArrayList();

        /// <summary>
        /// 新的预处理存储过程
        /// </summary>
        public string DP_BEFORE_NAME = "";


        /// <summary>
        /// 数据导入到中间表以后对，进入中间表的数据进行后处理的存储过程名
        /// </summary>
        public string DP_END_NAME = "";

        /// <summary>
        /// 数据预处理存储过程名
        /// </summary>
        public string DP_FUN_NAME = "";

        public DEDataSource DS = new DEDataSource();

        /// <summary>
        /// 是否在前处理扩展接口之后执行预处理存储过程（false 则在扩展程序之前执行后处理存储过程）
        /// </summary>
        public bool IsExeDPAfterInterface;

        /// <summary>
        /// 是否最后执行存储过程（false 则在扩展程序之前执行后处理存储过程）
        /// </summary>
        public bool IsExeDPAllEnd = true;

        // public ItemColumnMapping ItemTable;

        public ArrayList lstItemTables = new ArrayList();

        /// <summary>
        /// 存储零件关系的表、主要为了处理零件与材料之类的数据关系
        /// </summary>
        public ArrayList lstPartLink;

        public bool IsExportItem = true;
        public bool IsExportBom = true;
        //public bool isExportOther = true;

        /// <summary>
        /// 导出关联的数据表（工艺）
        /// </summary>
        public ArrayList lstRouteTables = new ArrayList();

        public DEDataConvert(DBParameter dbParam)
            : base(dbParam) {
        }

        /// <summary>
        /// 从配置文件KLERPIntegrateSVR.config中获取配置信息。
        /// </summary>
        /// <param name="deErp">导出配置文件选择与导出类型设置</param>
        /// <param name="ERPConfig">配置文件名称</param>
        public void GetConfiguration(string ERPConfig, DEErpExport deErp) {
            try {
                XmlDocument doc = new XmlDocument();
                doc.Load(ERPConfig);
                ArrayList lstAllPartType = GetSubClass(ConstProduct.CLASS_PART);

                foreach (XmlNode node in doc.DocumentElement.ChildNodes) {
                    #region 数据源

                    if (node.Name == "DataSource") {
                        foreach (XmlNode nd in node.ChildNodes) {
                            if (nd.Name == "DBType")
                                DS.ConnDBType = nd.InnerText.ToUpper();
                            if (nd.Name == "Server")
                                DS.Server = nd.InnerText;
                            if (nd.Name == "User")
                                DS.User = nd.InnerText;
                            if (nd.Name == "Database")
                                DS.DataBase = nd.InnerText;
                            if (nd.Name == "Password")
                                DS.PWD = nd.InnerText;
                        }
                    }

                    #endregion

                    #region DataMapping 建立数据映射

                    if (node.Name == "DataMapping") {
                        foreach (XmlNode nd in node.ChildNodes) {
                            #region Item 表的映射

                            if (nd.Name == "ItemColumnMapping") {
                                ItemColumnMapping ItemTable = new ItemColumnMapping();
                                if (!deErp.lstExportDataType.Contains("导出物料"))
                                    IsExportItem = false;
                                ItemTable.TableName = nd.Attributes["ItemTableName"].Value;
                                if (nd.Attributes["UpdateType"] != null)
                                    ItemTable.IsUpdate = nd.Attributes["UpdateType"].Value.ToUpper() == "INSERT"
                                                             ? false
                                                             : true;
                                foreach (XmlNode nd1 in nd.ChildNodes) {
                                    if (nd1.Name == "Column") {
                                        ColumnMappingItem item = new ColumnMappingItem();
                                        string strFormat = "";
                                        if (nd1.Attributes["Format"] != null)
                                            strFormat = nd1.Attributes["Format"].Value;
                                        item.DestCol = nd1.Attributes["DestName"].Value;
                                        item.DestDataType = ConvertXML(nd1.Attributes["DataType"].Value);
                                        if (nd1.Attributes["AllowNull"] != null)
                                            item.IsAllowNull = (nd1.Attributes["AllowNull"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["AllowCut"] != null)
                                            item.IsAllowCut = (nd1.Attributes["AllowCut"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["AllowConvertErr"] != null)
                                            item.IsAllowConvertErr = (nd1.Attributes["AllowConvertErr"].Value == "Y")
                                                                         ? true
                                                                         : false;
                                        //if (!item.IsAllowNull )
                                        //    item.IsAllowConvertErr = false;
                                        if (nd1.Attributes["Size"] != null)
                                            item.i_Size = Convert.ToInt32(nd1.Attributes["Size"].Value);
                                        if (nd1.Attributes["IsKey"] != null)
                                            item.IsKey = (nd1.Attributes["IsKey"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["IsID"] != null)
                                            ItemTable.ID_Item = nd1.Attributes["IsID"].Value == "Y" ? item : null;
                                        if (nd1.Attributes["IsExport"] != null)
                                            item.IsExport = (nd1.Attributes["IsExport"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["Order"] != null) {
                                            if (nd1.Attributes["Order"].Value.ToUpper() == "ASC")
                                                item.ordTy = OrderType.Asc;
                                            else if (nd1.Attributes["Order"].Value.ToUpper() == "DESC")
                                                item.ordTy = OrderType.Desc;
                                        }
                                        if (nd1.Attributes["Lable"] != null) {
                                            item.lb = nd1.Attributes["Lable"].Value;
                                        }
                                        if (nd1.ChildNodes.Count > 0) {
                                            foreach (XmlNode nd2 in nd1.ChildNodes) {
                                                if (nd2.Name == "SourceColumn") {
                                                    SourceItem srcItem = new SourceItem();
                                                    srcItem.SrcClass = nd2.Attributes["ClassName"].Value.ToUpper();
                                                    if (srcItem.SrcClass.Trim() != "" &&
                                                        !ItemTable.lstItemClass.Contains(srcItem.SrcClass))
                                                        ItemTable.lstItemClass.Add(srcItem.SrcClass);
                                                    srcItem.SrcCol = nd2.Attributes["ColName"].Value.ToUpper();
                                                    if (srcItem.SrcCol == "M.PLM_M_ID")
                                                        item.IsAllowConvertErr = false;
                                                    if (srcItem.SrcClass == "")
                                                        srcItem.SrcType = 0;
                                                    if (nd2.Attributes["DataType"] != null &&
                                                        nd2.Attributes["DataType"].Value.Trim() != "")
                                                        srcItem.SrcDataType =
                                                            ConvertXML(nd2.Attributes["DataType"].Value);
                                                    else
                                                        srcItem.SrcDataType = item.DestDataType;
                                                    if (nd2.Attributes["Value"] != null)
                                                        srcItem.Value = nd2.Attributes["Value"].Value;
                                                    if (String.IsNullOrEmpty(srcItem.Value) && !item.IsAllowNull) {
                                                        if (!item.IsAllowConvertErr)
                                                            item.IsAllowConvertErr = false;
                                                    }
                                                    if (strFormat != "")
                                                        srcItem.Format = strFormat;
                                                    else if (nd2.Attributes["Format"] != null)
                                                        srcItem.Format = nd2.Attributes["Format"].Value;
                                                    item.hsSrcItems[srcItem.SrcClass] = srcItem;
                                                    if (ItemTable.ID_Item == null && srcItem.SrcCol == "M.PLM_M_ID")
                                                        ItemTable.ID_Item = item;
                                                }
                                            }
                                        }
                                        ItemTable.MappingItemList.Add(item);
                                    }
                                }
                                lstItemTables.Add(ItemTable);
                            }

                            #endregion

                            #region BOM

                            if (nd.Name == "BomColumnMapping") {
                                ItemColumnMapping BomTable = new ItemColumnMapping();
                                if (!deErp.lstExportDataType.Contains("导出BOM"))
                                    IsExportBom = false;
                                BomTable.TableName = nd.Attributes["BomTableName"].Value;
                                if (nd.Attributes["UpdateType"] != null)
                                    BomTable.IsUpdate = nd.Attributes["UpdateType"].Value.ToUpper() == "INSERT"
                                                            ? false
                                                            : true;
                                foreach (XmlNode nd1 in nd.ChildNodes) {
                                    if (nd1.Name == "Column") {
                                        ColumnMappingItem item = new ColumnMappingItem();

                                        item.curSrcItem = new SourceItem();
                                        item.DestCol = nd1.Attributes["DestName"].Value;
                                        item.DestDataType = ConvertXML(nd1.Attributes["DataType"].Value.Trim().ToUpper());
                                        if (nd1.Attributes["PLMDataType"] != null)
                                            item.curSrcItem.SrcDataType =
                                                ConvertXML(nd1.Attributes["PLMDataType"].Value.Trim().ToUpper());
                                        else {
                                            item.curSrcItem.SrcDataType = item.DestDataType;
                                        }
                                        item.curSrcItem.SrcType = 0;
                                        if (nd1.Attributes["AllowNull"] != null)
                                            item.IsAllowNull = (nd1.Attributes["AllowNull"].Value.Trim().ToUpper() == "Y") ? true : false;
                                        if (nd1.Attributes["Size"] != null)
                                            item.i_Size = Convert.ToInt32(nd1.Attributes["Size"].Value);
                                        if (nd1.Attributes["Value"] != null)
                                            item.curSrcItem.Value = nd1.Attributes["Value"].Value.Trim().ToUpper();
                                        if (nd1.Attributes["DefaultValue"] != null)
                                            item.curSrcItem.DefaultValue =
                                                nd1.Attributes["DefaultValue"].Value.Trim().ToUpper();
                                        if (nd1.Attributes["Format"] != null)
                                            item.curSrcItem.Format = nd1.Attributes["Format"].Value.Trim();
                                        if (nd1.Attributes["IsKey"] != null)
                                            item.IsKey = (nd1.Attributes["IsKey"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["IsID"] != null)
                                            BomTable.ID_Item = nd1.Attributes["IsID"].Value == "Y" ? item : null;
                                        if (nd1.Attributes["AllowCut"] != null)
                                            item.IsAllowCut = (nd1.Attributes["AllowCut"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["IsExport"] != null)
                                            item.IsExport = (nd1.Attributes["IsExport"].Value == "Y") ? true : false;
                                        if (nd1.Attributes["Order"] != null) {
                                            if (nd1.Attributes["Order"].Value.ToUpper() == "ASC")
                                                item.ordTy = OrderType.Asc;
                                            else if (nd1.Attributes["Order"].Value.ToUpper() == "DESC")
                                                item.ordTy = OrderType.Desc;
                                        }
                                        if (nd1.Attributes["AllowConvertErr"] != null)
                                            item.IsAllowConvertErr = (nd1.Attributes["AllowConvertErr"].Value == "Y")
                                                                         ? true
                                                                         : false;
                                        if (item.curSrcItem.Value == "PID" || item.curSrcItem.Value == "CID" ||
                                            item.curSrcItem.Value == "NUMBER")
                                            item.IsAllowConvertErr = false;
                                        if (!item.IsAllowNull && item.curSrcItem.DefaultValue != null &&
                                            String.IsNullOrEmpty(item.curSrcItem.DefaultValue.ToString())) {
                                            item.IsAllowConvertErr = false;
                                        }


                                        BomTable.MappingItemList.Add(item);
                                        if (item.curSrcItem.Value != null && BomTable.ID_Item == null) {
                                            if (item.curSrcItem.Value == "PID" || item.curSrcItem.Value == "P.M.PLM_M_OID")
                                                BomTable.ID_Item = item;
                                        }
                                        if (nd1.Attributes["Lable"] != null) {
                                            item.lb = nd1.Attributes["Lable"].Value;
                                        }
                                    }
                                }
                                lstBomTables.Add(BomTable);
                            }

                            #endregion

                            #region Routing

                            if (nd.Name == "RoutingMapping") {
                                string lb;
                                if (nd.Attributes["Lable"] != null) {
                                    XmlAttribute attribute = nd.Attributes["Lable"];
                                    lb = attribute.Value;
                                } else lb = "导出工艺";
                                RelItemColumnMapping RoutTable = new RelItemColumnMapping();
                                if (!deErp.lstExportDataType.Contains(lb))
                                    RoutTable.IsExport = false;

                                foreach (XmlNode nd1 in nd.ChildNodes) {
                                    string linkNum; // = k.ToString();
                                    int k = 0;
                                    if (nd1.Name == "RelationLinks") {
                                        foreach (XmlNode nd2 in nd1.ChildNodes) {
                                            ArrayList lst = new ArrayList();
                                            if (nd2.Name == "RelationLink") {
                                                if (nd2.Attributes["LinkNum"] != null) {
                                                    linkNum = nd2.Attributes["LinkNum"].Value;
                                                } else {
                                                    linkNum = k.ToString();
                                                    k++;
                                                }
                                                foreach (XmlNode nd3 in nd2.ChildNodes) {
                                                    if (nd3.Name == "Relation") {
                                                        RelationLink link = new RelationLink();
                                                        link.RelationName =
                                                            nd3.Attributes["RelationName"].Value.ToUpper();
                                                        link.RightClassName =
                                                            nd3.Attributes["RightClassName"].Value.ToUpper();
                                                        if (nd3.Attributes["SiteInfo"] != null)
                                                            link.SiteInfo = nd3.Attributes["SiteInfo"].Value.ToUpper();
                                                        if (nd3.Attributes["SiteInfoDataType"] != null)
                                                            link.SiteInfoDataType =
                                                                ConvertXML(nd2.Attributes["DataType"].Value);
                                                        lst.Add(link);
                                                    }
                                                }
                                                RoutTable.hsLinks[linkNum] = new ArrayList(lst);
                                            }
                                            RoutTable.lstRelationLink.Add(lst);
                                        }
                                    }
                                    if (nd1.Name == "RoutingColumnMapping") {
                                        RoutTable.TableName = nd1.Attributes["RoutingTableName"].Value;
                                        if (nd.Attributes["UpdateType"] != null)
                                            RoutTable.IsUpdate = nd.Attributes["UpdateType"].Value.ToUpper() == "INSERT"
                                                                     ? false
                                                                     : true;
                                        RoutTable.SrcStartName = "PART";
                                        foreach (XmlNode nd2 in nd1.ChildNodes) {
                                            if (nd2.Name == "Column") {
                                                ColumnMappingItem item = new ColumnMappingItem();
                                                string strFormat = "";
                                                if (nd2.Attributes["Format"] != null)
                                                    strFormat = nd2.Attributes["Format"].Value;
                                                item.DestCol = nd2.Attributes["DestName"].Value;
                                                item.DestDataType = ConvertXML(nd2.Attributes["DataType"].Value);
                                                if (nd2.Attributes["AllowNull"] != null)
                                                    item.IsAllowNull = (nd2.Attributes["AllowNull"].Value == "Y")
                                                                           ? true
                                                                           : false;
                                                if (nd2.Attributes["AllowCut"] != null)
                                                    item.IsAllowCut = (nd2.Attributes["AllowCut"].Value == "Y")
                                                                          ? true
                                                                          : false;
                                                if (nd2.Attributes["AllowConvertErr"] != null)
                                                    item.IsAllowConvertErr = (nd2.Attributes["AllowConvertErr"].Value ==
                                                                              "Y")
                                                                                 ? true
                                                                                 : false;

                                                if (nd2.Attributes["Size"] != null)
                                                    item.i_Size = Convert.ToInt32(nd2.Attributes["Size"].Value);
                                                if (nd2.Attributes["IsKey"] != null)
                                                    item.IsKey = (nd2.Attributes["IsKey"].Value == "Y") ? true : false;
                                                if (nd2.Attributes["IsID"] != null)
                                                    RoutTable.ID_Item = nd2.Attributes["IsID"].Value == "Y" ? item : null;
                                                if (nd2.Attributes["IsExport"] != null)
                                                    item.IsExport = (nd2.Attributes["IsExport"].Value == "Y")
                                                                        ? true
                                                                        : false;
                                                if (nd2.Attributes["Order"] != null) {
                                                    if (nd2.Attributes["Order"].Value.ToUpper() == "ASC")
                                                        item.ordTy = OrderType.Asc;
                                                    else if (nd2.Attributes["Order"].Value.ToUpper() == "DESC")
                                                        item.ordTy = OrderType.Desc;
                                                }
                                                if (nd2.Attributes["Lable"] != null) {
                                                    item.lb = nd2.Attributes["Lable"].Value;
                                                }
                                                if (nd2.ChildNodes.Count > 0) {
                                                    foreach (XmlNode nd3 in nd2.ChildNodes) {
                                                        if (nd3.Name == "SourceColumn") {
                                                            SourceItem srcItem = new SourceItem();
                                                            srcItem.SrcClass =
                                                                nd3.Attributes["ClassName"].Value.ToUpper();

                                                            srcItem.LinkNum = nd3.Attributes["LinkNum"] == null ? SourceItem.DefaultLinkNum : nd3.Attributes["LinkNum"].Value;

                                                            srcItem.SrcCol = nd3.Attributes["ColName"].Value.ToUpper();
                                                            if (srcItem.SrcClass == "")
                                                                srcItem.SrcType = 0;
                                                            if (nd3.Attributes["DataType"] != null)
                                                                srcItem.SrcDataType =
                                                                    ConvertXML(nd3.Attributes["DataType"].Value);
                                                            else
                                                                srcItem.SrcDataType = item.DestDataType;
                                                            if (nd3.Attributes["Value"] != null)
                                                                srcItem.Value = nd3.Attributes["Value"].Value;
                                                            if (String.IsNullOrEmpty(srcItem.Value) && !item.IsAllowNull) {
                                                                if (item.IsAllowConvertErr)
                                                                    item.IsAllowConvertErr = false;
                                                            }
                                                            if (strFormat != "")
                                                                srcItem.Format = strFormat;
                                                            else if (nd3.Attributes["Format"] != null)
                                                                srcItem.Format = nd3.Attributes["Format"].Value;
                                                            item.hsSrcItems[srcItem.SrcClass] = srcItem;
                                                            if (lstAllPartType.Contains(srcItem.SrcClass) && srcItem.SrcCol == "M.PLM_M_ID") {
                                                                //RoutTable.SrcStartName = srcItem.SrcClass;
                                                                RoutTable.PART_ID_Item = item;
                                                                item.IsAllowConvertErr = false;
                                                                if (RoutTable.ID_Item == null) {
                                                                    RoutTable.ID_Item = item;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                RoutTable.MappingItemList.Add(item);
                                            }
                                        }
                                    }
                                    RoutTable.SetClassName();
                                }
                                if (RoutTable.PART_ID_Item == null)
                                    RoutTable.PART_ID_Item = RoutTable.ID_Item;

                                lstRouteTables.Add(RoutTable);
                            }

                            #endregion
                        }
                    } // DataMapping				

                    #endregion

                    #region 数据替换规则映射

                    if (node.Name == "DataConvertItems") {
                        foreach (XmlNode nd in node.ChildNodes) {
                            if (nd.Name == "DataConvertItem") {
                                if (nd.Attributes["SrcClassName"] != null && nd.Attributes["ColName"] != null) {
                                    string className = ((XmlElement)nd).GetAttribute("SrcClassName").ToUpper();
                                    string colName = ((XmlElement)nd).GetAttribute("ColName").ToUpper();

                                    for (int i = 0; i < lstItemTables.Count; i++) {
                                        ItemColumnMapping ItemTable = lstItemTables[i] as ItemColumnMapping;
                                        SetConvertDataMapping(nd, className, colName, ItemTable);
                                    }
                                    for (int i = 0; i < lstRouteTables.Count; i++) {
                                        ItemColumnMapping RoutTable = lstRouteTables[i] as ItemColumnMapping;
                                        SetConvertDataMapping2(nd, className, colName, RoutTable);
                                    }
                                    for (int i = 0; i < lstBomTables.Count; i++) {
                                        ItemColumnMapping ItemTable = lstBomTables[i] as ItemColumnMapping;
                                        SetConvertDataMapping3(nd, colName, ItemTable);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region 数据处理规则


                    if (node.Name == "BeforeEXPDataProcessing") {
                        DP_BEFORE_NAME = node.InnerText.Trim();
                        if (node.Attributes["IsExeDPAfterInterface"] != null) {
                            IsExeDPAfterInterface =
                                ((XmlElement)node).GetAttribute("IsExeDPAfterInterface").Trim().ToUpper() == "Y"
                                    ? true
                                    : false;
                        }
                    }
                    if (node.Name == "DataProcessing") {
                        DP_BEFORE_NAME = node.InnerText.Trim();
                        IsExeDPAfterInterface = false;
                    }
                    if (node.Name == "AfterEXPDataProcessing") {
                        DP_END_NAME = node.InnerText.Trim();
                        if (node.Attributes["IsExeDPAllEnd"] != null) {
                            IsExeDPAllEnd = ((XmlElement)node).GetAttribute("IsExeDPAllEnd").Trim().ToUpper() == "Y"
                                                ? true
                                                : false;
                        }
                    }

                    #endregion
                }
            } catch (Exception ex) {
                PLMEventLog.WriteExceptionLog(ex);
                throw new Exception("读取配置文件信息错误：" + ex.Message);
            }
        }

        private void SetConvertDataMapping3(XmlNode nd, string colName, ItemColumnMapping itemTable) {
            SourceItem srcItem;
            foreach (ColumnMappingItem item in itemTable.MappingItemList) {
                srcItem = item.curSrcItem;
                string v = srcItem.Value;
                string v1 = v;
                if (string.IsNullOrEmpty(v)) continue;
                if (v.StartsWith("P.") || v.StartsWith("C."))
                    v1 = v.Substring(2);
                if (colName == v || colName == v1) {
                    foreach (XmlNode nd1 in nd.ChildNodes) {
                        srcItem.DataConvertTable[nd1.Attributes["Source"].Value] =
                            nd1.Attributes["Destination"].Value;
                    }
                }
            }
        }

        /// <summary>
        /// 设置数据转换的映射
        /// </summary>
        /// <param name="nd"></param>
        /// <param name="className"></param>
        /// <param name="colName"></param>
        /// <param name="mapping"></param>
        private void SetConvertDataMapping(XmlNode nd, string className, string colName, ItemColumnMapping mapping) {
            ArrayList lstClass = GetSubClass(className);
            SourceItem srcItem;
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                foreach (string cls in item.lstClassName) {
                    srcItem = item.hsSrcItems[cls] as SourceItem;
                    if (!lstClass.Contains(cls) || srcItem == null) continue;
                    if (colName == srcItem.SrcCol) {
                        foreach (XmlNode nd1 in nd.ChildNodes) {
                            srcItem.DataConvertTable[nd1.Attributes["Source"].Value] =
                                nd1.Attributes["Destination"].Value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 设置数据转换的映射
        /// </summary>
        /// <param name="nd"></param>
        /// <param name="className"></param>
        /// <param name="colName"></param>
        /// <param name="mapping"></param>
        private void SetConvertDataMapping2(XmlNode nd, string className, string colName, ItemColumnMapping mapping) {
            // ArrayList lstClass = GetSubClass(className);
            SourceItem srcItem;
            foreach (ColumnMappingItem item in mapping.MappingItemList) {
                foreach (string cls in item.lstClassName) {
                    srcItem = item.hsSrcItems[cls] as SourceItem;
                    if (srcItem == null) continue;
                    if (colName == srcItem.SrcCol) {
                        foreach (XmlNode nd1 in nd.ChildNodes) {
                            srcItem.DataConvertTable[nd1.Attributes["Source"].Value] =
                                nd1.Attributes["Destination"].Value;
                        }
                    }
                }
            }
        }


        public DBDataType ConvertXML(string str) {
            switch (str) {
                case "1":
                    return DBDataType.Char;
                case "2":
                    return DBDataType.Varchar;
                case "3":
                    return DBDataType.DateTime;
                case "4":
                    return DBDataType.SmallInt;
                case "5":
                    return DBDataType.Integer;
                case "6":
                    return DBDataType.BigInt;
                case "7":
                    return DBDataType.Float;
                case "8":
                    return DBDataType.Guid;
            }
            return DBDataType.Unknown;
        }

        /// <summary>
        /// 转换自定义数据类型为系统类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public Type ConvertSystemType(DBDataType type) {
            switch (type) {
                case DBDataType.Char:
                case DBDataType.Varchar:
                    return typeof(String);
                case DBDataType.SmallInt:
                    return typeof(Int16);
                case DBDataType.Integer:
                    return typeof(Int32);
                case DBDataType.BigInt:
                    return typeof(Int64);
                case DBDataType.Float:
                    return typeof(Decimal);
                case DBDataType.DateTime:
                    return typeof(DateTime);
                case DBDataType.Guid:
                    return typeof(Guid);
                default:
                    return typeof(String);
            }
        }

        /// <summary>
        /// 转换自定义数据类型为系统类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public OleDbType ConvertToOleDBType(DBDataType type) {
            switch (type) {
                case DBDataType.Char:
                    return OleDbType.Char;
                case DBDataType.Varchar:
                    return OleDbType.VarChar;
                case DBDataType.SmallInt:
                    return OleDbType.SmallInt;
                case DBDataType.Integer:
                    return OleDbType.Integer;
                case DBDataType.BigInt:
                    return OleDbType.BigInt;
                case DBDataType.Float:
                    return OleDbType.Numeric;
                case DBDataType.DateTime:
                    return OleDbType.DBTimeStamp;
                case DBDataType.Guid:
                    return OleDbType.Guid;
                default:
                    return OleDbType.VarChar;
            }
        }

        /// <summary>
        /// 转换自定义数据类型为系统类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public OdbcType ConvertToOdbcType(DBDataType type) {
            switch (type) {
                case DBDataType.Char:
                    return OdbcType.Char;
                case DBDataType.Varchar:
                    return OdbcType.VarChar;
                case DBDataType.SmallInt:
                    return OdbcType.SmallInt;
                case DBDataType.Integer:
                    return OdbcType.Int;
                case DBDataType.BigInt:
                    return OdbcType.BigInt;
                case DBDataType.Float:
                    return OdbcType.Numeric;
                case DBDataType.DateTime:
                    return OdbcType.DateTime;
                case DBDataType.Guid:
                    return OdbcType.UniqueIdentifier;
                default:
                    return OdbcType.VarChar;
            }
        }

        /*		/// <summary>
                /// 转换自定义数据类型为系统类型
                /// </summary>
                /// <param name="type"></param>
                /// <returns></returns>
                public iDB2DbType ConvertToDB2Type(DBDataType type)
                {
                    switch (type)
                    {
                        case DBDataType.Char:
                            return iDB2DbType.iDB2Char;
                        case DBDataType.Varchar:
                            return iDB2DbType.iDB2VarChar;
                        case DBDataType.SmallInt:
                            return iDB2DbType.iDB2SmallInt;
                        case DBDataType.Integer:
                            return iDB2DbType.iDB2Integer;
                        case DBDataType.BigInt:
                            return iDB2DbType.iDB2BigInt;
                        case DBDataType.Float:
                            return iDB2DbType.iDB2Double;
                        case DBDataType.DateTime:
                            return iDB2DbType.iDB2Date;
                        default:
                            return iDB2DbType.iDB2VarChar;
                    }
                }
                */

        /// <summary>
        /// 转换自定义数据类型为系统类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public SqlDbType ConvertToSqlType(DBDataType type) {
            switch (type) {
                case DBDataType.Char:
                    return SqlDbType.Char;
                case DBDataType.Varchar:
                    return SqlDbType.VarChar;
                case DBDataType.SmallInt:
                    return SqlDbType.SmallInt;
                case DBDataType.Integer:
                    return SqlDbType.Int;
                case DBDataType.BigInt:
                    return SqlDbType.BigInt;
                case DBDataType.Float:
                    return SqlDbType.Float;
                case DBDataType.DateTime:
                    return SqlDbType.DateTime;
                case DBDataType.Guid:
                    return SqlDbType.UniqueIdentifier;
                default:
                    return SqlDbType.VarChar;
            }
        }

        /// <summary>
        /// 转换自定义数据类型为系统类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public OracleDbType ConvertToOracleType(DBDataType type) {
            switch (type) {
                case DBDataType.Char:
                    return OracleDbType.Char;
                case DBDataType.Varchar:
                    return OracleDbType.Varchar2;
                case DBDataType.SmallInt:
                    return OracleDbType.Int16;
                case DBDataType.Integer:
                    return OracleDbType.Int32;
                case DBDataType.BigInt:
                    return OracleDbType.Int64;
                case DBDataType.Float:
                    return OracleDbType.Double;
                case DBDataType.DateTime:
                    return OracleDbType.Date;
                case DBDataType.Guid:
                    return OracleDbType.Raw;
                default:
                    return OracleDbType.Varchar2;
            }
        }


        public object ConvertOracle(DBDataType type, OracleDataReader dr, int i) {
            switch (type) {
                case DBDataType.Char:
                    return Convert.ToChar(dr.GetOracleString(i).ToString()[0]);
                case DBDataType.Varchar:
                    return dr.GetOracleString(i).ToString();
                case DBDataType.SmallInt:
                    return dr.GetInt32(i);
                case DBDataType.Integer:
                    return dr.GetInt32(i);
                case DBDataType.BigInt:
                    return dr.GetOracleDecimal(i).ToInt64();
                case DBDataType.Float:
                    return dr.GetOracleDecimal(i).ToDouble();
                case DBDataType.DateTime:
                    return dr.GetOracleDate(i).Value;
                case DBDataType.Guid:
                    return new Guid(dr.GetOracleBinary(i).Value);
                default:
                    return dr.GetOracleValue(i);
            }
        }

        /// <summary>
        /// 将对象转化为指定类型数据
        /// </summary>
        /// <param name="type">类型</param>
        /// <param name="obj">对象</param>
        /// <returns>数据对象</returns>
        public object ConvertToSystem(DBDataType type, object obj) {
            switch (type) {
                case DBDataType.Char:
                    return Convert.ToChar(obj);
                case DBDataType.Varchar:
                    return Convert.ToString(obj);
                case DBDataType.SmallInt:
                    return Convert.ToInt16(obj);
                case DBDataType.Integer:
                    return Convert.ToInt32(obj);
                case DBDataType.BigInt:
                    return Convert.ToInt64(obj);
                case DBDataType.Float:
                    return Convert.ToDecimal(obj);
                case DBDataType.DateTime:
                    return Convert.ToDateTime(obj);
                case DBDataType.Guid:
                    if (obj is Guid) return ((Guid)obj).ToByteArray();
                    return obj;
                default:
                    return obj;
            }
        }

        #region 获取要进行数据转换的类型

        public ArrayList GetSubClass(string className) {
            ArrayList lstClass = new ArrayList();
            OracleCommand cmd = new OracleCommand();
            OracleDataReader dr = null;
            cmd.Connection = (OracleConnection)dbParam.Connection;
            cmd.CommandText =
                "SELECT PLM_NAME FROM PLM_SYS_METACLASS CONNECT BY PRIOR PLM_OID=PLM_PARENT START WITH PLM_NAME='" +
                className + "'";
            try {
                dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    string str = dr.GetString(0);
                    lstClass.Add(str);
                }
            } finally {
                if (dr != null)
                    dr.Close();
            }
            return lstClass;
        }

        #endregion
    }
}
