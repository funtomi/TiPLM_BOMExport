using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Thyt.TiPLM.DEL.Product;

namespace BOMExportCommon {
    public class ConstERP {
        //public const string RemotingURL = "TiPLM/ERPNewAddin/SVR/BFEntrance.rem";
        public const string RemotingURL = "BFProject/Interface.rems";
    }

    [Serializable]
    public class DEErpExport {
        public string configName;
        public string ExpProeName;
        public string dec;
        public string ExtendSVRDllName;
        public string ExtendCLTDllName;
        public ExpOption op = ExpOption.eDataBase;
        public ArrayList lstExportDataType;
        public bool IsExpOldItemAndBom;
    }

    [Serializable]
    /// <summary>
    /// 数据导出方式
    /// </summary>
    public enum ExpOption : int {
        eDataBase = 0,
        eXls = 1,
        eOther = 2,
        eXml =3
    }
    /// <summary>
    /// Interface 的摘要说明。
    /// </summary>
    public interface Interface {
        /// <summary>
        /// 导出产品结构到ERP系统中
        /// </summary>
        /// <param name="expEvent">导出事件</param>
        /// <param name="deErp">导出程序配置</param>
        /// <param name="strErrInfo">出错信息</param>
        /// <param name="ds">获取数据（可能为空，用于客户端需要展现或保存数据）</param>
        /// <param name="Inobjs">对于需要进行二次开发的企业，需要携带</param>
        /// <param name="outObjects">需要返回在扩展插件中进行二次开发处理的数据</param>
        /// <returns>如果成功，返回true；否则返回false</returns>
        /// <remarks>
        /// </remarks>
        bool ExportPSToERP(DEExportEvent expEvent, DEErpExport deErp, out StringBuilder strErrInfo, out StringBuilder strWarning, out List<DataSet> ds, Object Inobjs, out Object outObjects);

        /// <summary>
        /// 获取可用ERP导出程序
        /// </summary>
        /// <param name="userOid"></param>
        /// <returns></returns>
        ArrayList GetExpProe(Guid userOid);

    }
}

