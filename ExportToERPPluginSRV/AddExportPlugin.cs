using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using Thyt.TiPLM.BRL.Operation;
using Thyt.TiPLM.Common;
using Thyt.TiPLM.Common.Interface.Addin;
using Thyt.TiPLM.DAL.Common;
using Thyt.TiPLM.DAL.Operation;
using Thyt.TiPLM.DEL.Operation;
using Thyt.TiPLM.UIL.Common;
using Thyt.TiPLM.UIL.Common.Operation;

namespace ExportToERPPluginSRV {
    public class AddExportPlugin :  IAutoExec {
        public AddExportPlugin() {

        }
        #region 插入操作内容
        
        private const string OPERATION_ID = "PLM30_ExportToERP";
        private const string OPERATION_LABEL = "导入ERP";
        private const string OPERATION_TOOLTIP = "导入到ERP";
        private const string OPERATION_FILTER = "ExportToERPPlugin.dll,ExportToERPPlugin.ExportToERPFilter";
        private const string OPERATION_EVENTHANDLE = "ExportToERPPlugin.dll,ExportToERPPlugin.ExportHandlerHelper,OnExportToERP";
        #endregion

        public void Activate() {
            //var list = BPMClient.ActivityList;
            //BPMProcessor TheBPMProcessor =new BPMProcessor();
            //DELBPMEntityList list;
            //string theWorkItemState = "Open";
            //TheBPMProcessor.GetWorkItemListByUser(BPMClient.UserID, theWorkItemState, out list);
            //var id = ClientData.Session;
            //DEBusinessItem[] items = PLItem.Agent.get
            //var con =ActionContext2.GetInstance();
            //var iActionContent = con.GetCurrentContext();
            //ExtenderProviderService service = new ExtenderProviderService("bomeditor.toppanel");
            //IAction2[] actionArray = ActionRepository2.GetInstance().FindActionWithService(service);
            //MenuItemBuilder.
            //MenuBuilder.Instance.
            DEOperationItem item = new DEOperationItem() {
                Id = "PLM30_ExportToERP",
            };
            item.Label = "导入ERP";
            item.Tooltip = "导入到ERP";
            item.Filter = "ExportToERPPlugin.dll,ExportToERPPlugin.ExportToERPFilter";
            item.EventHandler = "ExportToERPPlugin.dll,ExportToERPPlugin.ExportHandlerHelper,OnExportToERP";
            item.Option = 0;

        }

        public void Config() {
        }

        public void Init() {
            try {
                PLMEventLog.WriteLog("成功进图", System.Diagnostics.EventLogEntryType.Information);
                var item = OperationConfigHelper.Instance.GetDEOperationItem(OPERATION_ID);
                if (item !=null) {
                PLMEventLog.WriteLog("成功进图", System.Diagnostics.EventLogEntryType.Information);

                    return;
                }
                AddNewOperation();
            } catch (Exception) {
                
                throw;
            }
        }

        private void AddNewOperation() {
            
            DEOperationItem item = new DEOperationItem() {
                Id = "PLM30_ExportToERP",
            };
            item.Label = "导入ERP";
            item.Tooltip = "导入到ERP";
            item.Filter = "ExportToERPPluginCLT.dll,ExportToERPPluginCLT.ExportToERPFilter";
            item.EventHandler = "ExportToERPPluginCLT.dll,ExportToERPPluginCLT.ExportHandlerHelper,OnExportToERP";
            item.Option = 0;
            var dbPara = DBUtil.GetDbParameter(true);
            try {
                //dbPara.Open();
                //new DAOperationDef(dbPara).CreateOperationItem(item);
                new PROperationDef().CreateOperationItem(item, Guid.NewGuid());
            } catch (Exception) {
                
                throw;
            }
        }

        public void UnInit() {
            
        }

        public System.Runtime.Remoting.WellKnownServiceTypeEntry[] RemoteTypes {
            get {
                WellKnownServiceTypeEntry[] entries = new WellKnownServiceTypeEntry[1];
                entries[0] = new WellKnownServiceTypeEntry(
                    typeof(AddExportPlugin),
                    "",
                    WellKnownObjectMode.SingleCall);
                return entries;
            }
        }
    }
}
