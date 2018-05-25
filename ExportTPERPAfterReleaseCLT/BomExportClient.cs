using ExportBOMToERP;
using System;
using System.Collections;
using Thyt.TiPLM.Common.Interface.Addin;
using Thyt.TiPLM.DEL.Product;
using Thyt.TiPLM.PLL.Product2;
using Thyt.TiPLM.UIL.Common;
using Thyt.TiPLM.UIL.Product.Common;

namespace ExportTPERPAfterReleaseCLT {
    public class BomExportClient : IAddinClientEntry, IAutoExec {
        public Delegate d_AfterReleased { get; set; }
        public Delegate d_AfterDeleted { get; set; }
        public static DEBusinessItem _item = null;

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

        #region 初始化

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
                        //ExportExecute(item);
                        var bItem = GetDEBusinessItem(item);
                        ExportService srv = new ExportService(bItem);
                        srv.AddOrEditItem();
                    }
                }
            }
        }

        #endregion
        
    }
}
