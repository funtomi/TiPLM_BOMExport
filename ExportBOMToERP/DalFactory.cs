using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Thyt.TiPLM.DEL.Product;

namespace ExportBOMToERP {
    public class DalFactory {
        private DalFactory() {
        }

        public static DalFactory Instance;
        static DalFactory() {
            Instance = new DalFactory();
        }
        /// <summary>
        /// 创建数据类型
        /// </summary>
        /// <param name="typeStr"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public BaseDal CreateDal(string typeStr,DEBusinessItem item) {
            if (item==null) {
                return null;
            }
            BaseDal dal = null;
            BusinessType type;
            var hasDefine= this.TryGetBusinessType(typeStr, out type);
            if (!hasDefine) {
                return null;
            }
            switch (type) {
                case BusinessType.UnitGroup:
                    dal = new UnitGroupDal(item);
                    break;
                case BusinessType.Unit:
                    dal = new UnitDal(item);
                    break;
                case BusinessType.Inventory:
                    dal = new InventoryDal(item);
                    break;
                case BusinessType.InventoryClass:
                    dal = new InventoryClassDal(item);
                    break;
                case BusinessType.Resource:
                    dal = new ResourceDal(item);
                    break;
                case BusinessType.Routing:
                    dal = new RoutingDal(item);
                    break;
                case BusinessType.WorkCenter:
                    dal = new WorkCenterDal(item);
                    break;
            }
            return dal;
        }

        /// <summary>
        /// 获取业务类型
        /// </summary>
        /// <param name="type"></param>
        /// <param name="businessType"></param>
        /// <returns></returns>
        public bool TryGetBusinessType(string type,out BusinessType businessType) {
            businessType = BusinessType.Unit;
            if (string.IsNullOrEmpty(type)) {
                return false;
            }
            switch (type.ToLower()) {
                case "unit":
                    businessType = BusinessType.Unit;
                    break;
            }
            return true;
        }

        
    }

    public enum BusinessType{
        UnitGroup,Unit,Inventory,InventoryClass,Resource,Routing,WorkCenter
    }
}
