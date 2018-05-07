using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Thyt.TiPLM.Common;

namespace BOMExportClient {
    public class ExportExcel {
        private DataSet ds;
        private Guid eventOid;
        private Process ps;
        private Guid userOid;
        public ExportExcel(DataSet ds, Guid eventOid, Guid userOid) {
            //
            // TODO: 在此处添加构造函数逻辑
            //
            this.ds = ds;
            this.eventOid = eventOid;
            this.userOid = userOid;
        }

        public static void ExportToExcel(DataSet ds, Guid eventOid, Guid userOid) {
            ExportExcel exp = new ExportExcel(ds, eventOid, userOid);
            string xlsPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string noFindXml = "";
            int t = 0;
            for (int i = 0; i < ds.Tables.Count; i++) {
                string tbName = ds.Tables[i].TableName;
                string xlsFile = Path.Combine(xlsPath, tbName + ".xls");
                if (!File.Exists(xlsFile))
                    noFindXml += tbName.ToUpper() + "\r\n";
                else
                    t++;
            }
            if (t == 0) {
                PLMEventLog.WriteLog(string.Format("由于缺少{0}导出模板，保存文件操作被取消", noFindXml), EventLogEntryType.Error);
                return;
            }
            string fPath = Path.GetFullPath("..\\ExportFiles");

            for (int i = 0; i < ds.Tables.Count; i++) {
                DataTable tb = ds.Tables[i];
                string xlsFile = Path.Combine(xlsPath, tb.TableName + ".xls");
                string newFile = Path.Combine(fPath, BomExportClient._item.Id + "_" + Path.GetFileName(xlsFile));
                if (File.Exists(xlsFile)) {
                    ds.Tables.Remove(tb);
                    ds.AcceptChanges();
                    i--;
                    exp.ExportExcle(tb, xlsFile, newFile);
                }
            }
        }

        private void ExportExcle(DataTable tb, string xlsFile, string newFile) {
            DataSet rs = new DataSet();
            rs.Tables.Add(tb);

            try {
                File.Copy(xlsFile, newFile, true);
                RegistryKey regKey = Registry.ClassesRoot.OpenSubKey("Excel.Application\\CurVer");
                string excelVer = (string)regKey.GetValue("");
                string vol = excelVer.Substring(excelVer.IndexOf(".", excelVer.IndexOf(".") + 1) + 1);
                Assembly ass = 
                    Assembly.LoadFrom(Application.StartupPath + "\\Office" + vol + "\\Thyt.TiPLM.UIL.Report" + vol +
                                      ".dll");
                Type type = ass.GetType("Thyt.TiPLM.UIL.Report" + vol + ".Report");

                ps =
                    (Process)
                    type.InvokeMember("GetRptResult",
                                      BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Static,
                                      null, null, new Object[] { rs, newFile }, null);
            } catch (Exception ex) {
                string str = ex.InnerException != null ? ex.InnerException.Message : "";
                MessageBox.Show("生成报表失败：" + ex.Message + "　" + str);
            } finally {
                try {
                    ps.Kill();
                } catch {
                }
            }
        }

    }
}
