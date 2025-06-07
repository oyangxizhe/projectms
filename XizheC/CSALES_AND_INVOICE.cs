using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using XizheC;

namespace XizheC
{
    public class CSALES_AND_INVOICE
    {
        #region nature
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        private string _RECEIVABLE_DATE;
        public string RECEIVABLE_DATE
        {
            set { _RECEIVABLE_DATE = value; }
            get { return _RECEIVABLE_DATE; }
        }
        private string _NO_TAX_UNIT_PRICE;
        public string NO_TAX_UNIT_PRICE
        {
            set { _NO_TAX_UNIT_PRICE = value; }
            get { return _NO_TAX_UNIT_PRICE; }
        }
        private string _sql;
        public string sql
        {
            set { _sql = value; }
            get { return _sql; }
        }
        private string _sqlo;
        public string sqlo
        {
            set { _sqlo = value; }
            get { return _sqlo; }

        }
        private string _sqlt;
        public string sqlt
        {
            set { _sqlt = value; }
            get { return _sqlt; }

        }
        private string _sqlth;
        public string sqlth
        {
            set { _sqlth = value; }
            get { return _sqlth; }

        }
        private string _sqlf;
        public string sqlf
        {
            set { _sqlf = value; }
            get { return _sqlf; }

        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }
        }
        private bool _IF_IMPORT;
        public bool IF_IMPORT
        {
            set { _IF_IMPORT = value; }
            get { return _IF_IMPORT; }

        }
        private string _PICK_RECEIVABLE_MAKER;
        public string PICK_RECEIVABLE_MAKER
        {

            set { _PICK_RECEIVABLE_MAKER = value; }
            get { return _PICK_RECEIVABLE_MAKER; }

        }
        private decimal _BILL_ID;
        public decimal BILL_ID
        {
            set { _BILL_ID = value; }
            get { return _BILL_ID; }
        }
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }
        }
        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }
        }
        private string _RCID;
        public string RCID
        {
            set { _RCID = value; }
            get { return _RCID; }
        }
        private string _PNID;
        public string PNID
        {
            set { _PNID = value; }
            get { return _PNID; }
        }
        #endregion
        #region sql
        string setsql = @"

";
        string setsqlo = @"

";


        string setsqlt = @"


";
        string setsqlth = @"


";
        string setsqlf = @"



";
        #endregion
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        StringBuilder sqb = new StringBuilder();
        CINVENTORY cinventory = new CINVENTORY();
        CINVOICE cinvoice = new CINVOICE();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        public CSALES_AND_INVOICE()
        {
            IFExecution_SUCCESS = true;
            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
        }
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("订单编号", typeof(string));
            dt.Columns.Add("客户名称", typeof(string));
            dt.Columns.Add("品号", typeof(string));
            dt.Columns.Add("订单数量", typeof(string));
            dt.Columns.Add("已入库", typeof(decimal));
            dt.Columns.Add("已出货", typeof(decimal));
            dt.Columns.Add("待出货", typeof(decimal));
            dt.Columns.Add("应开票", typeof(decimal));
            dt.Columns.Add("已开票", typeof(decimal));
            dt.Columns.Add("待开票", typeof(decimal));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("截止日期", typeof(string));
            return dt;
        }
        #endregion
        #region RETURN_SEARCH_DT
        public DataTable RETURN_SEARCH_DT(DataTable dtx)
        {
          DataTable dt = GetTableInfo();
          int i = 1;
          foreach (DataRow dr1 in dtx.Rows)
          {
              DataRow dr = dt.NewRow();
              dr["序号"] = i.ToString();
   
              dr["订单编号"] = dr1["订单编号"].ToString();
              dr["客户名称"] = dr1["客户名称"].ToString();
              dr["品号"] = dr1["品号"].ToString();
              dr["订单数量"] = dr1["生产数量"].ToString();
              sqb = new StringBuilder();
              sqb.AppendFormat(cinventory.sqlo);
              sqb.AppendFormat(" WHERE C.ORDER_ID='{0}'",dr1["订单编号"].ToString ());
              sqb.AppendFormat(" GROUP BY C.ORDER_ID,E.CName,C.WAREID, C.PRODUCTION_COUNT  ORDER BY C.ORDER_ID ASC");
              DataTable dtx1 = bc.getdt(sqb.ToString ());
              if (dtx1.Rows.Count > 0)
              {
                  dr["已入库"] = dtx1.Rows[0]["已入库数量"].ToString();
                  dr["已出货"] = dtx1.Rows[0]["已出货数量"].ToString();
                  dr["待出货"] = dtx1.Rows[0]["待出货数量"].ToString();
              }
              else
              {
                  dr["已入库"] = "0";
                  dr["已出货"] = "0";
                  dr["待出货"] = "0";
              }
              sqb = new StringBuilder();
              sqb.AppendFormat(cinvoice .sqlo);
              sqb.AppendFormat(" WHERE C.ORDER_ID='{0}'", dr1["订单编号"].ToString());
              sqb.Append(" GROUP BY C.ORDER_ID,E.CName,C.WAREID,C.HAVE_TAX_UNIT_PRICE, C.PRODUCTION_COUNT,B.TAX_RATE  ORDER BY C.ORDER_ID ASC");
              DataTable dtx2 = bc.getdt(sqb.ToString());
              if (dtx2.Rows.Count > 0)
              {
                  dr["应开票"] = dtx2.Rows[0]["应开票金额"].ToString();
                  dr["已开票"] = dtx2.Rows[0]["实开票金额"].ToString();
                  dr["待开票"] = dtx2.Rows[0]["待开金额"].ToString();
              }
              else
              {
                  dr["应开票"] = "0";
                  dr["已开票"] = "0";
                  dr["待开票"] = "0";
              }
              dr["截止日期"] = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
              dt.Rows.Add(dr);
              i = i + 1;
          }
            return dt;
        }
        #endregion
        #region RETURN_HAVE_ID_DT
        public DataTable RETURN_HAVE_ID_DT(DataTable dtx,string EMID,string POSITION)
        {
            DataTable dt = GetTableInfo();
            int i = 1;
            foreach (DataRow dr1 in dtx.Rows)
            {
                if (!bc.exists(ccustomer_info.sqlsi + " WHERE B.CNAME='"+dr1["客户名称"].ToString()+
                    "' AND A.USER_MAKERID='"+EMID +"' ") && POSITION =="AE")
                {

                }
                else
                {
                    DataRow dr = dt.NewRow();
                    dr["序号"] = i.ToString();
                    dr["订单编号"] = dr1["订单编号"].ToString();
                    dr["客户名称"] = dr1["客户名称"].ToString();
                    dr["品号"] = dr1["品号"].ToString();
                    dr["订单数量"] = dr1["订单数量"].ToString();
                    dr["已入库"] = dr1["已入库"].ToString();
                    dr["已出货"] = dr1["已出货"].ToString();
                    dr["待出货"] = dr1["待出货"].ToString();
                    dr["应开票"] = dr1["应开票"].ToString();
                    dr["已开票"] = dr1["已开票"].ToString();
                    dr["待开票"] = dr1["待开票"].ToString();
                    dr["截止日期"] = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
                    dt.Rows.Add(dr);
                    i = i + 1;
                }
            }
            return dt;
        }
        #endregion
    
    }

}
