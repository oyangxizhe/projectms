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
    public class CINVOICE_AND_RECEIVABLE
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
        CINVOICE cinvoice = new CINVOICE();
        CRECEIVABLE creceivable = new CRECEIVABLE();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        public CINVOICE_AND_RECEIVABLE()
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
            dt.Columns.Add("订单金额", typeof(string));
            dt.Columns.Add("开票金额", typeof(decimal));
            dt.Columns.Add("开票日期", typeof(string));
            dt.Columns.Add("领票人", typeof(string));
            dt.Columns.Add("已回款", typeof(decimal));
            dt.Columns.Add("待收款", typeof(decimal));
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
              dr["订单数量"] = dr1["订单数量"].ToString();
              dr["订单金额"] = dr1["订单金额"].ToString();
              dr["开票金额"] = dr1["含税金额"].ToString();
              dr["开票日期"] = dr1["开票日期"].ToString();
              dr["领票人"] = dr1["领票人"].ToString();
              sqb = new StringBuilder();
              sqb.AppendFormat(creceivable.sqlo);
              sqb.AppendFormat(" WHERE C.ORDER_ID='{0}'", dr1["订单编号"].ToString());
              sqb.AppendFormat(" GROUP BY C.ORDER_ID,E.CName,C.WAREID, C.PRODUCTION_COUNT,C.HAVE_TAX_UNIT_PRICE  ORDER BY C.ORDER_ID ASC");
              DataTable dtx1 = bc.getdt(sqb.ToString());
              if (dtx1.Rows.Count > 0)
              {
                  dr["已回款"] = dtx1.Rows[0]["实收金额"].ToString();
                  dr["待收款"] = dtx1.Rows[0]["待收金额"].ToString(); ;
              }
              else
              {
                  dr["已回款"] = "0";
                  dr["待收款"] = "0";
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
                if (!bc.exists(ccustomer_info.sqlsi + " WHERE B.CNAME='" + dr1["客户名称"].ToString() +
                       "' AND A.USER_MAKERID='" + EMID + "' ") && POSITION =="AE")
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
                    dr["订单金额"] = dr1["订单金额"].ToString();
                    dr["开票金额"] = dr1["开票金额"].ToString();
                    dr["开票日期"] = dr1["开票日期"].ToString();
                    dr["领票人"] = dr1["领票人"].ToString();
                    dr["已回款"] = dr1["已回款"].ToString();
                    dr["待收款"] = dr1["待收款"].ToString();
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
