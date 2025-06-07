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
    public class CINVOICE
    {
        #region nature
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

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
        private string _sqlfi;
        public string sqlfi
        {
            set { _sqlfi = value; }
            get { return _sqlfi; }
        }
        private string _INVOICE_DATE;
        public string INVOICE_DATE
        {
            set { _INVOICE_DATE = value; }
            get { return _INVOICE_DATE; }
        }
        private string _NO_TAX_AMOUNT;
        public string NO_TAX_AMOUNT
        {
            set { _NO_TAX_AMOUNT = value; }
            get { return _NO_TAX_AMOUNT; }
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
        private string _PICK_INVOICE_MAKER;
        public string PICK_INVOICE_MAKER
        {

            set { _PICK_INVOICE_MAKER = value; }
            get { return _PICK_INVOICE_MAKER; }

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
        private string _IVID;
        public string IVID
        {
            set { _IVID = value; }
            get { return _IVID; }
        }
        private string _PNID;
        public string PNID
        {
            set { _PNID = value; }
            get { return _PNID; }
        }
        #endregion
        #region sql
        string INKEY;
        string setsql = @"

SELECT 
C.ORDER_ID AS 订单编号,
E.CName AS 客户名称,
C.WAREID AS 品号,
C.PRODUCTION_COUNT AS 订单数量,
C.HAVE_TAX_UNIT_PRICE AS 含税单价,
CASE WHEN C.HAVE_TAX_UNIT_PRICE IS NOT NULL AND C.HAVE_TAX_UNIT_PRICE<>'' THEN C.HAVE_TAX_UNIT_PRICE*C.PRODUCTION_COUNT
ELSE 0
END  AS 订单金额,
B.INVOICE_DATE  AS 开票日期,
B.NO_TAX_AMOUNT AS 未税金额,
B.TAX_RATE AS 税率,
RTRIM(CONVERT(DECIMAL(18,2),B.NO_TAX_AMOUNT*(1+B.TAX_RATE/100))) AS 含税金额,
B.INVOICE_NO AS 发票号码,
B.PICK_INVOICE_MAKER AS 领票人,
B.PICK_INVOICE_DATE AS 领票日期,
B.REMARK AS 备注,
A.MAKERID AS 制单人编号
FROM INVOICE_MST A
LEFT JOIN INVOICE_DET B ON A.IVID=B.IVID 
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS C ON A.PNID=C.PNID
LEFT JOIN PROJECT_INFO D ON C.PIID=D.PIID
LEFT JOIN CustomerInfo_MST E ON D.CUID=E.CUID

";
        string setsqlo = @"
SELECT 
C.ORDER_ID AS 订单编号,
E.CName AS 客户名称,
C.WAREID AS 品号,
C.PRODUCTION_COUNT AS 订单数量,
C.HAVE_TAX_UNIT_PRICE AS 含税单价,
CASE WHEN C.HAVE_TAX_UNIT_PRICE IS NOT NULL AND C.HAVE_TAX_UNIT_PRICE<>'' THEN C.HAVE_TAX_UNIT_PRICE*C.PRODUCTION_COUNT
ELSE 0
END  AS 应开票金额,
SUM(B.NO_TAX_AMOUNT*(1+B.TAX_RATE/100)) AS 实开票金额,
CASE WHEN C.HAVE_TAX_UNIT_PRICE IS NOT NULL AND C.HAVE_TAX_UNIT_PRICE<>'' THEN C.HAVE_TAX_UNIT_PRICE*C.PRODUCTION_COUNT-SUM(B.NO_TAX_AMOUNT*(1+B.TAX_RATE/100))
ELSE 0
END  AS 待开金额,
CONVERT(varchar(12) , getdate(), 111 )  AS 截止日期
FROM INVOICE_MST A
LEFT JOIN INVOICE_DET B ON A.IVID=B.IVID 
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS C ON A.PNID=C.PNID
LEFT JOIN PROJECT_INFO D ON C.PIID=D.PIID
LEFT JOIN CustomerInfo_MST E ON D.CUID=E.CUID
";


        string setsqlt = @"INSERT INTO INVOICE_MST(

IVID,
PNID,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
) VALUES 

(
@IVID,
@PNID,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY

)

";
        string setsqlth = @"UPDATE INVOICE_MST SET 
IVID=@IVID,
PNID=@PNID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH,
DAY=@DAY

";
        string setsqlf = @"INSERT INTO INVOICE_DET(
IVKEY,
IVID,
SN,
INVOICE_DATE,
NO_TAX_AMOUNT,
TAX_RATE,
INVOICE_NO,
PICK_INVOICE_MAKER,
PICK_INVOICE_DATE,
REMARK,
YEAR,
MONTH,
DAY
)
VALUES (
@IVKEY,
@IVID,
@SN,
@INVOICE_DATE,
@NO_TAX_AMOUNT,
@TAX_RATE,
@INVOICE_NO,
@PICK_INVOICE_MAKER,
@PICK_INVOICE_DATE,
@REMARK,
@YEAR,
@MONTH,
@DAY
)

";
        /*含开票日期*/
        string setsqlfi = @"
SELECT 
C.ORDER_ID AS 订单编号,
E.CName AS 客户名称,
C.WAREID AS 品号,
C.PRODUCTION_COUNT AS 订单数量,
C.HAVE_TAX_UNIT_PRICE AS 含税单价,
CASE WHEN C.HAVE_TAX_UNIT_PRICE IS NOT NULL AND C.HAVE_TAX_UNIT_PRICE<>'' THEN C.HAVE_TAX_UNIT_PRICE*C.PRODUCTION_COUNT
ELSE 0
END  AS 应开票金额,
SUM(B.NO_TAX_AMOUNT*(1+B.TAX_RATE/100)) AS 实开票金额,
CASE WHEN C.HAVE_TAX_UNIT_PRICE IS NOT NULL AND C.HAVE_TAX_UNIT_PRICE<>'' THEN C.HAVE_TAX_UNIT_PRICE*C.PRODUCTION_COUNT-SUM(B.NO_TAX_AMOUNT*(1+B.TAX_RATE/100))
ELSE 0
END  AS 待开金额,
CONVERT(varchar(12) , getdate(), 111 )  AS 截止日期
FROM INVOICE_MST A
LEFT JOIN INVOICE_DET B ON A.IVID=B.IVID 
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS C ON A.PNID=C.PNID
LEFT JOIN PROJECT_INFO D ON C.PIID=D.PIID
LEFT JOIN CustomerInfo_MST E ON D.CUID=E.CUID
";
      

        #endregion
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        StringBuilder sqb = new StringBuilder();
        public CINVOICE()
        {
            IFExecution_SUCCESS = true;
            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
            sqlfi = setsqlfi;
        }
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYMD(12, 4, "0001", "select * from INVOICE_NO", "IVID", "IV");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO INVOICE_NO(IVID,DATE,YEAR,MONTH,DAY) VALUES ('" + v1 + "','"+varDate +"','"+year +"','"+month +"','"+day +"')");
            }
            return GETID;
        }
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("开票日期", typeof(string));
            dt.Columns.Add("未税金额", typeof(decimal));
            dt.Columns.Add("税率", typeof(decimal));
            dt.Columns.Add("税金", typeof(decimal));
            dt.Columns.Add("含税金额", typeof(decimal));
            dt.Columns.Add("发票号码", typeof(string));
            dt.Columns.Add("领票人", typeof(string));
            dt.Columns.Add("领票日期", typeof(string));
            dt.Columns.Add("待开金额", typeof(decimal));
            dt.Columns.Add("备注", typeof(string));
            return dt;
        }
        #endregion
        #region GetTableInfo_2
        public DataTable GetTableInfo_2()
        {
            dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("订单编号", typeof(string));
            dt.Columns.Add("客户名称", typeof(string));
            dt.Columns.Add("品号", typeof(string));
            dt.Columns.Add("订单数量", typeof(decimal));
            dt.Columns.Add("含税单价", typeof(decimal));
            dt.Columns.Add("应开票金额", typeof(decimal));
            dt.Columns.Add("实开票金额", typeof(decimal));
            dt.Columns.Add("待开金额", typeof(decimal));
            dt.Columns.Add("截止日期", typeof(string));
            return dt;
        }
        #endregion
        #region GetTableInfo_3
        public DataTable GetTableInfo_3()
        {
            dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("品号", typeof(string));
            dt.Columns.Add("品号", typeof(string));
            dt.Columns.Add("开票日期", typeof(decimal));
            dt.Columns.Add("未税金额", typeof(decimal));
            dt.Columns.Add("税率", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("税金", typeof(string));
            dt.Columns.Add("含税金额", typeof(string));
            dt.Columns.Add("发票号码", typeof(string));
            dt.Columns.Add("备注", typeof(string));
            return dt;
        }
        #endregion
        #region GetTableInfo_4
        public DataTable GetTableInfo_4()
        {
            dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("编号", typeof(string));
            dt.Columns.Add("订单编号", typeof(string));
            dt.Columns.Add("客户名称", typeof(string));
            dt.Columns.Add("品号", typeof(string));
            dt.Columns.Add("订单数量", typeof(decimal));
            dt.Columns.Add("已开票日期", typeof(decimal));
            dt.Columns.Add("已出货数量", typeof(decimal));
            dt.Columns.Add("待出货数量", typeof(decimal));
            dt.Columns.Add("税率", typeof(string));
            dt.Columns.Add("截止日期", typeof(string));
            return dt;
        }
        #endregion
        #region save
        public void save(string TABLENAME_MST, string TABLENAME_DET, string COLUMNID,
           string IDVALUE, DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID;
            basec.getcoms("DELETE " + TABLENAME_DET + " WHERE " + COLUMNID + "='" + IDVALUE + "'");
            SQlcommandE(sqlf, dt);
            if (!bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME_DET + " WHERE " + COLUMNID + "='" + IDVALUE + "'"))
            {
                return;
            }
            else if (!bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME_MST + " WHERE " + COLUMNID + "='" + IDVALUE + "'"))
            {
                SQlcommandE(
                    sqlt,
                    IDVALUE);
            }
            else
            {
                SQlcommandE(sqlth + " WHERE " + COLUMNID + "='" + IDVALUE + "'", IDVALUE);
            }
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql,DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            int i=1;
            foreach (DataRow dr in dt.Rows)
            {
               
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                INKEY = bc.numYMD(20, 12, "000000000001", "select * from INVOICE_DET", "IVKEY", "IV");
                sqlcom.Parameters.Add("@IVKEY", SqlDbType.VarChar, 20).Value = INKEY;
                sqlcom.Parameters.Add("@IVID", SqlDbType.VarChar, 20).Value = IVID;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = i.ToString();
                DateTime date1 = Convert.ToDateTime(dr["开票日期"].ToString());
           
                sqlcom.Parameters.Add("@INVOICE_DATE", SqlDbType.VarChar, 20).Value = date1.ToString("yyyy/MM/dd").Replace("-", "/");
                if (!string.IsNullOrEmpty(dr["未税金额"].ToString()))
                {
                    sqlcom.Parameters.Add("@NO_TAX_AMOUNT", SqlDbType.VarChar, 20).Value = dr["未税金额"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("@NO_TAX_AMOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr["税率"].ToString()))
                {
                    sqlcom.Parameters.Add("@TAX_RATE", SqlDbType.VarChar, 20).Value = dr["税率"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("@TAX_RATE", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("@INVOICE_NO", SqlDbType.VarChar, 20).Value = dr["发票号码"].ToString();
                if (dr["领票日期"].ToString() != "")
                {
                    DateTime date2 = Convert.ToDateTime(dr["领票日期"].ToString());
                    sqlcom.Parameters.Add("@PICK_INVOICE_DATE", SqlDbType.VarChar, 20).Value = date2.ToString("yyyy/MM/dd").Replace("-", "/");
                }
                else
                {
                    sqlcom.Parameters.Add("@PICK_INVOICE_DATE", SqlDbType.VarChar, 20).Value = dr["领票日期"].ToString();
                }

                sqlcom.Parameters.Add("@PICK_INVOICE_MAKER", SqlDbType.VarChar, 20).Value = dr["领票人"].ToString();
                sqlcom.Parameters.Add("@REMARK", SqlDbType.VarChar, 1000).Value = dr["备注"].ToString();
                sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcon.Open();
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
                i = i + 1;
            }
        
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql, string v1)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@IVID", SqlDbType.VarChar, 20).Value = v1;
            sqlcom.Parameters.Add("@PNID", SqlDbType.VarChar, 20).Value = PNID;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region  GET_CALCULATE
        public DataTable GET_CALCULATE(DataTable dt) /*流水账余额TABLE*/
        {
            DataTable dtt = GetTableInfo();
            if (dt.Rows.Count > 0)
            {
                decimal SUM = 0;
                int i = 1;
                decimal d3 = 0;

                foreach (DataRow dr1 in dt.Rows)
                {
                    decimal d1 = 0, d2 = 0;
                    if (!string.IsNullOrEmpty(dr1["未税金额"].ToString()))
                    {
                        d1 = decimal.Parse(dr1["未税金额"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr1["税率"].ToString()))
                    {
                        d2 = decimal.Parse(dr1["税率"].ToString());
                    }
                    d3 = d3+ d1*(1+d2/100);
                    SUM = decimal.Parse(dr1["订单金额"].ToString()) - d3;
                    DataRow dr = dtt.NewRow();
                    dr["项次"] = i.ToString();
                    dr["开票日期"] = dr1["开票日期"].ToString();
                    if (!string.IsNullOrEmpty(dr1["未税金额"].ToString()))
                    {
                        dr["未税金额"] = dr1["未税金额"].ToString();
                    }
                    else
                    {
                        dr["未税金额"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dr1["税率"].ToString()))
                    {
                        dr["税率"] = dr1["税率"].ToString();
                    }
                    else
                    {
                        dr["税率"] = DBNull.Value;
                    }
                    dr["税金"] = (decimal.Parse(dr1["未税金额"].ToString()) * decimal.Parse(dr1["税率"].ToString()) / 100).ToString ("0.00");
                    dr["含税金额"] = (decimal.Parse(dr1["未税金额"].ToString()) * (1 + decimal.Parse(dr1["税率"].ToString()) / 100)).ToString ("0.00");
                    dr["发票号码"] = dr1["发票号码"].ToString();
                    dr["领票人"] = dr1["领票人"].ToString();
                    dr["领票日期"] = dr1["领票日期"].ToString();
                    dr["待开金额"] = SUM.ToString ("0.00");
                    dr["备注"] = dr1["备注"].ToString();
                    dtt.Rows.Add(dr);
                    i = i + 1;
                }
            }
            return dtt;
        }
        #endregion
        #region RETURN_HAVE_ID_DT
        public DataTable RETURN_HAVE_ID_DT(DataTable dtx,bool IF_HAVE_ID)
        {
          DataTable dt = GetTableInfo_2();
          int i = 1;
          foreach (DataRow dr1 in dtx.Rows)
          {
              DataRow dr = dt.NewRow();
              dr["序号"] = i.ToString();
              if (IF_HAVE_ID)
              {
                  dr["编号"] = dr1["编号"].ToString();
              }
              dr["订单编号"] = dr1["订单编号"].ToString();
              dr["客户名称"] = dr1["客户名称"].ToString();
              dr["品号"] = dr1["品号"].ToString();
              dr["订单数量"] = dr1["订单数量"].ToString();
              if (!string.IsNullOrEmpty(dr1["含税单价"].ToString()))
              {
                  dr["含税单价"] = dr1["含税单价"].ToString();
              }
              else
              {
                  dr["含税单价"] = DBNull.Value;
              }
              dr["应开票金额"] = dr1["应开票金额"].ToString();
              dr["实开票金额"] = dr1["实开票金额"].ToString();
              dr["待开金额"] = dr1["待开金额"].ToString();
              dr["截止日期"] = dr1["截止日期"].ToString();
              
              dt.Rows.Add(dr);
              i = i + 1;
          }
            return dt;
        }
        #endregion
        #region EXCEL_DT_TO_CSHART_D
        public DataTable EXCEL_DT_TO_CSHART_DT(DataTable dtx)
        {
            DataTable dt = GetTableInfo_3();
    
            for (int i = 1; i < dtx.Rows.Count; i++)
            {
               
                DataRow dr = dt.NewRow();
                dr["品号"] = dtx.Rows[i]["F1"].ToString();
                dr["品号"] = dtx.Rows[i]["F2"].ToString();
                if (!string.IsNullOrEmpty(dtx.Rows [i]["F4"].ToString()))
                {
                    dr["开票日期"] = dtx.Rows[i]["F4"].ToString();
                }
                else
                {
                    dr["开票日期"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dtx.Rows [i]["F5"].ToString()))
                {
                    dr["未税金额"] = dtx.Rows[i]["F5"].ToString();
                }
                else
                {
                    dr["未税金额"] = DBNull.Value;
                }
                dr["税金"] = dtx.Rows[i]["F7"].ToString();
                dr["含税金额"] = dtx.Rows[i]["F8"].ToString();
                dr["日期"] = dtx.Rows[i]["F9"].ToString();
                dt.Rows.Add(dr);
            }
            return dt;
        }
        #endregion
        #region showdata
        public void showdata(string path)
        {
            DataSet ds = new DataSet();
            string tablename = ExcelToCSHARP.GetExcelFirstTableName(path);
            ds = ExcelToCSHARP.importExcelToDataSet(path, tablename);
            DataTable dt = ds.Tables[0];
            dt = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "F1 IS NOT NULL");
            dt = EXCEL_DT_TO_CSHART_DT(dt);
            IVID = GETID();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (IVID == "Exceed Limited")
                {
                    MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (JuageFormat(i))
                {
                }
                else
                {
                    PNID = "PN16050016";
                    save("INVOICE_MST", "INVOICE_DET", "IVID", IVID, dt);
                    IFExecution_SUCCESS = true;
                }
            }
        }
        #endregion
        #region JuageFormat()
        public bool JuageFormat(int i)
        {

            bool b = false;
            return b;
        }
        #endregion
    
    }

}
