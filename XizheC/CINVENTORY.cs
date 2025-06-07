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
    public class CINVENTORY
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
        private string _WAREID;
        public string WAREID
        {
            set { _WAREID = value; }
            get { return _WAREID; }
        }
        private string _WNAME;
        public string WNAME
        {
            set { _WNAME = value; }
            get { return _WNAME; }
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
        private string _BILL_DATE;
        public string BILL_DATE
        {

            set { _BILL_DATE = value; }
            get { return _BILL_DATE; }

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
        private string _INID;
        public string INID
        {
            set { _INID = value; }
            get { return _INID; }
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
A.INID AS 编号,
D.ORDER_ID AS 订单编号,
A.GECOUNT AS 入库数量,
A.MRCOUNT AS 出库数量,
A.BILL_DATE AS 日期,
A.BILL_TYPE AS 单据类型,
A.BILL_ID AS 单号,
C.ENAME AS 制单人,
A.REMARK AS 备注,
B.MAKERID AS 制单人编号
FROM INVENTORY_DET A
LEFT JOIN INVENTORY_MST B ON A.INID=B.INID 
LEFT JOIN EMPLOYEEINFO C ON B.MAKERID=C.EMID
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS D ON B.PNID=D.PNID


";
        string setsqlo = @"
SELECT 
C.ORDER_ID AS 订单编号,
E.CName AS 客户名称,
C.WAREID AS 品号,
ISNULL(C.PRODUCTION_COUNT,0) AS 订单数量,
CASE WHEN SUM(GECOUNT) IS NOT NULL THEN SUM(GECOUNT)
ELSE 0
END  AS 已入库数量,
CASE WHEN SUM(MRCOUNT) IS NOT NULL THEN SUM(MRCOUNT)
ELSE 0
END   AS 已出货数量,
CASE WHEN C.PRODUCTION_COUNT IS NOT NULL AND SUM(MRCOUNT) IS NOT NULL THEN C.PRODUCTION_COUNT -SUM(MRCOUNT)
ELSE ISNULL(C.PRODUCTION_COUNT,0)
END AS 待出货数量,
CASE WHEN SUM(GECOUNT) IS NOT NULL AND SUM(MRCOUNT) IS  NULL THEN SUM(GECOUNT)
WHEN SUM(GECOUNT) IS NOT NULL AND SUM(MRCOUNT) IS  NOT NULL THEN SUM(GECOUNT)-SUM(MRCOUNT)
ELSE 0
END AS 库存结余,
CONVERT(varchar(12) , getdate(), 111 )  AS 截止日期
FROM INVENTORY_MST A
LEFT JOIN INVENTORY_DET B ON A.INID=B.INID 
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS C ON A.PNID=C.PNID
LEFT JOIN PROJECT_INFO D ON C.PIID=D.PIID
LEFT JOIN CustomerInfo_MST E ON D.CUID=E.CUID


";
        string setsqlt = @"INSERT INTO INVENTORY_MST(

INID,
PNID,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
) VALUES 

(
@INID,
@PNID,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY

)

";
        string setsqlth = @"UPDATE INVENTORY_MST SET 
INID=@INID,
PNID=@PNID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH,
DAY=@DAY

";
        string setsqlf = @"INSERT INTO INVENTORY_DET(
INKEY,
INID,
SN,
WAREID,
WNAME,
GECOUNT,
MRCOUNT,
BILL_DATE,
BILL_TYPE,
BILL_ID,
REMARK,
YEAR,
MONTH,
DAY
)
VALUES (
@INKEY,
@INID,
@SN,
@WAREID,
@WNAME,
@GECOUNT,
@MRCOUNT,
@BILL_DATE,
@BILL_TYPE,
@BILL_ID,
@REMARK,
@YEAR,
@MONTH,
@DAY
)

";
        string setsqlfi = @"
SELECT 
A.INID AS 编号,
C.ORDER_ID AS 订单编号,
E.CName AS 客户名称,
C.WAREID AS 品号,
ISNULL(C.PRODUCTION_COUNT,0) AS 订单数量,
CASE WHEN SUM(GECOUNT) IS NOT NULL THEN SUM(GECOUNT)
ELSE 0
END  AS 已入库数量,
CASE WHEN SUM(MRCOUNT) IS NOT NULL THEN SUM(MRCOUNT)
ELSE 0
END   AS 已出货数量,
CASE WHEN C.PRODUCTION_COUNT IS NOT NULL AND SUM(MRCOUNT) IS NOT NULL THEN C.PRODUCTION_COUNT -SUM(MRCOUNT)
ELSE ISNULL(C.PRODUCTION_COUNT,0)
END AS 待出货数量,
CASE WHEN SUM(GECOUNT) IS NOT NULL AND SUM(MRCOUNT) IS  NULL THEN SUM(GECOUNT)
WHEN SUM(GECOUNT) IS NOT NULL AND SUM(MRCOUNT) IS  NOT NULL THEN SUM(GECOUNT)-SUM(MRCOUNT)
ELSE 0
END AS 库存结余,
CONVERT(varchar(12) , getdate(), 111 )  AS 截止日期
FROM INVENTORY_MST A
LEFT JOIN INVENTORY_DET B ON A.INID=B.INID 
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS C ON A.PNID=C.PNID
LEFT JOIN PROJECT_INFO D ON C.PIID=D.PIID
LEFT JOIN CustomerInfo_MST E ON D.CUID=E.CUID
";

        #endregion
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        CPN_PRODUCTION_INSTRUCTIONS cpn_production_instructions = new CPN_PRODUCTION_INSTRUCTIONS();
        public CINVENTORY()
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
            string v1 = bc.numYMD(12, 4, "0001", "select * from INVENTORY_NO", "INID", "IN");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO INVENTORY_NO(INID,DATE,YEAR,MONTH,DAY) VALUES ('" + v1 + "','"+varDate +"','"+year +"','"+month +"','"+day +"')");
            }
            return GETID;
        }
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("入库数量", typeof(decimal));
            dt.Columns.Add("出库数量", typeof(decimal));
            dt.Columns.Add("库存结余", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("单据类型", typeof(string));
            dt.Columns.Add("单号", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
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
            dt.Columns.Add("已入库数量", typeof(decimal));
            dt.Columns.Add("已出货数量", typeof(decimal));
            dt.Columns.Add("待出货数量", typeof(decimal));
            dt.Columns.Add("库存结余", typeof(string));
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
            dt.Columns.Add("品名", typeof(string));
            dt.Columns.Add("入库数量", typeof(decimal));
            dt.Columns.Add("出库数量", typeof(decimal));
            dt.Columns.Add("库存结余", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("单据类型", typeof(string));
            dt.Columns.Add("单号", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
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
            dt.Columns.Add("已入库数量", typeof(decimal));
            dt.Columns.Add("已出货数量", typeof(decimal));
            dt.Columns.Add("待出货数量", typeof(decimal));
            dt.Columns.Add("库存结余", typeof(string));
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
                if(IF_IMPORT)
                {
                    WAREID = dr["品号"].ToString();
                    WNAME = dr["品名"].ToString();
                }
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                INKEY = bc.numYMD(20, 12, "000000000001", "select * from INVENTORY_DET", "INKEY", "IN");
                sqlcom.Parameters.Add("@INKEY", SqlDbType.VarChar, 20).Value = INKEY;
                sqlcom.Parameters.Add("@INID", SqlDbType.VarChar, 20).Value = INID;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = i.ToString();
                sqlcom.Parameters.Add("@WAREID", SqlDbType.VarChar, 20).Value = WAREID;
                sqlcom.Parameters.Add("@WNAME", SqlDbType.VarChar, 50).Value = WNAME;
                if (!string.IsNullOrEmpty(dr["入库数量"].ToString()))
                {   sqlcom.Parameters.Add("@GECOUNT", SqlDbType.VarChar, 20).Value = dr["入库数量"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("@GECOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr["出库数量"].ToString()))
                {
                    sqlcom.Parameters.Add("@MRCOUNT", SqlDbType.VarChar, 20).Value = dr["出库数量"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("@MRCOUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("@BILL_DATE", SqlDbType.VarChar, 20).Value =dr["日期"].ToString();
                sqlcom.Parameters.Add("@BILL_TYPE", SqlDbType.VarChar, 20).Value =dr["单据类型"].ToString();
                sqlcom.Parameters.Add("@BILL_ID", SqlDbType.VarChar, 20).Value =dr["单号"].ToString();
                sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.Parameters.Add("@REMARK", SqlDbType.VarChar, 1000).Value = dr["备注"].ToString();
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
            sqlcom.Parameters.Add("@INID", SqlDbType.VarChar, 20).Value = v1;
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
                foreach (DataRow dr1 in dt.Rows)
                {
                    decimal d1 = 0, d2 = 0;
                 
                    if (!string.IsNullOrEmpty(dr1["入库数量"].ToString()))
                    {
                        d1 = decimal.Parse(dr1["入库数量"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr1["出库数量"].ToString()))
                    {
                        d2 = decimal.Parse(dr1["出库数量"].ToString());
                    }
                    SUM = SUM + d1 - d2;
                    DataRow dr = dtt.NewRow();
                    dr["项次"] = i.ToString();
                    if (!string.IsNullOrEmpty(dr1["入库数量"].ToString()))
                    {
                        dr["入库数量"] = dr1["入库数量"].ToString();
                    }
                    else
                    {
                        dr["入库数量"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dr1["出库数量"].ToString()))
                    {
                        dr["出库数量"] = dr1["出库数量"].ToString();
                    }
                    else
                    {
                        dr["出库数量"] = DBNull.Value;
                    }
                    dr["库存结余"] = SUM;
                    dr["日期"] = dr1["日期"].ToString();
                    dr["单据类型"] = dr1["单据类型"].ToString();
                    dr["单号"] = dr1["单号"].ToString();
                    dr["制单人"] = dr1["制单人"].ToString();
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
            if (IF_HAVE_ID)
            {
                DataTable dt = GetTableInfo_4();
            }
            else
            {
                DataTable dt = GetTableInfo_2();
            }
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
                dr["已入库数量"] = dr1["已入库数量"].ToString();
                dr["已出货数量"] = dr1["已出货数量"].ToString();
                dr["待出货数量"] = dr1["待出货数量"].ToString();
                dr["库存结余"] = dr1["库存结余"].ToString();
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
                dr["品名"] = dtx.Rows[i]["F2"].ToString();
                if (!string.IsNullOrEmpty(dtx.Rows [i]["F4"].ToString()))
                {
                    dr["入库数量"] = dtx.Rows[i]["F4"].ToString();
                }
                else
                {
                    dr["入库数量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dtx.Rows [i]["F5"].ToString()))
                {
                    dr["出库数量"] = dtx.Rows[i]["F5"].ToString();
                }
                else
                {
                    dr["出库数量"] = DBNull.Value;
                }
                dr["单据类型"] = dtx.Rows[i]["F7"].ToString();
                dr["单号"] = dtx.Rows[i]["F8"].ToString();
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
            INID = GETID();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (INID == "Exceed Limited")
                {
                    MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (JuageFormat(i))
                {
                }
                else
                {
                    DataTable dtx = bc.getdt(cpn_production_instructions.sql + " WHERE A.ORDER_ID='"+dt.Rows [i][0].ToString ()+"'");
                    if (dtx.Rows.Count > 0)
                    {
                        PNID = dtx.Rows[0]["编号"].ToString();
                    }
                    else
                    {
                        PNID = cpn_production_instructions.GETID();
                        bc.getcom(@"
INSERT INTO 
PN_PRODUCTION_INSTRUCTIONS
(PNID,ORDER_ID,WAREID) 
VALUES ('" + PNID + "','" + dt.Rows[i][0].ToString() + "','" + dt.Rows[i]["品号"].ToString() + "')");
                    }
                    save("INVENTORY_MST", "INVENTORY_DET", "INID", INID, dt);
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
