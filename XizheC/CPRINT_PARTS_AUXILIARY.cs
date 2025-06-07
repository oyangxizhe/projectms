using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.SqlClient;
using XizheC;
using System.Windows.Forms;

namespace XizheC
{
    public class CPRINT_PARTS_AUXILIARY:IGETID 
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

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

        private string _sqlsi;
        public string sqlsi
        {
            set { _sqlsi = value; }
            get { return _sqlsi; }

        }
        private string _MAKERID;
        public string MAKERID
        {
            set { _MAKERID = value; }
            get { return _MAKERID; }

        }

        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }

        private string _PARTS_AUXILIARY;
        public string PARTS_AUXILIARY
        {
            set { _PARTS_AUXILIARY = value; }
            get { return _PARTS_AUXILIARY; }

        }
        private string _DPID;
        public string DPID
        {
            set { _DPID = value; }
            get { return _DPID; }

        }
        private string _STARTING_PRICE;
        public string STARTING_PRICE
        {
            set { _STARTING_PRICE = value; }
            get { return _STARTING_PRICE; }

        }
        private string _STARTING_PRICE_UNIT;
        public string STARTING_PRICE_UNIT
        {
            set { _STARTING_PRICE_UNIT = value; }
            get { return _STARTING_PRICE_UNIT; }

        }
        private string _UNIT_PRICE_UNIT;
        public string UNIT_PRICE_UNIT
        {
            set { _UNIT_PRICE_UNIT = value; }
            get { return _UNIT_PRICE_UNIT; }

        }
        private string _MAX_PRICE;
        public string MAX_PRICE
        {
            set { _MAX_PRICE = value; }
            get { return _MAX_PRICE; }

        }
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }

        }
        private string _UNIT_PRICE;
        public string UNIT_PRICE
        {

            set { _UNIT_PRICE = value; }
            get { return _UNIT_PRICE; }

        }
        private string _SN;
        public string SN
        {
            set { _SN = value; }
            get { return _SN; }

        }

        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
  
        #endregion
        #region sql
        string setsql = @"
SELECT 
A.PFID AS 编号,
A.PARTS_AUXILIARY AS 配件名,
A.DOSAGE AS 用量,
A.PRICE AS 单价,
A.UNIT AS 单位,
A.REMARK AS 备注,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM PRINT_PARTS_AUXILIARY A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID
LEFT JOIN PRINTING_OFFER_MST C ON A.PFID=C.PFID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_PARTS_AUXILIARY
(
PAID,
PFID,
PARTS_AUXILIARY,
DOSAGE,
PRICE,
UNIT,
REMARK,
MakerID,
Date,
Year,
Month,
DAY
)
VALUES
(
@PAID,
@PFID,
@PARTS_AUXILIARY,
@DOSAGE,
@PRICE,
@UNIT,
@REMARK,
@MakerID,
@Date,
@Year,
@Month,
@DAY

)
";
        string setsqlth = @"




";

        string setsqlf = @"

";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        DataTable dt = new DataTable();
        CPARTS_AUXILIARY cparts_auxiliary = new CPARTS_AUXILIARY();
        public CPRINT_PARTS_AUXILIARY()
        {
            string year, month, day;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            day = DateTime.Now.ToString("dd");
            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
            sqlfi = setsqlfi;
            sqlsi = setsqlsi;
        }
        public string GETID()
        {
            string v1 = bc.numYMD(12, 4, "0001", "SELECT * FROM PRINT_PARTS_AUXILIARY", "PAID", "PA");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region save
        public void save(DataGridView dgv)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            basec.getcoms("DELETE PRINT_PARTS_AUXILIARY WHERE PFID='" + PFID + "'");
            SQlcommandE(sqlt, dgv);
            IFExecution_SUCCESS = true;
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql,DataGridView dgv)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcon.Open();
                sqlcom.Parameters.Add("PAID", SqlDbType.VarChar, 20).Value = GETID();
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                sqlcom.Parameters.Add("PARTS_AUXILIARY", SqlDbType.VarChar, 20).Value = dgv["配件名", i].FormattedValue.ToString();
                if (!string.IsNullOrEmpty(dgv["用量", i].FormattedValue.ToString()))
                {
                    sqlcom.Parameters.Add("DOSAGE", SqlDbType.VarChar, 20).Value = dgv["用量", i].FormattedValue.ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("DOSAGE", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                DataTable dtx1 = bc.getdt("SELECT * FROM PARTS_AUXILIARY A WHERE A.PARTS_AUXILIARY='" + dgv["配件名", i].FormattedValue.ToString() + "'");
                decimal d1 = 0, d2 = 0, d3 = 0;
                if (dtx1.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))
                    {
                        d1 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                    {

                        d2 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                    }

                    d3 = d1 / (1 + d2 / 100);
                    //此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10
                }
                if (i == 0 || i == 1 || i == 2 || i == 3 || i == 4 || i == 5 || i == 6 || i == 7)
                {
                    if (d3 != 0)
                    {
                        sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = d3;//前8行的单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10

                    }
                    else
                    {
                        sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;

                    }
                }
                else if (!string.IsNullOrEmpty(dgv["单价", i].FormattedValue.ToString()))
                {
                    sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = dgv["单价", i].FormattedValue.ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("UNIT", SqlDbType.VarChar, 20).Value = dgv["单位", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("REMARK", SqlDbType.VarChar, 20).Value = dgv["备注", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
                sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
            }
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("配件名", typeof(string));
            dt.Columns.Add("用量", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            dt.Columns.Add("单位", typeof(string));
            dt.Columns.Add("备注", typeof(string));
            return dt;
        }

        #endregion
        #region RETURN_NO_FREE_KEY_DT
        public DataTable RETURN_NO_FREE_KEY_DT(DataTable dt)
        {

            DataTable dtt = GetTableInfo();
            DataTable dtx = bc.getdt(cparts_auxiliary .sql );
            if (dt.Rows.Count > 0)
            {
               
                foreach (DataRow dr1 in dt.Rows)
                {
                    DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtx, "", string.Format("配件名='{0}'", dr1["配件名"].ToString()));
                    if (dtt1.Rows.Count > 0)
                    {
                        DataRow dr = dtt.NewRow();
                        dr["配件名"] = dr1["配件名"].ToString();
                        dr["用量"] = dr1["用量"].ToString();
                        dr["单价"] = dr1["单价"].ToString();
                        dr["单位"] = dr1["单位"].ToString();
                        dr["备注"] = dr1["备注"].ToString();
                        dtt.Rows.Add(dr);
                    }
                    else
                    {
                      
                    }
                }
            }

            return dtt;
        }
        #endregion
        #region RETURN_FREE_KEY_DT
        public DataTable RETURN_FREE_KEY_DT(DataTable dt)
        {

            DataTable dtt = GetTableInfo();
            DataTable dtx = bc.getdt(cparts_auxiliary .sql );
            if (dt.Rows.Count > 0)
            {

                foreach (DataRow dr1 in dt.Rows)
                {
                    DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtx, "", string.Format("配件名='{0}'", dr1["配件名"].ToString()));
                    if (dtt1.Rows.Count > 0)
                    {
                    
                    }
                    else
                    {
                        DataRow dr = dtt.NewRow();
                        dr["配件名"] = dr1["配件名"].ToString();
                        dr["用量"] = dr1["用量"].ToString();
                        dr["单价"] = dr1["单价"].ToString();
                        dr["单位"] = dr1["单位"].ToString();
                        dr["备注"] = dr1["备注"].ToString();
                        dtt.Rows.Add(dr);
                    }
                }
            }

            return dtt;
        }
        #endregion
    }
}
