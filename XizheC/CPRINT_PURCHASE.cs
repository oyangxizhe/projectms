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
    public class CPRINT_PURCHASE:IGETID 
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

        private string _PURCHASE;
        public string PURCHASE
        {
            set { _PURCHASE = value; }
            get { return _PURCHASE; }

        }
        private string _DPID;
        public string DPID
        {
            set { _DPID = value; }
            get { return _DPID; }

        }
        private string _STARTING_PURCHASE_ONE;
        public string STARTING_PURCHASE_ONE
        {
            set { _STARTING_PURCHASE_ONE = value; }
            get { return _STARTING_PURCHASE_ONE; }

        }
        private string _STARTING_PURCHASE_ONE_UNIT;
        public string STARTING_PURCHASE_ONE_UNIT
        {
            set { _STARTING_PURCHASE_ONE_UNIT = value; }
            get { return _STARTING_PURCHASE_ONE_UNIT; }

        }
        private string _UNIT_PURCHASE_ONE_UNIT;
        public string UNIT_PURCHASE_ONE_UNIT
        {
            set { _UNIT_PURCHASE_ONE_UNIT = value; }
            get { return _UNIT_PURCHASE_ONE_UNIT; }

        }
        private string _MAX_PURCHASE_ONE;
        public string MAX_PURCHASE_ONE
        {
            set { _MAX_PURCHASE_ONE = value; }
            get { return _MAX_PURCHASE_ONE; }

        }
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }

        }
        private string _UNIT_PURCHASE_ONE;
        public string UNIT_PURCHASE_ONE
        {

            set { _UNIT_PURCHASE_ONE = value; }
            get { return _UNIT_PURCHASE_ONE; }

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
A.TYPE_ONE AS 类型一,
A.PURCHASE_ONE AS 外购价一,
A.MANAGE_COST_ONE AS 管理费一,
A.SUBTOTAL_ONE AS 小计一,
A.TYPE_TWO AS 类型二,
A.PURCHASE_TWO AS 外购价二,
A.MANAGE_COST_TWO AS 管理费二,
A.SUBTOTAL_TWO AS 小计二
FROM PRINT_PURCHASE A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID
LEFT JOIN PRINTING_OFFER_MST C ON A.PFID=C.PFID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_PURCHASE
(
PPID,
PFID,
TYPE_ONE,
PURCHASE_ONE,
MANAGE_COST_ONE,
SUBTOTAL_ONE,
TYPE_TWO,
PURCHASE_TWO,
MANAGE_COST_TWO,
SUBTOTAL_TWO,
MakerID,
Date,
Year,
Month,
DAY
)
VALUES
(
@PPID,
@PFID,
@TYPE_ONE,
@PURCHASE_ONE,
@MANAGE_COST_ONE,
@SUBTOTAL_ONE,
@TYPE_TWO,
@PURCHASE_TWO,
@MANAGE_COST_TWO,
@SUBTOTAL_TWO,
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
        decimal d = 0;
        CARTIFICIAL cartificial = new CARTIFICIAL();
        public CPRINT_PURCHASE()
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
            string v1 = bc.numYMD(12, 4, "0001", "SELECT * FROM PRINT_PURCHASE", "PPID", "PP");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region save
        public void save(DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            basec.getcoms("DELETE PRINT_PURCHASE WHERE PFID='" + PFID + "'");
            SQlcommandE(sqlt, dt);
            IFExecution_SUCCESS = true;
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql,DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            foreach (DataRow dr in dt.Rows )
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcon.Open();
                sqlcom.Parameters.Add("PPID", SqlDbType.VarChar, 20).Value = GETID();
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                sqlcom.Parameters.Add("TYPE_ONE", SqlDbType.VarChar, 20).Value = dr["类型一"].ToString();
                sqlcom.Parameters.Add("PURCHASE_ONE", SqlDbType.VarChar, 20).Value = dr["外购价一"].ToString();
                sqlcom.Parameters.Add("MANAGE_COST_ONE", SqlDbType.VarChar, 20).Value = dr["管理费一"].ToString();
                sqlcom.Parameters.Add("SUBTOTAL_ONE", SqlDbType.VarChar, 20).Value = dr["小计一"].ToString();
                sqlcom.Parameters.Add("TYPE_TWO", SqlDbType.VarChar, 20).Value = dr["类型二"].ToString();
                sqlcom.Parameters.Add("PURCHASE_TWO", SqlDbType.VarChar, 20).Value = dr["外购价二"].ToString();
                sqlcom.Parameters.Add("MANAGE_COST_TWO", SqlDbType.VarChar, 20).Value = dr["管理费二"].ToString();
                sqlcom.Parameters.Add("SUBTOTAL_TWO", SqlDbType.VarChar, 20).Value = dr["小计二"].ToString();   
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

        #region emptydatatable_T
        public DataTable emptydatatable_T()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("类型一", typeof(string));
            dt.Columns.Add("外购价一", typeof(string));
            dt.Columns.Add("管理费一", typeof(string));
            dt.Columns.Add("小计一", typeof(string));
            dt.Columns.Add("类型二", typeof(string));
            dt.Columns.Add("外购价二", typeof(string));
            dt.Columns.Add("管理费二", typeof(string));
            dt.Columns.Add("小计二", typeof(string));
            return dt;
        }
        #endregion
        #region RETURN_HAVE_ID_DT
        public DataTable RETURN_HAVE_ID_DT(DataTable dtx)//用于将小数位多的外购件管理费及小计截断多余位数显示
        {
            DataTable dt = emptydatatable_T();
            int i = 1;
            foreach (DataRow dr1 in dtx.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["类型一"] = dr1["类型一"].ToString();
                dr["外购价一"] = dr1["外购价一"].ToString();
                if (!string.IsNullOrEmpty(dr1["管理费一"].ToString()))
                {
                    d = decimal.Parse(dr1["管理费一"].ToString());
                    dr["管理费一"] = d.ToString("0.00");
                }
                if (!string.IsNullOrEmpty(dr1["小计一"].ToString()))
                {
                    d = decimal.Parse(dr1["小计一"].ToString());
                    dr["小计一"] = d.ToString("0.00");
                }
                dr["类型二"] = dr1["类型二"].ToString();
                dr["外购价二"] = dr1["外购价二"].ToString();
                dr["管理费二"] = dr1["管理费二"].ToString();
                dr["小计二"] = dr1["小计二"].ToString();
                dt.Rows.Add(dr);
                i = i + 1;
            }
            return dt;
        }
        #endregion
   
    }
}
