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
    public class CPRINT_PORTRAY:IGETID 
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

        private string _PORTRAY_TYPE;
        public string PORTRAY_TYPE
        {
            set { _PORTRAY_TYPE = value; }
            get { return _PORTRAY_TYPE; }

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
A.PORTRAY_TYPE AS 写真类型,
A.PORTRAY_LENGTH AS 长,
A.PORTRAY_WIDTH AS 宽,
A.TOTAL AS 总数量,
A.PRICE AS 单价,
A.SUBTOTAL AS 小计
FROM PRINT_PORTRAY A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID
LEFT JOIN PRINTING_OFFER_MST C ON A.PFID=C.PFID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_PORTRAY
(
PPID,
PFID,
PORTRAY_TYPE,
PORTRAY_LENGTH,
PORTRAY_WIDTH,
TOTAL,
PRICE,
SUBTOTAL,
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
@PORTRAY_TYPE,
@PORTRAY_LENGTH,
@PORTRAY_WIDTH,
@TOTAL,
@PRICE,
@SUBTOTAL,
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
        CPORTRAY cportray = new CPORTRAY();
        public CPRINT_PORTRAY()
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
            string v1 = bc.numYMD(12, 4, "0001", "SELECT * FROM PRINT_PORTRAY", "PPID", "PP");
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
            basec.getcoms("DELETE PRINT_PORTRAY WHERE PFID='" + PFID + "'");
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
            int i = 1;
            foreach(DataRow dr in dt.Rows )
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcon.Open();
                sqlcom.Parameters.Add("PPID", SqlDbType.VarChar, 20).Value = GETID();
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                sqlcom.Parameters.Add("SN", SqlDbType.VarChar, 20).Value = i;
                sqlcom.Parameters.Add("PORTRAY_TYPE", SqlDbType.VarChar, 20).Value = dr["写真类型"].ToString();
                sqlcom.Parameters.Add("PORTRAY_LENGTH", SqlDbType.VarChar, 20).Value = dr["长"].ToString();
                sqlcom.Parameters.Add("PORTRAY_WIDTH", SqlDbType.VarChar, 20).Value = dr["宽"].ToString();
                sqlcom.Parameters.Add("TOTAL", SqlDbType.VarChar, 20).Value = dr["总数量"].ToString();
                sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = dr["单价"].ToString();
                if (!string.IsNullOrEmpty(dr["小计"].ToString()))
                {
                    sqlcom.Parameters.Add("SUBTOTAL", SqlDbType.VarChar, 20).Value = dr["小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("SUBTOTAL", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
                sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
                i = i + 1;
            }
          
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("写真类型", typeof(string));
            dt.Columns.Add("长", typeof(string));
            dt.Columns.Add("宽", typeof(string));
            dt.Columns.Add("总数量", typeof(string));
            dt.Columns.Add("单价", typeof(decimal));
            dt.Columns.Add("小计", typeof(string));
            return dt;
        }

        #endregion
        #region RETURN_DT
        public DataTable RETURN_DT(DataTable dt)//用于前台单价截取小数位后显示 16/01/05
        {
            DataTable dtt = GetTableInfo();
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    DataRow dr = dtt.NewRow();
                    dr["写真类型"] = dr1["写真类型"].ToString();
                    dr["长"] = dr1["长"].ToString();
                    dr["宽"] = dr1["宽"].ToString();
                    dr["总数量"] = dr1["总数量"].ToString();
                    if (!string.IsNullOrEmpty(dr1["单价"].ToString()))
                    {
                        decimal d1 = decimal.Parse(dr1["单价"].ToString());
                        dr["单价"] = d1.ToString("0.00");
                    }
                    else
                    {
                        dr["单价"] = DBNull.Value;

                    }
                    dr["小计"] = dr1["小计"].ToString();
                    dtt.Rows.Add(dr);
                }
            }
            return dtt;
        }
        #endregion
    }
}
