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
    public class CPRINT_TRANSPORT:IGETID 
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

        private string _TOTAL_BOXS_COUNT;
        public string TOTAL_BOXS_COUNT
        {
            set { _TOTAL_BOXS_COUNT = value; }
            get { return _TOTAL_BOXS_COUNT; }

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
A.PACK_LENGTH AS 长,
A.PACK_WIDTH AS 宽,
A.PACK_HEIGHT AS 高,
A.TOTAL_BOXS_COUNT AS 总箱数,
A.TOTAL_CUBIC_NUMBER AS 总立方数,
A.TRANSPORT AS 运输方式,
A.PRICE AS 单价,
A.SUBTOTAL AS 小计
FROM PRINT_TRANSPORT A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID
LEFT JOIN PRINTING_OFFER_MST C ON A.PFID=C.PFID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_TRANSPORT
(
PTID,
PFID,
TOTAL_BOXS_COUNT,
TRANSPORT,
PACK_LENGTH,
PACK_WIDTH,
PACK_HEIGHT,
TOTAL_CUBIC_NUMBER,
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
@PTID,
@PFID,
@TOTAL_BOXS_COUNT,
@TRANSPORT,
@PACK_LENGTH,
@PACK_WIDTH,
@PACK_HEIGHT,
@TOTAL_CUBIC_NUMBER,
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
        CTRANSPORT cTRANSPORT = new CTRANSPORT();
        public CPRINT_TRANSPORT()
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
            string v1 = bc.numYMD(12, 4, "0001", "SELECT * FROM PRINT_TRANSPORT", "PTID", "PT");
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
            basec.getcoms("DELETE PRINT_TRANSPORT WHERE PFID='" + PFID + "'");
            SQlcommandE(sqlt, dt);
            IFExecution_SUCCESS = true;
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql, DataTable dt)
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
                sqlcom.Parameters.Add("PTID", SqlDbType.VarChar, 20).Value = GETID();
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                if (!string.IsNullOrEmpty(dr["总箱数"].ToString()))
                {
                    sqlcom.Parameters.Add("TOTAL_BOXS_COUNT", SqlDbType.VarChar, 20).Value = dr["总箱数"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("TOTAL_BOXS_COUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("TRANSPORT", SqlDbType.VarChar, 20).Value = dr["运输方式"].ToString();
                sqlcom.Parameters.Add("PACK_LENGTH", SqlDbType.VarChar, 20).Value = dr["长"].ToString();
                sqlcom.Parameters.Add("PACK_WIDTH", SqlDbType.VarChar, 20).Value = dr["宽"].ToString();
                sqlcom.Parameters.Add("PACK_HEIGHT", SqlDbType.VarChar, 20).Value = dr["高"].ToString();
                sqlcom.Parameters.Add("TOTAL_CUBIC_NUMBER", SqlDbType.VarChar, 20).Value = dr["总立方数"].ToString();
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
            }
        }
        #endregion
       
    }
}
