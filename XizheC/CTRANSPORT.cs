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
    public class CTRANSPORT:IGETID 
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _CUSTOMER_TYPE;
        public string CUSTOMER_TYPE
        {
            set { _CUSTOMER_TYPE = value; }
            get { return _CUSTOMER_TYPE; }
        }
        private string _TAX_UNIT_PRICE_TWO;
        public string TAX_UNIT_PRICE_TWO
        {
            set { _TAX_UNIT_PRICE_TWO = value; }
            get { return _TAX_UNIT_PRICE_TWO; }

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

        private string _TRANSPORT;
        public string TRANSPORT
        {
            set { _TRANSPORT = value; }
            get { return _TRANSPORT; }

        }
        private string _TRID;
        public string TRID
        {
            set { _TRID = value; }
            get { return _TRID; }

        }
        private string _TAX_UNIT_PRICE_ONE;
        public string TAX_UNIT_PRICE_ONE
        {
            set { _TAX_UNIT_PRICE_ONE = value; }
            get { return _TAX_UNIT_PRICE_ONE; }

        }
        private string _UNIT;
        public string UNIT
        {
            set { _UNIT = value; }
            get { return _UNIT; }

        }
        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

        }
        private string _TAX_TRANSPORT_COST;
        public string TAX_TRANSPORT_COST
        {
            set { _TAX_TRANSPORT_COST = value; }
            get { return _TAX_TRANSPORT_COST; }

        }
        private string _TAX_RATE;
        public string TAX_RATE
        {
            set { _TAX_RATE = value; }
            get { return _TAX_RATE; }

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
A.TRANSPORT AS 物流运输,
RTRIM(CONVERT(DECIMAL(18,2),A.TAX_UNIT_PRICE_ONE/(1+A.TAX_RATE/100))) AS 未税单价一,
RTRIM(CONVERT(DECIMAL(18,2),A.TAX_UNIT_PRICE_TWO/(1+A.TAX_RATE/100))) AS 未税单价二,
CASE WHEN A.TAX_TRANSPORT_COST IS NOT NULL THEN RTRIM(CONVERT(DECIMAL(18,2),
A.TAX_TRANSPORT_COST/(1+A.TAX_RATE/100)))          
ELSE ''
END AS 未税起运费,
A.TAX_UNIT_PRICE_ONE AS 含税单价一,
A.TAX_UNIT_PRICE_TWO AS 含税单价二,
A.TAX_TRANSPORT_COST AS 含税起运费,
RTRIM(CONVERT(DECIMAL(18,0),A.TAX_RATE))+'%' AS 税率,
A.REMARK AS 说明,
A.CUSTOMER_TYPE AS 客户类别,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM TRANSPORT A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO TRANSPORT
(
TRID,
TRANSPORT,
TAX_UNIT_PRICE_ONE,
TAX_UNIT_PRICE_TWO,
TAX_TRANSPORT_COST,
TAX_RATE,
REMARK,
CUSTOMER_TYPE,
MakerID,
Date,
Year,
Month
)
VALUES
(
@TRID,
@TRANSPORT,
@TAX_UNIT_PRICE_ONE,
@TAX_UNIT_PRICE_TWO,
@TAX_TRANSPORT_COST,
@TAX_RATE,
@CUSTOMER_TYPE,
@REMARK,
@MakerID,
@Date,
@Year,
@Month
)
";
        string setsqlth = @"
UPDATE TRANSPORT SET 
TRID=@TRID,
TRANSPORT=@TRANSPORT,
TAX_UNIT_PRICE_ONE=@TAX_UNIT_PRICE_ONE,
TAX_UNIT_PRICE_TWO=@TAX_UNIT_PRICE_TWO,
TAX_TRANSPORT_COST=@TAX_TRANSPORT_COST,
TAX_RATE=@TAX_RATE,
REMARK=@REMARK,
CUSTOMER_TYPE=@CUSTOMER_TYPE,
MakerID=@MakerID,
Date=@Date,
Year=@Year,
Month=@Month

";

        string setsqlf = @"

";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        DataTable dt = new DataTable();
      
        public CTRANSPORT()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM TRANSPORT", "TRID", "TR");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region save
        public void save()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string GET_TRANSPORT = bc.getOnlyString("SELECT TRANSPORT FROM TRANSPORT WHERE  TRID='" + TRID + "'");
            string GET_CUSTOMER_TYPE = bc.getOnlyString("SELECT CUSTOMER_TYPE FROM TRANSPORT WHERE  TRID='" + TRID + "'");
            if (!bc.exists("SELECT TRID FROM TRANSPORT WHERE TRID='" + TRID + "'"))
            {
                if (bc.exists("SELECT * FROM TRANSPORT where TRANSPORT='" + TRANSPORT + "' AND CUSTOMER_TYPE='"+CUSTOMER_TYPE +"'"))
                {
                    ErrowInfo = string.Format("物流运输：{0} + 客户类别：{1} 组合" + "已经存在系统", TRANSPORT,CUSTOMER_TYPE );
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    SQlcommandE(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_TRANSPORT != TRANSPORT || GET_CUSTOMER_TYPE !=CUSTOMER_TYPE )
            {
                if (bc.exists("SELECT * FROM TRANSPORT where TRANSPORT='" + TRANSPORT + "' AND CUSTOMER_TYPE='" + CUSTOMER_TYPE + "'"))
                {
                    ErrowInfo = string.Format("物流运输：{0} + 客户类别：{1} 组合" + "已经存在系统", TRANSPORT,CUSTOMER_TYPE );
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE(sqlth + " WHERE TRID='" + TRID + "'");
                    IFExecution_SUCCESS = true;
                }
            }

            else
            {

                SQlcommandE(sqlth + " WHERE TRID='" + TRID + "'");
                IFExecution_SUCCESS = true;
            }
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcon.Open();
            sqlcom.Parameters.Add("TRID", SqlDbType.VarChar, 20).Value = TRID;
            sqlcom.Parameters.Add("TRANSPORT", SqlDbType.VarChar, 20).Value = TRANSPORT;
            sqlcom.Parameters.Add("TAX_UNIT_PRICE_ONE", SqlDbType.VarChar, 20).Value = TAX_UNIT_PRICE_ONE;
            sqlcom.Parameters.Add("TAX_UNIT_PRICE_TWO", SqlDbType.VarChar, 20).Value = TAX_UNIT_PRICE_TWO;
            sqlcom.Parameters.Add("TAX_TRANSPORT_COST", SqlDbType.VarChar, 20).Value = TAX_TRANSPORT_COST;
            sqlcom.Parameters.Add("TAX_RATE", SqlDbType.VarChar, 20).Value = TAX_RATE;
            sqlcom.Parameters.Add("REMARK", SqlDbType.VarChar, 1000).Value = REMARK;
            sqlcom.Parameters.Add("CUSTOMER_TYPE", SqlDbType.VarChar, 20).Value = CUSTOMER_TYPE;
            sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region emptydatatable_T
        public DataTable emptydatatable_T()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("物流运输", typeof(string));
            dt.Columns.Add("未税单价一", typeof(string));
            dt.Columns.Add("未税单价二", typeof(string));
            dt.Columns.Add("未税起运费", typeof(string));
            dt.Columns.Add("税率", typeof(string));
            dt.Columns.Add("含税单价一", typeof(string));
            dt.Columns.Add("含税单价二", typeof(string));
            dt.Columns.Add("含税起运费", typeof(string));
            dt.Columns.Add("说明", typeof(string));
            dt.Columns.Add("客户类别", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            dt.Columns.Add("制单日期", typeof(string));
            return dt;
        }
        #endregion
        #region RETURN_HAVE_ID_DT
        public DataTable RETURN_HAVE_ID_DT(DataTable dtx)
        {
            DataTable dt = emptydatatable_T();
            int i = 1;
            foreach (DataRow dr1 in dtx.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["序号"] = i.ToString();
                dr["物流运输"] = dr1["物流运输"].ToString();
                dr["未税单价一"] = dr1["未税单价一"].ToString();
                dr["未税单价二"] = dr1["未税单价二"].ToString();
                dr["未税起运费"] = dr1["未税起运费"].ToString();
                dr["税率"] = dr1["税率"].ToString();
                dr["含税单价一"] = dr1["含税单价一"].ToString();
                dr["含税单价二"] = dr1["含税单价二"].ToString();
                dr["含税起运费"] = dr1["含税起运费"].ToString();
                dr["说明"] = dr1["说明"].ToString();
                dr["客户类别"] = dr1["客户类别"].ToString();
                dr["制单人"] = dr1["制单人"].ToString();
                dr["制单日期"] = dr1["制单日期"].ToString();
                dt.Rows.Add(dr);
                i = i + 1;
            }
            return dt;
        }
        #endregion
    }
}
