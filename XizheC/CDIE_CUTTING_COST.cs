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
    public class CDIE_CUTTING_COST:IGETID 
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
        private string _TAX_RATE;
        public string TAX_RATE
        {
            set { _TAX_RATE = value; }
            get { return _TAX_RATE; }
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

        private string _DIE_CUTTING;
        public string DIE_CUTTING
        {
            set { _DIE_CUTTING = value; }
            get { return _DIE_CUTTING; }

        }
        private string _DUID;
        public string DUID
        {
            set { _DUID = value; }
            get { return _DUID; }

        }
        private string _TAX_UNIT_PRICE;
        public string TAX_UNIT_PRICE
        {
            set { _TAX_UNIT_PRICE = value; }
            get { return _TAX_UNIT_PRICE; }

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
        private string _TAX_MACHINE_COST;
        public string TAX_MACHINE_COST
        {
            set { _TAX_MACHINE_COST = value; }
            get { return _TAX_MACHINE_COST; }

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
A.DIE_CUTTING AS 项目,
RTRIM(CONVERT(DECIMAL(18,2),A.TAX_UNIT_PRICE/(1+A.TAX_RATE/100))) AS 未税单价,
CASE WHEN A.TAX_MACHINE_COST IS NOT NULL THEN RTRIM(CONVERT(DECIMAL(18,2),A.TAX_MACHINE_COST/(1+A.TAX_RATE/100)))          
ELSE ''
END AS 未税起机费,
A.TAX_UNIT_PRICE AS 含税单价,
A.UNIT AS 单位,
A.TAX_MACHINE_COST AS 含税起机费,
RTRIM(CONVERT(DECIMAL(18,0),A.TAX_RATE))+'%' AS 税率,
A.REMARK AS 说明,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM DIE_CUTTING_COST A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO DIE_CUTTING_COST
(
DUID,
DIE_CUTTING,
TAX_UNIT_PRICE,
UNIT,
TAX_MACHINE_COST,
TAX_RATE,
REMARK,
MakerID,
Date,
Year,
Month
)
VALUES
(
@DUID,
@DIE_CUTTING,
@TAX_UNIT_PRICE,
@UNIT,
@TAX_MACHINE_COST,
@TAX_RATE,
@REMARK,
@MakerID,
@Date,
@Year,
@Month
)
";
        string setsqlth = @"
UPDATE DIE_CUTTING_COST SET 
DUID=@DUID,
DIE_CUTTING=@DIE_CUTTING,
TAX_UNIT_PRICE=@TAX_UNIT_PRICE,
UNIT=@UNIT,
TAX_MACHINE_COST=@TAX_MACHINE_COST,
TAX_RATE=@TAX_RATE,
REMARK=@REMARK,
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
      
        public CDIE_CUTTING_COST()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM DIE_CUTTING_COST", "DUID", "DU");
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
            string GET_DIE_CUTTING = bc.getOnlyString("SELECT DIE_CUTTING FROM DIE_CUTTING_COST WHERE  DUID='" + DUID + "'");
            if (!bc.exists("SELECT DUID FROM DIE_CUTTING_COST WHERE DUID='" + DUID + "'"))
            {
                if (bc.exists("SELECT * FROM DIE_CUTTING_COST where DIE_CUTTING='" + DIE_CUTTING + "'"))
                {
                    ErrowInfo = string.Format("刀模：{0}" + " 已经存在系统", DIE_CUTTING);
                    IFExecution_SUCCESS = false;
                }

                else
                {
                    SQlcommandE(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_DIE_CUTTING != DIE_CUTTING)
            {
                if (bc.exists("SELECT * FROM DIE_CUTTING_COST where DIE_CUTTING='" + DIE_CUTTING + "'"))
                {
                    ErrowInfo = string.Format("刀模：{0}" + " 已经存在系统", DIE_CUTTING);
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE(sqlth + " WHERE DUID='" + DUID + "'");
                    IFExecution_SUCCESS = true;
                }
            }

            else
            {

                SQlcommandE(sqlth + " WHERE DUID='" + DUID + "'");
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
            sqlcom.Parameters.Add("DUID", SqlDbType.VarChar, 20).Value = DUID;
            sqlcom.Parameters.Add("DIE_CUTTING", SqlDbType.VarChar, 20).Value = DIE_CUTTING;
            sqlcom.Parameters.Add("TAX_UNIT_PRICE", SqlDbType.VarChar, 20).Value = TAX_UNIT_PRICE;
            sqlcom.Parameters.Add("UNIT", SqlDbType.VarChar, 20).Value = UNIT;
            if (!string.IsNullOrEmpty(TAX_MACHINE_COST))
            {
                sqlcom.Parameters.Add("TAX_MACHINE_COST", SqlDbType.VarChar, 20).Value = TAX_MACHINE_COST;
            }
            else
            {
                sqlcom.Parameters.Add("TAX_MACHINE_COST", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
      
            sqlcom.Parameters.Add("TAX_RATE", SqlDbType.VarChar, 20).Value = TAX_RATE;
            sqlcom.Parameters.Add("REMARK", SqlDbType.VarChar, 1000).Value = REMARK;
            sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
    
    }
}
