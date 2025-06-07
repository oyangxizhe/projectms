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
    public class CMATERIAL_PRICE:IGETID 
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

        private string _MATERIAL_TYPE;
        public string MATERIAL_TYPE
        {
            set { _MATERIAL_TYPE = value; }
            get { return _MATERIAL_TYPE; }

        }
        private string _MRID;
        public string MRID
        {
            set { _MRID = value; }
            get { return _MRID; }

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
        private string _MAX_PRICE_UNIT;
        public string MAX_PRICE_UNIT
        {
            set { _MAX_PRICE_UNIT = value; }
            get { return _MAX_PRICE_UNIT; }

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
A.MATERIAL_TYPE AS 类型,
A.STARTING_PRICE AS 起步价,
A.STARTING_PRICE_UNIT AS 起步价单位,
A.UNIT_PRICE AS 单位计价,
A.UNIT_PRICE_UNIT AS 单位计价单位,
A.MAX_PRICE AS 封顶金额,
A.MAX_PRICE_UNIT AS 封顶金额单位,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM MATERIAL_PRICE A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO MATERIAL_PRICE
(
MRID,
MATERIAL_TYPE,
STARTING_PRICE,
STARTING_PRICE_UNIT,
UNIT_PRICE,
UNIT_PRICE_UNIT,
MAX_PRICE,
MAX_PRICE_UNIT,
MakerID,
Date,
Year,
Month
)
VALUES
(
@MRID,
@MATERIAL_TYPE,
@STARTING_PRICE,
@STARTING_PRICE_UNIT,
@UNIT_PRICE,
@UNIT_PRICE_UNIT,
@MAX_PRICE,
@MAX_PRICE_UNIT,
@MakerID,
@Date,
@Year,
@Month
)
";
        string setsqlth = @"
UPDATE MATERIAL_PRICE SET 
MRID=@MRID,
MATERIAL_TYPE=@MATERIAL_TYPE,
STARTING_PRICE=@STARTING_PRICE,
STARTING_PRICE_UNIT=@STARTING_PRICE_UNIT,
UNIT_PRICE=@UNIT_PRICE,
UNIT_PRICE_UNIT=@UNIT_PRICE_UNIT,
MAX_PRICE=@MAX_PRICE,
MAX_PRICE_UNIT=@MAX_PRICE_UNIT,
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
       
      
        public CMATERIAL_PRICE()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM MATERIAL_PRICE", "MRID", "MR");
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
            string GET_MATERIAL_TYPE = bc.getOnlyString("SELECT MATERIAL_TYPE FROM MATERIAL_PRICE WHERE  MRID='" + MRID + "'");

            if (!bc.exists("SELECT MRID FROM MATERIAL_PRICE WHERE MRID='" + MRID + "'"))
            {
                if (bc.exists("SELECT * FROM MATERIAL_PRICE where MATERIAL_TYPE='" + MATERIAL_TYPE+ "'"))
                {
                    ErrowInfo = string.Format("类型：{0}" + " 已经存在系统", MATERIAL_TYPE);
                    IFExecution_SUCCESS = false;
              
                }

                else
                {
                    SQlcommandE(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_MATERIAL_TYPE != MATERIAL_TYPE)
            {
                if (bc.exists("SELECT * FROM MATERIAL_PRICE where MATERIAL_TYPE='" + MATERIAL_TYPE + "'"))
                {

                    ErrowInfo = string.Format("类型：{0}" + " 已经存在系统", MATERIAL_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE(sqlth + " WHERE MRID='" + MRID + "'");
                    IFExecution_SUCCESS = true;
                }
            }

            else
            {

                SQlcommandE(sqlth + " WHERE MRID='" + MRID + "'");
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
            sqlcom.Parameters.Add("MRID", SqlDbType.VarChar, 20).Value = MRID;
            sqlcom.Parameters.Add("MATERIAL_TYPE", SqlDbType.VarChar, 20).Value = MATERIAL_TYPE;
            sqlcom.Parameters.Add("STARTING_PRICE", SqlDbType.VarChar, 20).Value = STARTING_PRICE;
            sqlcom.Parameters.Add("STARTING_PRICE_UNIT", SqlDbType.VarChar, 20).Value = STARTING_PRICE_UNIT;
            sqlcom.Parameters.Add("UNIT_PRICE", SqlDbType.VarChar, 20).Value = UNIT_PRICE;
            sqlcom.Parameters.Add("UNIT_PRICE_UNIT", SqlDbType.VarChar, 20).Value = UNIT_PRICE_UNIT;
            sqlcom.Parameters.Add("MAX_PRICE", SqlDbType.VarChar, 20).Value = MAX_PRICE;
            sqlcom.Parameters.Add("MAX_PRICE_UNIT", SqlDbType.VarChar, 20).Value = MAX_PRICE_UNIT;
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
