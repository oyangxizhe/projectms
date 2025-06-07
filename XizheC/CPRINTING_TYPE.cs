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
    public class CPRINTING_TYPE
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _SQUARE_PRICE;
        public string SQUARE_PRICE
        {
            set { _SQUARE_PRICE = value; }
            get { return _SQUARE_PRICE; }

        }
        private string _OUT_OF_PRINT;
        public string OUT_OF_PRINT
        {
            set { _OUT_OF_PRINT = value; }
            get { return _OUT_OF_PRINT; }

        }
        private string _CTP_EDITION;
        public string CTP_EDITION
        {
            set { _CTP_EDITION = value; }
            get { return _CTP_EDITION; }

        }
        private string _SUN_SCREEN_INK;
        public string SUN_SCREEN_INK
        {
            set { _SUN_SCREEN_INK = value; }
            get { return _SUN_SCREEN_INK; }

        }
        private string _TAX_RATE;
        public string TAX_RATE
        {
            set { _TAX_RATE = value; }
            get { return _TAX_RATE; }

        }

        private string _MONOCHROME_PRINTING;
        public string MONOCHROME_PRINTING
        {
            set { _MONOCHROME_PRINTING = value; }
            get { return _MONOCHROME_PRINTING; }

        }
        private string _PTID;
        public string PTID
        {
            set { _PTID = value; }
            get { return _PTID; }

        }
        private string _SIZE;
        public string SIZE
        {
            set { _SIZE = value; }
            get { return _SIZE; }
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
        private string _MIN_PRINTING;
        public string MIN_PRINTING
        {
            set { _MIN_PRINTING = value; }
            get { return _MIN_PRINTING; }

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
        private string _PTKEY;
        public string PTKEY
        {
            set { _PTKEY = value; }
            get { return _PTKEY; }

        }
        private string _MACHINE_TYPE;
        public string MACHINE_TYPE
        {
            set { _MACHINE_TYPE = value; }
            get { return _MACHINE_TYPE; }

        }
        private  bool _IFExecutionSUCCESS;
        public  bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
     

        private string _SN;
        public string SN
        {
            set { _SN = value; }
            get { return _SN; }

        }
        private string _CUSTOMER_TYPE;
        public string CUSTOMER_TYPE
        {
            set { _CUSTOMER_TYPE = value; }
            get { return _CUSTOMER_TYPE; }

        }
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }



        private string _MACHINE_FREE;
        public string MACHINE_FREE
        {

            set { _MACHINE_FREE = value; }
            get { return _MACHINE_FREE; }

        }
        #endregion
        DataTable dt = new DataTable();
        #region sql
        string setsql = @"
SELECT 
B.SIZE AS 尺寸,
B.MACHINE_TYPE AS 机型,
B.MIN_PRINTING AS 起印数,
A.SN AS 项次,
RTRIM(CONVERT(DECIMAL(18,2),B.MONOCHROME_PRINTING/(1+B.TAX_RATE/100)))  AS 单色印刷,
B.MONOCHROME_PRINTING AS 单色印刷含税,
RTRIM(CONVERT(DECIMAL(18,3),B.OUT_OF_PRINT/(1+B.TAX_RATE/100))) AS 超出印工,
B.OUT_OF_PRINT AS 超出印工含税,
RTRIM(CONVERT(DECIMAL(18,2),B.CTP_EDITION/(1+B.TAX_RATE/100))) AS CTP版,
B.CTP_EDITION AS CTP版含税,
RTRIM(CONVERT(DECIMAL(18,2),B.SUN_SCREEN_INK/(1+B.TAX_RATE/100))) AS 防晒油墨,
B.SUN_SCREEN_INK AS 防晒油墨含税,
RTRIM(CONVERT(DECIMAL(18,2),B.MACHINE_FREE/(1+B.TAX_RATE/100))) AS 起机费,
B.MACHINE_FREE AS 起机费含税,
A.SURFACE_PROCESSING AS 表面处理,
RTRIM(CONVERT(DECIMAL(18,2),A.SURFACE_PROCESSING_PRICE/(1+B.TAX_RATE/100))) AS 表面处理单价,
A.SURFACE_PROCESSING_PRICE AS 表面处理含税,
RTRIM(CONVERT(DECIMAL(18,1),B.TAX_RATE))+'%' AS 税率,
B.CUSTOMER_TYPE AS 客户类别
FROM PRINTING_TYPE_DET A 
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID

";


        string setsqlo = @"
INSERT INTO PRINTING_TYPE_DET
(
PTKEY,
PTID,
SN,
SURFACE_PROCESSING,
SURFACE_PROCESSING_PRICE,
MakerID,
Date,
YEAR,
MONTH,
DAY
)
VALUES
(
@PTKEY,
@PTID,
@SN,
@SURFACE_PROCESSING,
@SURFACE_PROCESSING_PRICE,
@MakerID,
@Date,
@YEAR,
@MONTH,
@DAY
)


";

        string setsqlt = @"

INSERT INTO PRINTING_TYPE_MST
(
PTID,
SIZE,
MACHINE_TYPE,
MIN_PRINTING,
MONOCHROME_PRINTING,
OUT_OF_PRINT,
CTP_EDITION,
SUN_SCREEN_INK,
TAX_RATE,
MACHINE_FREE,
CUSTOMER_TYPE,
DATE,
MAKERID,
YEAR,
MONTH,
DAY
)
VALUES
(
@PTID,
@SIZE,
@MACHINE_TYPE,
@MIN_PRINTING,
@MONOCHROME_PRINTING,
@OUT_OF_PRINT,
@CTP_EDITION,
@SUN_SCREEN_INK,
@TAX_RATE,
@MACHINE_FREE,
@CUSTOMER_TYPE,
@DATE,
@MAKERID,
@YEAR,
@MONTH,
@DAY
)
";
        string setsqlth = @"
UPDATE PRINTING_TYPE_MST SET 
SIZE=@SIZE,
MACHINE_TYPE=@MACHINE_TYPE,
MIN_PRINTING=@MIN_PRINTING,
MONOCHROME_PRINTING=@MONOCHROME_PRINTING,
OUT_OF_PRINT=@OUT_OF_PRINT,
CTP_EDITION=@CTP_EDITION,
SUN_SCREEN_INK=@SUN_SCREEN_INK,
TAX_RATE=@TAX_RATE,
MACHINE_FREE=@MACHINE_FREE,
CUSTOMER_TYPE=@CUSTOMER_TYPE,
MAKERID=@MAKERID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH,
DAY=@DAY
";

        string setsqlf = @"

";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        public CPRINTING_TYPE()
        {
            string year, month, day;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            day = DateTime.Now.ToString("dd");
            //GETID =bc.numYM(10, 4, "0001", "SELECT * FROM WORKORDER_PICKING_MST", "WPID", "WP");

            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
            sqlfi = setsqlfi;
            sqlsi = setsqlsi;
        }
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("表面处理", typeof(string));
            dt.Columns.Add("表面处理单价", typeof(decimal));
            return dt;
        }
 
        #endregion
   
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from PRINTING_TYPE_MST", "PTID", "PT");
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
            string GET_CUSTOMER_TYPE = bc.getOnlyString("SELECT CUSTOMER_TYPE FROM PRINTING_TYPE_MST WHERE  PTID='" + PTID + "'");
            string GET_MACHINE_TYPE = bc.getOnlyString("SELECT MACHINE_TYPE FROM PRINTING_TYPE_MST WHERE  PTID='" + PTID + "'");

            if (!bc.exists("SELECT PTID FROM PRINTING_TYPE_DET WHERE PTID='" + PTID + "'"))
            {
                if (bc.exists("SELECT * FROM PRINTING_TYPE_MST where  MACHINE_TYPE='" + MACHINE_TYPE + "' AND CUSTOMER_TYPE='" + CUSTOMER_TYPE + "'"))
                {
                    ErrowInfo = string.Format("机型：{0} + 客户类别：{1} 组合已经存在系统", MACHINE_TYPE,CUSTOMER_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE_DET(sqlo, dt);
                    SQlcommandE_MST(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
        
            else if (GET_CUSTOMER_TYPE != CUSTOMER_TYPE || GET_MACHINE_TYPE != MACHINE_TYPE)
            {
                if (bc.exists("SELECT * FROM PRINTING_TYPE_MST where  MACHINE_TYPE='" + MACHINE_TYPE + "' AND CUSTOMER_TYPE='" + CUSTOMER_TYPE + "'"))
                {
                    ErrowInfo = string.Format("机型：{0} + 客户类别：{1} 组合已经存在系统", MACHINE_TYPE, CUSTOMER_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    SQlcommandE_DET(sqlo, dt);
                    SQlcommandE_MST(sqlth + " WHERE PTID='" + PTID + "'");
                    IFExecution_SUCCESS = true;
                }
            }
            else
            {
                SQlcommandE_DET(sqlo, dt);
                SQlcommandE_MST(sqlth + " WHERE PTID='" + PTID + "'");
                IFExecution_SUCCESS = true;
            }
        }
        #endregion
    
        #region SQlcommandE_DET
        protected void SQlcommandE_DET(string sql,DataTable dt)
        {
            decimal d1 = 0, d2 = 0;
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
         
            basec.getcoms("DELETE PRINTING_TYPE_DET WHERE PTID='"+PTID+"'");
            foreach (DataRow dr in dt.Rows)
            {
                SqlConnection sqlcon = bc.getcon();
                sqlcon.Open();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                PTKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM PRINTING_TYPE_DET", "PTKEY", "PT");
                sqlcom.Parameters.Add("@PTKEY", SqlDbType.VarChar, 20).Value = PTKEY;
                sqlcom.Parameters.Add("@PTID", SqlDbType.VarChar, 20).Value = PTID;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = dr["项次"].ToString();
                sqlcom.Parameters.Add("@SURFACE_PROCESSING", SqlDbType.VarChar, 20).Value = dr["表面处理"].ToString();
                if (!string.IsNullOrEmpty(dr["表面处理单价"].ToString()) && TAX_RATE !=null)
                {
                    d1 = decimal.Parse(dr["表面处理单价"].ToString());
                    d2 = decimal.Parse(TAX_RATE );
                    sqlcom.Parameters.Add("@SURFACE_PROCESSING_PRICE", SqlDbType.VarChar, 20).Value = d1;
                }
                else
                {
                    sqlcom.Parameters.Add("@SURFACE_PROCESSING_PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
                sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
            }
          
        }
        #endregion
        #region SQlcommandE_MST
        protected void SQlcommandE_MST(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcon.Open();
            sqlcom.Parameters.Add("@PTID", SqlDbType.VarChar, 20).Value = PTID;
            sqlcom.Parameters.Add("@SIZE", SqlDbType.VarChar, 20).Value = SIZE;
            if (!string.IsNullOrEmpty(MIN_PRINTING ))
            {
                sqlcom.Parameters.Add("@MIN_PRINTING", SqlDbType.VarChar, 20).Value = MIN_PRINTING;
            }
            else
            {
                sqlcom.Parameters.Add("@MIN_PRINTING", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            sqlcom.Parameters.Add("@MACHINE_TYPE", SqlDbType.VarChar, 20).Value = MACHINE_TYPE;
        
       
            if (!string.IsNullOrEmpty(MONOCHROME_PRINTING))
            {
                sqlcom.Parameters.Add("@MONOCHROME_PRINTING", SqlDbType.VarChar, 20).Value = MONOCHROME_PRINTING;
                
            }
            else
            {
                sqlcom.Parameters.Add("@MONOCHROME_PRINTING", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            if (!string.IsNullOrEmpty(OUT_OF_PRINT ))
            {
                sqlcom.Parameters.Add("@OUT_OF_PRINT", SqlDbType.VarChar, 20).Value = OUT_OF_PRINT;
            }
            else
            {
                sqlcom.Parameters.Add("@OUT_OF_PRINT", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            if (!string.IsNullOrEmpty(CTP_EDITION ))
            {
                sqlcom.Parameters.Add("@CTP_EDITION", SqlDbType.VarChar, 20).Value = CTP_EDITION;
            }
            else
            {
                sqlcom.Parameters.Add("@CTP_EDITION", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            if (!string.IsNullOrEmpty(SUN_SCREEN_INK ))
            {
                sqlcom.Parameters.Add("@SUN_SCREEN_INK", SqlDbType.VarChar, 20).Value = SUN_SCREEN_INK;
            }
            else
            {
                sqlcom.Parameters.Add("@SUN_SCREEN_INK", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
   
            if (!string.IsNullOrEmpty(TAX_RATE))
            {
                sqlcom.Parameters.Add("@TAX_RATE", SqlDbType.VarChar, 20).Value = TAX_RATE;
            }
            else
            {
                sqlcom.Parameters.Add("@TAX_RATE", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            if (!string.IsNullOrEmpty(MACHINE_FREE))
            {
                sqlcom.Parameters.Add("@MACHINE_FREE", SqlDbType.VarChar, 20).Value =MACHINE_FREE;
            }
            else
            {
                sqlcom.Parameters.Add("@MACHINE_FREE", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            sqlcom.Parameters.Add("@CUSTOMER_TYPE", SqlDbType.VarChar, 20).Value = CUSTOMER_TYPE;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
  
     
    
    }
}
