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
    public class CMACHINING
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
        private string _DIE_CUTTING;
        public string DIE_CUTTING
        {
            set { _DIE_CUTTING = value; }
            get { return _DIE_CUTTING; }

        }
        private string _TAX_RATE;
        public string TAX_RATE
        {
            set { _TAX_RATE = value; }
            get { return _TAX_RATE; }

        }
        private string _CUSTOMER_TYPE;
        public string CUSTOMER_TYPE
        {
            set { _CUSTOMER_TYPE = value; }
            get { return _CUSTOMER_TYPE; }
        }
        private string _MAID;
        public string MAID
        {
            set { _MAID = value; }
            get { return _MAID; }

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
        private string _MAKEY;
        public string MAKEY
        {
            set { _MAKEY = value; }
            get { return _MAKEY; }

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
   
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }

        private string _SUN_SCREEN_INK;
        public string SUN_SCREEN_INK
        {
            set { _SUN_SCREEN_INK = value; }
            get { return _SUN_SCREEN_INK; }

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
A.SN AS 项次,
RTRIM(CONVERT(DECIMAL(18,3),B.DIE_CUTTING/(1+B.TAX_RATE/100))) AS 模切,
B.DIE_CUTTING AS 模切含税,
RTRIM(CONVERT(DECIMAL(18,2),B.MACHINE_FREE/(1+B.TAX_RATE/100))) AS 起机费,
B.MACHINE_FREE AS 起机费含税,
A.LAMINATING_PROCESS AS 裱纸,
RTRIM(CONVERT(DECIMAL(18,2),A.LAMINATING_PROCESS_PRICE/(1+B.TAX_RATE/100))) AS 裱纸单价,
RTRIM(CONVERT(DECIMAL(18,1),B.TAX_RATE))+'%' AS 税率,
A.LAMINATING_PROCESS_PRICE AS 裱纸含税价,
B.CUSTOMER_TYPE AS 客户类别
FROM MACHINING_DET A 
LEFT JOIN MACHINING_MST B ON A.MAID=B.MAID


";


        string setsqlo = @"
INSERT INTO MACHINING_DET
(
MAKEY,
MAID,
SN,
LAMINATING_PROCESS,
LAMINATING_PROCESS_PRICE,
MakerID,
Date,
YEAR,
MONTH,
DAY
)
VALUES
(
@MAKEY,
@MAID,
@SN,
@LAMINATING_PROCESS,
@LAMINATING_PROCESS_PRICE,
@MakerID,
@Date,
@YEAR,
@MONTH,
@DAY
)


";

        string setsqlt = @"

INSERT INTO MACHINING_MST
(
MAID,
SIZE,
MACHINE_TYPE,
DIE_CUTTING,
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
@MAID,
@SIZE,
@MACHINE_TYPE,
@DIE_CUTTING,
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
UPDATE MACHINING_MST SET 
SIZE=@SIZE,
MACHINE_TYPE=@MACHINE_TYPE,
DIE_CUTTING=@DIE_CUTTING,
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
        public CMACHINING()
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
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("裱纸", typeof(string));
            dt.Columns.Add("裱纸含税价", typeof(decimal));
            return dt;
        }
 
        #endregion
   
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from MACHINING_MST", "MAID", "MA");
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
            string GET_CUSTOMER_TYPE = bc.getOnlyString("SELECT CUSTOMER_TYPE FROM MACHINING_MST WHERE  MAID='" + MAID + "'");
            string GET_MACHINE_TYPE = bc.getOnlyString("SELECT MACHINE_TYPE FROM MACHINING_MST WHERE  MAID='" + MAID + "'");
            if (!bc.exists("SELECT MAID FROM MACHINING_DET WHERE MAID='" + MAID + "'"))
            {
               if (bc.exists("SELECT * FROM MACHINING_MST where MACHINE_TYPE='" + MACHINE_TYPE + "' AND CUSTOMER_TYPE='" + CUSTOMER_TYPE + "' "))
                {
                    ErrowInfo = string.Format("机型：{0}" + " + 客户类别：{1} 组合已经存在系统", MACHINE_TYPE,CUSTOMER_TYPE );
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
                if (bc.exists("SELECT * FROM MACHINING_MST where MACHINE_TYPE='" + MACHINE_TYPE + "' AND CUSTOMER_TYPE='" + CUSTOMER_TYPE + "' "))
                {
                    ErrowInfo = string.Format("机型：{0}" + " + 客户类别：{1} 组合已经存在系统", MACHINE_TYPE, CUSTOMER_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    SQlcommandE_DET(sqlo, dt);
                    SQlcommandE_MST(sqlth + " WHERE MAID='" + MAID + "'");
                    IFExecution_SUCCESS = true;
                }
            }
   
            else
            {
                SQlcommandE_DET(sqlo, dt);
                SQlcommandE_MST(sqlth + " WHERE MAID='" + MAID + "'");
                IFExecution_SUCCESS = true;
            }
        }
        #endregion
    
        #region SQlcommandE_DET
        protected void SQlcommandE_DET(string sql,DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
         
            basec.getcoms("DELETE MACHINING_DET WHERE MAID='"+MAID+"'");
            foreach (DataRow dr in dt.Rows)
            {
              
                SqlConnection sqlcon = bc.getcon();
                sqlcon.Open();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                MAKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM MACHINING_DET", "MAKEY", "MA");
                sqlcom.Parameters.Add("@MAKEY", SqlDbType.VarChar, 20).Value = MAKEY;
                sqlcom.Parameters.Add("@MAID", SqlDbType.VarChar, 20).Value = MAID;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = dr["项次"].ToString();
                sqlcom.Parameters.Add("@LAMINATING_PROCESS", SqlDbType.VarChar, 20).Value = dr["裱纸"].ToString();
                if (!string.IsNullOrEmpty(dr["裱纸含税价"].ToString()))
                {
                    sqlcom.Parameters.Add("@LAMINATING_PROCESS_PRICE", SqlDbType.VarChar, 20).Value = dr["裱纸含税价"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("@LAMINATING_PROCESS_PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;
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
            sqlcom.Parameters.Add("@MAID", SqlDbType.VarChar, 20).Value = MAID;
            sqlcom.Parameters.Add("@SIZE", SqlDbType.VarChar, 20).Value = SIZE;
            sqlcom.Parameters.Add("@MACHINE_TYPE", SqlDbType.VarChar, 20).Value = MACHINE_TYPE;
            if (!string.IsNullOrEmpty(DIE_CUTTING))
            {
                sqlcom.Parameters.Add("@DIE_CUTTING", SqlDbType.VarChar, 20).Value = DIE_CUTTING;
            }
            else
            {
                sqlcom.Parameters.Add("@DIE_CUTTING", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            sqlcom.Parameters.Add("@TAX_RATE", SqlDbType.VarChar, 20).Value = TAX_RATE;
            if (!string.IsNullOrEmpty(MACHINE_FREE))
            {
                sqlcom.Parameters.Add("@MACHINE_FREE", SqlDbType.VarChar, 20).Value = MACHINE_FREE;
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
