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
    public class CPROCESSING_TECHNOLOGY
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
  
        private string _PTID;
        public string PTID
        {
            set { _PTID = value; }
            get { return _PTID; }

        }
        private string _MATERIAL_TYPE;
        public string MATERIAL_TYPE
        {
            set { _MATERIAL_TYPE = value; }
            get { return _MATERIAL_TYPE; }
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
B.MATERIAL_TYPE AS 加工内容,
A.SN AS 项次,
A.TECHNOLOGY AS 工艺
FROM PROCESSING_TECHNOLOGY_DET A 
LEFT JOIN PROCESSING_TECHNOLOGY_MST B ON A.PTID=B.PTID


";


        string setsqlo = @"
INSERT INTO PROCESSING_TECHNOLOGY_DET
(
PTKEY,
PTID,
SN,
TECHNOLOGY,
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
@TECHNOLOGY,
@MakerID,
@Date,
@YEAR,
@MONTH,
@DAY
)


";

        string setsqlt = @"

INSERT INTO PROCESSING_TECHNOLOGY_MST
(
PTID,
MATERIAL_TYPE,
DATE,
MAKERID,
YEAR,
MONTH,
DAY
)
VALUES
(
@PTID,
@MATERIAL_TYPE,
@DATE,
@MAKERID,
@YEAR,
@MONTH,
@DAY
)
";
        string setsqlth = @"
UPDATE PROCESSING_TECHNOLOGY_MST SET 
MATERIAL_TYPE=@MATERIAL_TYPE,
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
        public CPROCESSING_TECHNOLOGY()
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
            dt.Columns.Add("工艺", typeof(string));
            return dt;
        }
 
        #endregion
   
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from PROCESSING_TECHNOLOGY_MST", "PTID", "PT");
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
            string GET_MATERIAL_TYPE = bc.getOnlyString("SELECT MATERIAL_TYPE FROM PROCESSING_TECHNOLOGY_MST WHERE  PTID='" + PTID + "'");
  
            if (!bc.exists("SELECT PTID FROM PROCESSING_TECHNOLOGY_DET WHERE PTID='" + PTID + "'"))
            {
                if (bc.exists("SELECT * FROM PROCESSING_TECHNOLOGY_MST where MATERIAL_TYPE='" + MATERIAL_TYPE + "'"))
                {
                    ErrowInfo = string.Format("加工内容：{0}" + " 已经存在系统", MATERIAL_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE_DET(sqlo, dt);
                    SQlcommandE_MST(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_MATERIAL_TYPE != MATERIAL_TYPE )
            {
                if (bc.exists("SELECT * FROM PROCESSING_TECHNOLOGY_MST where MATERIAL_TYPE='" + MATERIAL_TYPE + "'"))
                {
                    ErrowInfo = string.Format("加工内容：{0}" + " 已经存在系统", MATERIAL_TYPE);
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
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
            basec.getcoms("DELETE PROCESSING_TECHNOLOGY_DET WHERE PTID='"+PTID+"'");
            foreach (DataRow dr in dt.Rows)
            {
                SqlConnection sqlcon = bc.getcon();
                sqlcon.Open();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                PTKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM PROCESSING_TECHNOLOGY_DET", "PTKEY", "PT");
                sqlcom.Parameters.Add("@PTKEY", SqlDbType.VarChar, 20).Value = PTKEY;
                sqlcom.Parameters.Add("@PTID", SqlDbType.VarChar, 20).Value = PTID;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = dr["项次"].ToString();
                sqlcom.Parameters.Add("@TECHNOLOGY", SqlDbType.VarChar, 20).Value = dr["工艺"].ToString();
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
            sqlcom.Parameters.Add("@MATERIAL_TYPE", SqlDbType.VarChar, 20).Value = MATERIAL_TYPE;
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
