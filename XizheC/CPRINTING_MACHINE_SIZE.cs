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
    public class CPRINTING_MACHINE_SIZE:IGETID 
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
        private string _MACHINE_TYPE;
        public string MACHINE_TYPE
        {
            set { _MACHINE_TYPE = value; }
            get { return _MACHINE_TYPE; }
        }
        private string _PMID;
        public string PMID
        {
            set { _PMID = value; }
            get { return _PMID; }
        }
        private string _SN;
        public string SN
        {
            set { _SN = value; }
            get { return _SN; }
        }
        private string _MAX_WIDTH;
        public string MAX_WIDTH
        {
            set { _MAX_WIDTH = value; }
            get { return _MAX_WIDTH; }
        }
        private string _MAX_LENGTH;
        public string MAX_LENGTH
        {
            set { _MAX_LENGTH = value; }
            get { return _MAX_LENGTH; }
        }
        private string _MIN_WIDTH;
        public string MIN_WIDTH
        {
            set { _MIN_WIDTH = value; }
            get { return _MIN_WIDTH; }

        }
        private string _MIN_LENGTH;
        public string MIN_LENGTH
        {
            set { _MIN_LENGTH = value; }
            get { return _MIN_LENGTH; }
        }
        private string _ErrowInfo;
        public string ErrowInfo
        {
            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }
        }
        private string _PRINTING_PAPER;
        public string PRINTING_PAPER
        {
            set { _PRINTING_PAPER = value; }
            get { return _PRINTING_PAPER; }
        }
        #endregion
        #region sql
        string setsql = @"
SELECT 
A.MACHINE_TYPE AS 机器型号,
A.MAX_WIDTH AS 最大宽,
A.MAX_LENGTH AS 最大长,
A.MIN_WIDTH AS 最小宽,
A.MIN_LENGTH AS 最小长,
A.PRINTING_PAPER AS 印刷用纸,
B.ENAME AS 制单人,
A.Date AS 制单日期
FROM PRINTING_MACHINE_SIZE A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINTING_MACHINE_SIZE
(
PMID,
MACHINE_TYPE,
MAX_WIDTH,
MAX_LENGTH,
MIN_WIDTH,
MIN_LENGTH,
PRINTING_PAPER,
MakerID,
Date,
Year,
Month
)
VALUES
(
@PMID,
@MACHINE_TYPE,
@MAX_WIDTH,
@MAX_LENGTH,
@MIN_WIDTH,
@MIN_LENGTH,
@PRINTING_PAPER,
@MakerID,
@Date,
@Year,
@Month
)
";
        string setsqlth = @"
UPDATE PRINTING_MACHINE_SIZE SET 
MACHINE_TYPE=@MACHINE_TYPE,
MAX_WIDTH=@MAX_WIDTH,
MAX_LENGTH=@MAX_LENGTH,
MIN_WIDTH=@MIN_WIDTH,
MIN_LENGTH=@MIN_LENGTH,
PRINTING_PAPER=@PRINTING_PAPER,
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
      
        public CPRINTING_MACHINE_SIZE()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM PRINTING_MACHINE_SIZE", "PMID", "PM");
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
            string GET_MACHINE_TYPE = bc.getOnlyString("SELECT MACHINE_TYPE FROM PRINTING_MACHINE_SIZE WHERE  PMID='" + PMID + "'");

            if (!bc.exists("SELECT PMID FROM PRINTING_MACHINE_SIZE WHERE PMID='" + PMID + "'"))
            {
                if (bc.exists("SELECT * FROM PRINTING_MACHINE_SIZE where MACHINE_TYPE='" + MACHINE_TYPE + "'"))
                {
                    ErrowInfo = string.Format("印刷机型：{0}" + " 已经存在系统", MACHINE_TYPE);
                    IFExecution_SUCCESS = false;
                    //MessageBox.Show(PRINTING_MACHINE_SIZE);
                }

                else
                {
                    SQlcommandE(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_MACHINE_TYPE != MACHINE_TYPE)
            {
                if (bc.exists("SELECT * FROM PRINTING_MACHINE_SIZE where MACHINE_TYPE='" + MACHINE_TYPE + "'"))
                {

                    ErrowInfo = string.Format("印刷机型：{0}" + " 已经存在系统", MACHINE_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE(sqlth + " WHERE PMID='" + PMID + "'");
                    IFExecution_SUCCESS = true;
                }
            }

            else
            {

                SQlcommandE(sqlth + " WHERE PMID='" + PMID + "'");
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
            sqlcom.Parameters.Add("PMID", SqlDbType.VarChar, 20).Value = PMID;
            sqlcom.Parameters.Add("MACHINE_TYPE", SqlDbType.VarChar, 20).Value = MACHINE_TYPE;
            sqlcom.Parameters.Add("MAX_WIDTH", SqlDbType.VarChar, 20).Value = MAX_WIDTH;
            sqlcom.Parameters.Add("MAX_LENGTH", SqlDbType.VarChar, 20).Value = MAX_LENGTH;
            sqlcom.Parameters.Add("MIN_WIDTH", SqlDbType.VarChar, 20).Value = MIN_WIDTH;
            sqlcom.Parameters.Add("MIN_LENGTH", SqlDbType.VarChar, 20).Value = MIN_LENGTH;
            sqlcom.Parameters.Add("PRINTING_PAPER", SqlDbType.VarChar, 20).Value = PRINTING_PAPER;
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
