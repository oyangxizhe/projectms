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
    public class CPAPER_CORE_OPTION:IGETID 
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
        private string _DPID;
        public string DPID
        {
            set { _DPID = value; }
            get { return _DPID; }
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
A.PAPER_CORE AS 芯纸选项,
A.PAPER_CORE_A AS 芯纸内耗1到300,
A.PAPER_CORE_B AS 芯纸内耗大于300,
B.ENAME AS 制单人,
A.Date AS 制单日期
FROM PAPER_CORE_OPTION A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PAPER_CORE_OPTION
(
PCID,
PAPER_CORE,
PAPER_CORE_A,
PAPER_CORE_B,
MakerID,
Date,
YEAR,
MONTH
)
VALUES
(
@PCID,
@PAPER_CORE,
@PAPER_CORE_A,
@PAPER_CORE_B,
@MakerID,
@Date,
@YEAR,
@MONTH
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

        public CPAPER_CORE_OPTION()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM PAPER_CORE_OPTION", "PCID", "PC");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region emptydatatable_T
        public DataTable emptydatatable_T()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("芯纸选项", typeof(string));
            dt.Columns.Add("芯纸内耗1到300", typeof(string));
            dt.Columns.Add("芯纸内耗大于300", typeof(string));
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
                dr["项次"] = i.ToString();
                dr["芯纸选项"] = dr1["芯纸选项"].ToString();
                dr["芯纸内耗1到300"] = dr1["芯纸内耗1到300"].ToString();
                dr["芯纸内耗大于300"] = dr1["芯纸内耗大于300"].ToString();
                dr["制单人"] = dr1["制单人"].ToString();
                dr["制单日期"] = dr1["制单日期"].ToString();
                dt.Rows.Add(dr);
                i = i + 1;
            }
            return dt;
        }
        #endregion
        #region save
        public void save(DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            basec.getcoms("DELETE PAPER_CORE_OPTION");
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

            foreach (DataRow dr in dt.Rows)
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcon.Open();
                if (dr["芯纸选项"].ToString() != "")
                {
                    sqlcom.Parameters.Add("PCID", SqlDbType.VarChar, 20).Value = GETID();
                    sqlcom.Parameters.Add("PAPER_CORE", SqlDbType.VarChar, 20).Value = dr["芯纸选项"].ToString();
                    sqlcom.Parameters.Add("PAPER_CORE_A", SqlDbType.VarChar, 20).Value = dr["芯纸内耗1到300"].ToString();
                    sqlcom.Parameters.Add("PAPER_CORE_B", SqlDbType.VarChar, 20).Value = dr["芯纸内耗大于300"].ToString();
                    sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
                    sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
                    sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
                    sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
                    sqlcom.ExecuteNonQuery();
                }
                sqlcon.Close();
            }

        }
        #endregion
    
    }
}
