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
    public class CPRINT_DIE_CUTTING:IGETID 
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

        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

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
A.PROJECT_NAME AS 项目,
A.DIE_CUTTING_LENGTH_METRE AS 刀模长米,
A.YUAN_METRE AS 元米,
A.ROUND_HOLE_COUNT AS 圆孔个数,
A.YUAN_PERSON AS 元个,
A.SUBTOTAL AS 小计
FROM PRINT_DIE_CUTTING A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID
LEFT JOIN PRINTING_OFFER_MST C ON A.PFID=C.PFID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_DIE_CUTTING
(
PDID,
PFID,
PROJECT_NAME,
DIE_CUTTING_LENGTH_METRE,
YUAN_METRE,
ROUND_HOLE_COUNT,
YUAN_PERSON,
SUBTOTAL,
MakerID,
Date,
Year,
Month,
day
)
VALUES
(
@PDID,
@PFID,
@PROJECT_NAME,
@DIE_CUTTING_LENGTH_METRE,
@YUAN_METRE,
@ROUND_HOLE_COUNT,
@YUAN_PERSON,
@SUBTOTAL,
@MakerID,
@Date,
@Year,
@Month,
@day

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
      
        public CPRINT_DIE_CUTTING()
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
            string v1 = bc.numYMD(12, 4, "0001", "SELECT * FROM PRINT_DIE_CUTTING", "PDID", "PD");
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
            basec.getcoms("DELETE PRINT_DIE_CUTTING WHERE PFID='" + PFID + "'");
            SQlcommandE(sqlt,dt);
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
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcon.Open();
                sqlcom.Parameters.Add("PDID", SqlDbType.VarChar, 20).Value = GETID();
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                sqlcom.Parameters.Add("PROJECT_NAME", SqlDbType.VarChar, 20).Value = dt.Rows[i]["项目"].ToString();
                if (i == 2)
                {
                    sqlcom.Parameters.Add("DIE_CUTTING_LENGTH_METRE", SqlDbType.VarChar, 20).Value = "";
                }
                else
                {
                    sqlcom.Parameters.Add("DIE_CUTTING_LENGTH_METRE", SqlDbType.VarChar, 20).Value = dt.Rows[i]["刀模长米"].ToString();
                }
                sqlcom.Parameters.Add("YUAN_METRE", SqlDbType.VarChar, 20).Value = dt.Rows[i]["元米"].ToString();
                if (!string.IsNullOrEmpty(dt.Rows[i]["圆孔个数"].ToString()))//圆孔个数在软件费用明细显示整数位，需要DECIMAL类型，不能换成STRING 15/12/27
                {
                    sqlcom.Parameters.Add("ROUND_HOLE_COUNT", SqlDbType.VarChar, 20).Value = dt.Rows[i]["圆孔个数"].ToString();
                }
                else
                {
                    //圆孔个数在软件费用明细显示整数位，需要DECIMAL类型，不能换成STRING 15/12/27
                    sqlcom.Parameters.Add("ROUND_HOLE_COUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
             
                sqlcom.Parameters.Add("YUAN_PERSON", SqlDbType.VarChar, 20).Value = dt.Rows[i]["元个"].ToString();
                sqlcom.Parameters.Add("SUBTOTAL", SqlDbType.VarChar, 20).Value = dt.Rows[i]["小计"].ToString();
                sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
                sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("day", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
            }
          
        }
        #endregion
    
    }
}
