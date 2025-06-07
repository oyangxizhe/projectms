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
    public class CDOOR_PARAMETERS
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
        private string _RMKEY;
        public string RMKEY
        {
            set { _RMKEY = value; }
            get { return _RMKEY; }

        }

        private string _PRICE;
        public string PRICE
        {
            set { _PRICE = value; }
            get { return _PRICE; }

        }
        private string _PHONE;
        public string PHONE
        {
            set { _PHONE = value; }
            get { return _PHONE; }

        }
 
     
       
        private string _DPID;
        public string DPID
        {
            set { _DPID = value; }
            get { return _DPID; }

        }
      
        private string _DOOR_PARAMETERS;
        public string DOOR_PARAMETERS
        {
            set { _DOOR_PARAMETERS = value; }
            get { return _DOOR_PARAMETERS; }

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
        private string _DPKEY;
        public string DPKEY
        {
            set { _DPKEY = value; }
            get { return _DPKEY; }

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

        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

        }
        #endregion
       
        #region sql
        string setsql = @"
SELECT 
B.DOOR_PARAMETERS AS 印刷用纸或芯纸,
A.SN AS 项次,
A.PRICE AS 值,
A.CUSTOMER_TYPE AS 客户类别,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 制单人,
B.DATE AS 制单日期,
A.REMARK AS 备注
FROM DOOR_PARAMETERS_DET A 
LEFT JOIN DOOR_PARAMETERS_MST B ON A.DPID=B.DPID

";


        string setsqlo = @"
INSERT INTO DOOR_PARAMETERS_DET
(
DPKEY,
DPID,
SN,
PRICE,
CUSTOMER_TYPE,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
)
VALUES
(
@DPKEY,
@DPID,
@SN,
@PRICE,
@CUSTOMER_TYPE,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY

)


";

        string setsqlt = @"

INSERT INTO DOOR_PARAMETERS_MST
(
DPID,
DOOR_PARAMETERS,
DATE,
MAKERID,
YEAR,
MONTH,
DAY
)
VALUES
(
@DPID,
@DOOR_PARAMETERS,
@DATE,
@MAKERID,
@YEAR,
@MONTH,
@DAY
)
";
        string setsqlth = @"
UPDATE DOOR_PARAMETERS_MST SET 
DOOR_PARAMETERS=@DOOR_PARAMETERS,
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
        public CDOOR_PARAMETERS()
        {
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
            DataTable dt = new DataTable(); 
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("值", typeof(string));
            dt.Columns.Add("客户类别", typeof(string));
            return dt;
        }
 
        #endregion
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from DOOR_PARAMETERS_MST", "DPID", "DP");
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
            string GET_DOOR_PARAMETERS = bc.getOnlyString("SELECT DOOR_PARAMETERS FROM DOOR_PARAMETERS_MST WHERE DPID='"+DPID +"'");
            if (!bc.exists("SELECT DPID FROM DOOR_PARAMETERS_DET WHERE DPID='" + DPID + "'"))
            {
                if (bc.exists("SELECT * FROM  DOOR_PARAMETERS_MST where DOOR_PARAMETERS='"+DOOR_PARAMETERS +"'"))
                {

                    ErrowInfo = string.Format("印刷用纸或芯纸：{0}"  + "已经存在系统",DOOR_PARAMETERS );
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    ACTION_DET(dt);
                    SQlcommandE_MST(sqlt);
                    IFExecution_SUCCESS = true;
                }

            }
            else if (DOOR_PARAMETERS != GET_DOOR_PARAMETERS)
            {

                if (bc.exists("SELECT * FROM  DOOR_PARAMETERS_MST where DOOR_PARAMETERS='" + DOOR_PARAMETERS + "'"))
                {

                    ErrowInfo = string.Format("印刷用纸或芯纸：{0}" + "已经存在系统", DOOR_PARAMETERS);
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    ACTION_DET(dt);
                    SQlcommandE_MST(sqlth + " WHERE DPID='" + DPID + "'");
                    IFExecution_SUCCESS = true;
                }

            }
            else
            {

              ACTION_DET(dt);
              SQlcommandE_MST(sqlth + " WHERE DPID='" + DPID + "'");
              IFExecution_SUCCESS = true;
            }
            
        }
        #endregion
    
        #region SQlcommandE_DET
        protected void SQlcommandE_DET(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
            SqlConnection sqlcon = bc.getcon();
            sqlcon.Open();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@DPKEY", SqlDbType.VarChar, 20).Value = DPKEY;
            sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = SN;
            sqlcom.Parameters.Add("@DPID", SqlDbType.VarChar, 20).Value = DPID;
            sqlcom.Parameters.Add("@PRICE", SqlDbType.VarChar, 20).Value = PRICE;
            sqlcom.Parameters.Add("@CUSTOMER_TYPE", SqlDbType.VarChar, 20).Value = CUSTOMER_TYPE;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
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
            sqlcom.Parameters.Add("@DPID", SqlDbType.VarChar, 20).Value = DPID;
            sqlcom.Parameters.Add("@DOOR_PARAMETERS", SqlDbType.VarChar, 20).Value = DOOR_PARAMETERS;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        private void ACTION_DET(DataTable dt)
        {
           
            basec.getcoms("DELETE DOOR_PARAMETERS_DET WHERE DPID='" + DPID + "'");
            foreach (DataRow dr in dt.Rows)
            {
                DPKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM DOOR_PARAMETERS_DET", "DPKEY", "DP");
                PRICE = dr["值"].ToString();
                CUSTOMER_TYPE = dr["客户类别"].ToString();
                SN = dr["项次"].ToString();
                SQlcommandE_DET(sqlo);
            }
        }

    
    }
}
