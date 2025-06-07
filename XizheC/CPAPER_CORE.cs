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
    public class CPAPER_CORE
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
        private decimal _TAX_RATE;
        public decimal TAX_RATE
        {
            set { _TAX_RATE = value; }
            get { return _TAX_RATE; }
        }
        private string _RMKEY;
        public string RMKEY
        {
            set { _RMKEY = value; }
            get { return _RMKEY; }

        }

        private string _SPEC;
        public string SPEC
        {
            set { _SPEC = value; }
            get { return _SPEC; }

        }
        private string _PHONE;
        public string PHONE
        {
            set { _PHONE = value; }
            get { return _PHONE; }

        }
 
     
       
        private string _PCID;
        public string PCID
        {
            set { _PCID = value; }
            get { return _PCID; }

        }
      
        private string _PAPER_CORE;
        public string PAPER_CORE
        {
            set { _PAPER_CORE = value; }
            get { return _PAPER_CORE; }

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
        private string _PCKEY;
        public string PCKEY
        {
            set { _PCKEY = value; }
            get { return _PCKEY; }

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
B.PAPER_CORE AS 芯纸,
A.SN AS 项次,
A.SPEC AS 规格,
CASE WHEN A.PRICE IS NOT NULL THEN RTRIM(CONVERT(DECIMAL(18,3),A.PRICE/(1+B.TAX_RATE/100))) 
ELSE NULL
END AS 单价,
B.TAX_RATE AS 税率,
A.PRICE AS 含税单价,
A.UNIT AS 单位,
A.PAPER_CORE_DOOR AS 芯纸门幅,
B.CUSTOMER_TYPE AS 客户类别,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 制单人,
B.DATE AS 制单日期,
A.REMARK AS 备注
FROM PAPER_CORE_DET A 
LEFT JOIN PAPER_CORE_MST B ON A.PCID=B.PCID

";


        string setsqlo = @"
INSERT INTO PAPER_CORE_DET
(
PCKEY,
PCID,
SN,
SPEC,
PRICE,
UNIT,
PAPER_CORE_DOOR,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
)
VALUES
(
@PCKEY,
@PCID,
@SN,
@SPEC,
@PRICE,
@UNIT,
@PAPER_CORE_DOOR,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY

)


";

        string setsqlt = @"

INSERT INTO PAPER_CORE_MST
(
PCID,
PAPER_CORE,
TAX_RATE,
CUSTOMER_TYPE,
DATE,
MAKERID,
YEAR,
MONTH,
DAY
)
VALUES
(
@PCID,
@PAPER_CORE,
@TAX_RATE,
@CUSTOMER_TYPE,
@DATE,
@MAKERID,
@YEAR,
@MONTH,
@DAY
)
";
        string setsqlth = @"
UPDATE PAPER_CORE_MST SET 
PAPER_CORE=@PAPER_CORE,
TAX_RATE=@TAX_RATE,
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
        public CPAPER_CORE()
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
            dt.Columns.Add("规格", typeof(string));
            dt.Columns.Add("含税单价", typeof(string));
            dt.Columns.Add("单位", typeof(string));
            dt.Columns.Add("芯纸门幅", typeof(string));
            return dt;
        }
 
        #endregion
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from PAPER_CORE_MST", "PCID", "PC");
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
            string GET_PAPER_CORE = bc.getOnlyString("SELECT PAPER_CORE FROM PAPER_CORE_MST WHERE PCID='"+PCID +"'");
            string GET_CUSTOMER_TYPE = bc.getOnlyString("SELECT CUSTOMER_TYPE FROM PAPER_CORE_MST WHERE PCID='" + PCID + "'");
            if (!bc.exists("SELECT PCID FROM PAPER_CORE_DET WHERE PCID='" + PCID + "'"))
            {
                if (bc.exists("SELECT * FROM  PAPER_CORE_MST where PAPER_CORE='" + PAPER_CORE + "' AND CUSTOMER_TYPE='"+CUSTOMER_TYPE +"'"))
                {

                    ErrowInfo = string.Format("芯纸：{0}"  + " + 客户类别：{1} 组合已经存在系统",PAPER_CORE,CUSTOMER_TYPE  );
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    SQlcommandE_DET(sqlo,dt );
                    SQlcommandE_MST(sqlt);
                    IFExecution_SUCCESS = true;
                }

            }
            else if (PAPER_CORE != GET_PAPER_CORE  || CUSTOMER_TYPE !=GET_CUSTOMER_TYPE  )
            {

              if (bc.exists("SELECT * FROM  PAPER_CORE_MST where PAPER_CORE='" + PAPER_CORE + "' AND CUSTOMER_TYPE='"+CUSTOMER_TYPE +"'"))
                {

                    ErrowInfo = string.Format("芯纸：{0}"  + " + 客户类别：{1} 组合已经存在系统",PAPER_CORE,CUSTOMER_TYPE  );
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    SQlcommandE_DET(sqlo,dt );
                    SQlcommandE_MST(sqlth + " WHERE PCID='" + PCID + "'");
                    IFExecution_SUCCESS = true;
                }

            }
            else
            {

              SQlcommandE_DET(sqlo, dt);
              SQlcommandE_MST(sqlth + " WHERE PCID='" + PCID + "'");
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
            SqlConnection sqlcon = bc.getcon();
            sqlcon.Open();
            basec.getcoms("DELETE PAPER_CORE_DET WHERE PCID='" + PCID + "'");
            foreach (DataRow dr in dt.Rows)
            {
                PCKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM PAPER_CORE_DET", "PCKEY", "PC");
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcom.Parameters.Add("@PCKEY", SqlDbType.VarChar, 20).Value = PCKEY;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value =dr["项次"].ToString();
                sqlcom.Parameters.Add("@PCID", SqlDbType.VarChar, 20).Value = PCID;
                sqlcom.Parameters.Add("@SPEC", SqlDbType.VarChar, 20).Value = dr["规格"].ToString();
                if (dr["含税单价"].ToString()!="")
                {
                    sqlcom.Parameters.Add("@PRICE", SqlDbType.VarChar, 20).Value = dr["含税单价"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("@PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
              
                sqlcom.Parameters.Add("@UNIT", SqlDbType.VarChar, 20).Value = dr["单位"].ToString();
                sqlcom.Parameters.Add("@PAPER_CORE_DOOR", SqlDbType.VarChar, 20).Value = dr["芯纸门幅"].ToString();
                sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
                sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
            }
           
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
            sqlcom.Parameters.Add("@PCID", SqlDbType.VarChar, 20).Value = PCID;
            sqlcom.Parameters.Add("@PAPER_CORE", SqlDbType.VarChar, 20).Value = PAPER_CORE;
            sqlcom.Parameters.Add("@TAX_RATE", SqlDbType.VarChar, 20).Value = TAX_RATE;
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
