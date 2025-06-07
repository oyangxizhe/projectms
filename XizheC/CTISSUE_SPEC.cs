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
    public class CTISSUE_SPEC
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
        private string _SQUARE_PRICE;
        public string SQUARE_PRICE
        {
            set { _SQUARE_PRICE = value; }
            get { return _SQUARE_PRICE; }

        }
        private string _TON_PRICE;
        public string TON_PRICE
        {
            set { _TON_PRICE = value; }
            get { return _TON_PRICE; }

        }
        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

        }
        private string _WEIGHT;
        public string WEIGHT
        {
            set { _WEIGHT = value; }
            get { return _WEIGHT; }

        }
        private string _TSID;
        public string TSID
        {
            set { _TSID = value; }
            get { return _TSID; }

        }
        private string _TISSUE_SPEC;
        public string TISSUE_SPEC
        {
            set { _TISSUE_SPEC = value; }
            get { return _TISSUE_SPEC; }
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
        private string _TSKEY;
        public string TSKEY
        {
            set { _TSKEY = value; }
            get { return _TSKEY; }

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
      
   
        #endregion
        DataTable dt = new DataTable();
        #region sql
        string setsql = @"
SELECT 
B.TISSUE_SPEC AS 印刷,
A.SN AS 项次,
A.WEIGHT AS 克重,
A.TON_PRICE AS 吨价,
CASE WHEN  A.TON_PRICE IS NOT NULL THEN 
RTRIM(CONVERT(DECIMAL(18,2),A.TON_PRICE+(A.TON_PRICE*B.TAX_RATE/100))) 
ELSE ''
END  AS 含税吨价,
B.TAX_RATE AS 税率,
B.CUSTOMER_TYPE AS 客户类别,
A.REMARK AS 说明,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 制单人,
B.DATE AS 制单日期
FROM TISSUE_SPEC_DET A 
LEFT JOIN TISSUE_SPEC_MST B ON A.TSID=B.TSID

";


        string setsqlo = @"
INSERT INTO TISSUE_SPEC_DET
(
TSKEY,
TSID,
SN,
WEIGHT,
TON_PRICE,
REMARK,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
)
VALUES
(
@TSKEY,
@TSID,
@SN,
@WEIGHT,
@TON_PRICE,
@REMARK,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY

)


";

        string setsqlt = @"

INSERT INTO TISSUE_SPEC_MST
(
TSID,
TISSUE_SPEC,
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
@TSID,
@TISSUE_SPEC,
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
UPDATE TISSUE_SPEC_MST SET 
TISSUE_SPEC=@TISSUE_SPEC,
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
        public CTISSUE_SPEC()
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
            dt.Columns.Add("克重", typeof(string));
            dt.Columns.Add("含税吨价", typeof(string));
            dt.Columns.Add("说明", typeof(string));
            return dt;
        }
 
        #endregion
        #region GetTableInfo_t
        public DataTable GetTableInfo_t()
        {
            dt = new DataTable();
            dt.Columns.Add("印刷", typeof(string));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("克重", typeof(string));
            dt.Columns.Add("平方价", typeof(decimal));
            dt.Columns.Add("含税平方价", typeof(decimal));
            dt.Columns.Add("吨价", typeof(decimal));
            dt.Columns.Add("含税吨价", typeof(decimal));
            dt.Columns.Add("税率", typeof(decimal));
            dt.Columns.Add("客户类别", typeof(string));
            dt.Columns.Add("说明", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            dt.Columns.Add("制单日期", typeof(string));
            return dt;
        }

        #endregion
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from TISSUE_SPEC_MST", "TSID", "TS");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
              
            }
            return GETID;
        }
        public DataTable RETURN_DT(DataTable dtx)
        {
            if (dtx.Rows.Count > 0)
            {
                dt = GetTableInfo_t();
                foreach (DataRow dr1 in dtx.Rows)
                {
                    decimal d1 = 0, d2 = 0,d3=0;
                    DataRow dr = dt.NewRow();
                    dr["项次"] = dr1["项次"].ToString();
                    dr["印刷"] = dr1["印刷"].ToString();
                    dr["克重"] = dr1["克重"].ToString();
                 
                    if (!string.IsNullOrEmpty(dr1["吨价"].ToString()))
                    {
                        dr["吨价"] = dr1["吨价"].ToString();
                    }
                    else
                    {
                        dr["吨价"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dr1["吨价"].ToString()))
                    {
                        dr["含税吨价"] = (decimal.Parse(dr1["吨价"].ToString()) + decimal.Parse(dr1["吨价"].ToString()) * 
                            decimal.Parse(dr1["税率"].ToString()) / 100).ToString ("0.00");
                    }
                    else
                    {
                        dr["含税吨价"] = DBNull.Value;
                    }
                    if (!string.IsNullOrEmpty(dr1["税率"].ToString()))
                    {
                        dr["税率"] = dr1["税率"].ToString();
                    }
                    else
                    {
                        dr["税率"] = DBNull.Value;
                    }
                    dr["客户类别"] = dr1["客户类别"].ToString();
                    dr["说明"] = dr1["说明"].ToString();
                    dr["制单人"] = dr1["制单人"].ToString();
                    dr["制单日期"] = dr1["制单日期"].ToString();
                    if (!string.IsNullOrEmpty(dr["吨价"].ToString()) && !string.IsNullOrEmpty(dr["克重"].ToString()))
                    {
                        d1 = decimal.Parse(dr["吨价"].ToString());
                        d2 = decimal.Parse(dr["克重"].ToString());
                        dr["平方价"] = (d1 * d2 / 1000000).ToString("0.00");
                    }
                    else
                    {
                        dr["平方价"] = DBNull.Value;

                    }
                    if (!string.IsNullOrEmpty(dr["吨价"].ToString()) && !string.IsNullOrEmpty(dr["克重"].ToString()))
                    {
                        d1 = decimal.Parse(dr["吨价"].ToString());
                        d2 = decimal.Parse(dr["克重"].ToString());
                        d3 = decimal.Parse(dr["税率"].ToString());
                        dr["含税平方价"] = ((d1 + d1 * d3 / 100) * d2 / 1000000).ToString("0.00");
                    }
                    else
                    {
                        dr["含税平方价"] = DBNull.Value;

                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        #region save
        public void save(DataTable dt)
        {

            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string GET_TISSUE_SPEC = bc.getOnlyString("SELECT TISSUE_SPEC FROM TISSUE_SPEC_MST WHERE TSID='"+TSID +"'");
            string GET_CUSTOMER_TYPE = bc.getOnlyString("SELECT CUSTOMER_TYPE FROM TISSUE_SPEC_MST WHERE TSID='" + TSID + "'");
            if (!bc.exists("SELECT TSID FROM TISSUE_SPEC_DET WHERE TSID='" + TSID + "'"))
            {
                if (bc.exists("SELECT * FROM  TISSUE_SPEC_MST where TISSUE_SPEC='"+TISSUE_SPEC +"' AND CUSTOMER_TYPE='"+CUSTOMER_TYPE +"'"))
                {

                    ErrowInfo = string.Format("印刷：{0} + 客户类别：{1} "  + "组合已经存在系统",TISSUE_SPEC,CUSTOMER_TYPE  );
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    ACTION_DET(dt);
                    SQlcommandE_MST(sqlt);
                    IFExecution_SUCCESS = true;
                }

            }
            else if (TISSUE_SPEC != GET_TISSUE_SPEC || CUSTOMER_TYPE !=GET_CUSTOMER_TYPE )
            {

                if (bc.exists("SELECT * FROM  TISSUE_SPEC_MST where TISSUE_SPEC='" + TISSUE_SPEC + "' AND CUSTOMER_TYPE='" + CUSTOMER_TYPE + "'"))
                {

                    ErrowInfo = string.Format("印刷：{0} + 客户类别：{1} " + "组合已经存在系统", TISSUE_SPEC, CUSTOMER_TYPE);
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    ACTION_DET(dt);
                    SQlcommandE_MST(sqlth + " WHERE TSID='" + TSID + "'");
                    IFExecution_SUCCESS = true;
                }

            }
            else
            {

              ACTION_DET(dt);
              SQlcommandE_MST(sqlth + " WHERE TSID='" + TSID + "'");
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
            sqlcom.Parameters.Add("@TSKEY", SqlDbType.VarChar, 20).Value = TSKEY;
            sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = SN;
            sqlcom.Parameters.Add("@TSID", SqlDbType.VarChar, 20).Value = TSID;
            sqlcom.Parameters.Add("@WEIGHT", SqlDbType.VarChar, 20).Value = WEIGHT;
            if (!string.IsNullOrEmpty (TON_PRICE ))
            {
                sqlcom.Parameters.Add("@TON_PRICE", SqlDbType.VarChar, 20).Value = TON_PRICE;
            }
            else
            {
                sqlcom.Parameters.Add("@TON_PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }
            sqlcom.Parameters.Add("@REMARK", SqlDbType.VarChar, 20).Value = REMARK;
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
            sqlcom.Parameters.Add("@TSID", SqlDbType.VarChar, 20).Value = TSID;
            sqlcom.Parameters.Add("@TISSUE_SPEC", SqlDbType.VarChar, 20).Value = TISSUE_SPEC;
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
        private void ACTION_DET(DataTable dt)
        {
           
            basec.getcoms("DELETE TISSUE_SPEC_DET WHERE TSID='" + TSID + "'");
            foreach (DataRow dr in dt.Rows)
            {
                TSKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM TISSUE_SPEC_DET", "TSKEY", "TS");
                WEIGHT = dr["克重"].ToString();
                if (!string.IsNullOrEmpty(dr["含税吨价"].ToString()))
                {
                    TON_PRICE = (decimal.Parse(dr["含税吨价"].ToString()) / (1 + decimal.Parse(TAX_RATE) / 100)).ToString();
                }
                else
                {
                    TON_PRICE = "";
                }
               
                REMARK = dr["说明"].ToString();
                SN = dr["项次"].ToString();
                SQlcommandE_DET(sqlo);
            }
        }
     
    
    }
}
