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
    public class CPARTS_AUXILIARY:IGETID 
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

        private string _PARTS_AUXILIARY;
        public string PARTS_AUXILIARY
        {
            set { _PARTS_AUXILIARY = value; }
            get { return _PARTS_AUXILIARY; }

        }
        private string _PAID;
        public string PAID
        {
            set { _PAID = value; }
            get { return _PAID; }

        }
        private string _TAX_UNIT_PRICE;
        public string TAX_UNIT_PRICE
        {
            set { _TAX_UNIT_PRICE = value; }
            get { return _TAX_UNIT_PRICE; }

        }
        private string _UNIT;
        public string UNIT
        {
            set { _UNIT = value; }
            get { return _UNIT; }

        }
        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

        }
        private string _TAX_MACHINE_COST;
        public string TAX_MACHINE_COST
        {
            set { _TAX_MACHINE_COST = value; }
            get { return _TAX_MACHINE_COST; }

        }
        private string _TAX_RATE;
        public string TAX_RATE
        {
            set { _TAX_RATE = value; }
            get { return _TAX_RATE; }

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
A.PARTS_AUXILIARY AS 配件名,
RTRIM(CONVERT(DECIMAL(18,2),A.TAX_UNIT_PRICE/(1+A.TAX_RATE/100))) AS 未税单价,
A.TAX_UNIT_PRICE AS 含税单价,
A.UNIT AS 单位,
RTRIM(CONVERT(DECIMAL(18,1),A.TAX_RATE))+'%' AS 税率,
A.REMARK AS 说明,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM PARTS_AUXILIARY A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PARTS_AUXILIARY
(
PAID,
PARTS_AUXILIARY,
TAX_UNIT_PRICE,
UNIT,
TAX_RATE,
REMARK,
MakerID,
Date,
Year,
Month
)
VALUES
(
@PAID,
@PARTS_AUXILIARY,
@TAX_UNIT_PRICE,
@UNIT,
@TAX_RATE,
@REMARK,
@MakerID,
@Date,
@Year,
@Month
)
";
        string setsqlth = @"
UPDATE PARTS_AUXILIARY SET 
PAID=@PAID,
PARTS_AUXILIARY=@PARTS_AUXILIARY,
TAX_UNIT_PRICE=@TAX_UNIT_PRICE,
UNIT=@UNIT,
TAX_RATE=@TAX_RATE,
REMARK=@REMARK,
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
      
        public CPARTS_AUXILIARY()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM PARTS_AUXILIARY", "PAID", "PA");
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
            string GET_PARTS_AUXILIARY = bc.getOnlyString("SELECT PARTS_AUXILIARY FROM PARTS_AUXILIARY WHERE  PAID='" + PAID + "'");

            if (!bc.exists("SELECT PAID FROM PARTS_AUXILIARY WHERE PAID='" + PAID + "'"))
            {
                if (bc.exists("SELECT * FROM PARTS_AUXILIARY where PARTS_AUXILIARY='" + PARTS_AUXILIARY+ "'"))
                {
                    ErrowInfo = string.Format("配件名：{0}" + " 已经存在系统", PARTS_AUXILIARY);
                    IFExecution_SUCCESS = false;
              
                }

                else
                {
                    SQlcommandE(sqlt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_PARTS_AUXILIARY != PARTS_AUXILIARY)
            {
                if (bc.exists("SELECT * FROM PARTS_AUXILIARY where PARTS_AUXILIARY='" + PARTS_AUXILIARY + "'"))
                {

                    ErrowInfo = string.Format("配件名：{0}" + " 已经存在系统", PARTS_AUXILIARY);
                    IFExecution_SUCCESS = false;
                }
                else
                {

                    SQlcommandE(sqlth + " WHERE PAID='" + PAID + "'");
                    IFExecution_SUCCESS = true;
                }
            }

            else
            {

                SQlcommandE(sqlth + " WHERE PAID='" + PAID + "'");
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
            sqlcom.Parameters.Add("PAID", SqlDbType.VarChar, 20).Value = PAID;
            sqlcom.Parameters.Add("PARTS_AUXILIARY", SqlDbType.VarChar, 20).Value = PARTS_AUXILIARY;
            sqlcom.Parameters.Add("TAX_UNIT_PRICE", SqlDbType.VarChar, 20).Value = TAX_UNIT_PRICE;
            sqlcom.Parameters.Add("UNIT", SqlDbType.VarChar, 20).Value = UNIT;
            sqlcom.Parameters.Add("TAX_RATE", SqlDbType.VarChar, 20).Value = TAX_RATE;
            sqlcom.Parameters.Add("REMARK", SqlDbType.VarChar, 1000).Value = REMARK;
            sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region emptydatatable_T
        public DataTable emptydatatable_T()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("配件名", typeof(string));
            dt.Columns.Add("未税单价", typeof(string));
            dt.Columns.Add("单位", typeof(string));
            dt.Columns.Add("税率", typeof(string));
            dt.Columns.Add("含税单价", typeof(string));
            dt.Columns.Add("说明", typeof(string));
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
                dr["序号"] = i.ToString();
                dr["配件名"] = dr1["配件名"].ToString();
                dr["未税单价"] = dr1["未税单价"].ToString();
                dr["单位"] = dr1["单位"].ToString();
                dr["税率"] = dr1["税率"].ToString();
                dr["含税单价"] = dr1["含税单价"].ToString();
                dr["说明"] = dr1["说明"].ToString();
                dr["制单人"] = dr1["制单人"].ToString();
                dr["制单日期"] = dr1["制单日期"].ToString();
                dt.Rows.Add(dr);
                i = i + 1;
            }
            return dt;
        }
        #endregion
    
    }
}
