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
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Collections.Generic;

namespace XizheC
{
    public class CNO_PAPER_OFFER
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _NPKEY;
        public string NPKEY
        {
            set { _NPKEY = value; }
            get { return _NPKEY; }
        }
        private string _WATER_CODE;
        public string WATER_CODE
        {
            set { _WATER_CODE = value; }
            get { return _WATER_CODE; }
        }
        private string _OFFER_DATE;
        public string OFFER_DATE
        {
            set { _OFFER_DATE = value; }
            get { return _OFFER_DATE; }

        }
        private string _SAMPLE_CODE_FIRST;
        public string SAMPLE_CODE_FIRST
        {
            set { _SAMPLE_CODE_FIRST = value; }
            get { return _SAMPLE_CODE_FIRST; }

        }
        private string _LOGIN_EMID;
        public string LOGIN_EMID
        {
            set { _LOGIN_EMID = value; }
            get { return _LOGIN_EMID; }

        }
        private string _SAMPLE_CODE;
        public string SAMPLE_CODE
        {
            set { _SAMPLE_CODE = value; }
            get { return _SAMPLE_CODE; }

        }
        private string _COUNT;
        public string COUNT
        {
            set { _COUNT = value; }
            get { return _COUNT; }

        }
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _OFFER_ID;
        public string OFFER_ID
        {
            set { _OFFER_ID = value; }
            get { return _OFFER_ID; }

        }
        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

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
        private string _NPID;
        public string NPID
        {
            set { _NPID = value; }
            get { return _NPID; }

        }
        private string _PROJECT_ID;
        public string PROJECT_ID
        {
            set { _PROJECT_ID = value; }
            get { return _PROJECT_ID; }

        }
        private string _AUDIT_STATUS;
        public string AUDIT_STATUS
        {
            set { _AUDIT_STATUS = value; }
            get { return _AUDIT_STATUS; }

        }
        private string _UNIT_PRICE;
        public string UNIT_PRICE
        {
            set { _UNIT_PRICE = value; }
            get { return _UNIT_PRICE; }

        }
        private string _PIID;
        public string PIID
        {
            set { _PIID = value; }
            get { return _PIID; }

        }
        private string _OFFER_ID_SENVEN;
        public string OFFER_ID_SENVEN
        {
            set { _OFFER_ID_SENVEN = value; }
            get { return _OFFER_ID_SENVEN; }

        }
        private string _OFFER_TYPE_CODE;
        public string OFFER_TYPE_CODE
        {
            set { _OFFER_TYPE_CODE = value; }
            get { return _OFFER_TYPE_CODE; }

        }

        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        #endregion
        DataTable dt = new DataTable();
        StringBuilder sqb = new StringBuilder();
        List<string> list2 = new List<string>();
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        int i;
        #region sql
        string setsql = @"

SELECT
D.NPKEY AS 编号,
A.PROJECT_ID AS 项目号,
D.COUNT AS 数量,
D.UNIT_PRICE AS 报出价,
D.OFFER_ID AS 报价编号,
B.PIID AS 项目ID,
B.PROJECT_NAME AS 项目名称,
CASE WHEN A.AUDIT_STATUS='Y' THEN '已审核'
ELSE '未审核'
END AS 审核状态,
C.CName  AS 客户名称,
B.BRAND AS 品牌,
A.OFFER_DATE AS 报价日期,
(SELECT EName  FROM EMPLOYEEINFO WHERE EMID=B.AE_MAKERID_1 ) AS AE,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.STRUCTURE_MAKERID_1) AS 结构设计,
(SELECT EName  FROM EMPLOYEEINFO WHERE EMID=B.PLANE_MAKERID_1) AS 平面设计,
D.UNIT_PRICE AS 报出价,
A.REMARK AS 备注,
A.MakerID AS 制单人工号,
A.Date AS 制单日期
FROM NO_PAPER_OFFER_MST A
LEFT JOIN PROJECT_INFO B ON A.PROJECT_ID=B.PROJECT_ID 
LEFT JOIN CustomerInfo_MST C ON C.CUID=B.CUID
LEFT JOIN NO_PAPER_OFFER_DET D ON A.NPID=D.NPID  



";


        string setsqlo = @"

INSERT INTO NO_PAPER_OFFER_DET
(
NPKEY,
NPID,
SN,
COUNT,
UNIT_PRICE,
OFFER_ID,
PIID,
MakerID,
Date,
YEAR,
MONTH,
DAY
)
VALUES
(
@NPKEY,
@NPID,
@SN,
@COUNT,
@UNIT_PRICE,
@OFFER_ID,
@PIID,
@MakerID,
@Date,
@YEAR,
@MONTH,
@DAY
)

";

        string setsqlt = @"

INSERT INTO NO_PAPER_OFFER_MST
(
NPID,
PROJECT_ID,
OFFER_DATE,
AUDIT_STATUS,
WATER_CODE,
REMARK,
MakerID,
Date,
YEAR,
MONTH
)
VALUES
(
@NPID,
@PROJECT_ID,
@OFFER_DATE,
@AUDIT_STATUS,
@WATER_CODE,
@REMARK,
@MakerID,
@Date,
@YEAR,
@MONTH

)
";
        string setsqlth = @"
UPDATE NO_PAPER_OFFER_MST SET 
NPID=@NPID,
PROJECT_ID=@PROJECT_ID,
OFFER_DATE=@OFFER_DATE,
AUDIT_STATUS=@AUDIT_STATUS,
WATER_CODE=@WATER_CODE,
REMARK=@REMARK,
MakerID=@MakerID,
Date=@Date,
YEAR=@YEAR,
MONTH=@MONTH

";

        string setsqlf = @"

";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        public CNO_PAPER_OFFER()
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

        #region GetTableInfo_SEARCH
        public DataTable GetTableInfo_SEARCH()
        {
            dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("项目号", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("报价编号", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("报出价", typeof(string));
            dt.Columns.Add("客户名称", typeof(string));
            dt.Columns.Add("品牌", typeof(string));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("平面设计", typeof(string));
            dt.Columns.Add("结构设计", typeof(string));
            dt.Columns.Add("备注", typeof(string));
            dt.Columns.Add("报价日期", typeof(string));
            return dt;
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("报出价", typeof(string));
            dt.Columns.Add("报价编号", typeof(string));
            return dt;
        }
        #endregion
        #region GETID
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from NO_PAPER_OFFER_NO", "NPID", "NP");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO NO_PAPER_OFFER_NO(NPID,DATE,YEAR,MONTH) VALUES ('" + v1 + "','" + varDate + "','" + year +
                    "','" + month + "')");
            }
            return GETID;
        }
        #endregion
        #region GETID_OFFER_ID
        public void GETID_OFFER_ID(string OFFER_TYPE_CODE)
        { 
 
            string vOFFER_TYPE_CODE = bc.getOnlyString("SELECT SUBSTRING(OFFER_ID,5,1) FROM NO_PAPER_OFFER_DET WHERE NPID='" + NPID  + "'");
            //1610Z002-03-ADM
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string v1 = "", v2 = "",v3="",v4="";
            SAMPLE_CODE = bc.getOnlyString("SELECT SAMPLE_CODE FROM EMPLOYEEINFO WHERE EMID='" + MAKERID + "'");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            /*add offer_type_code juage 161012  start         */
            if (!bc.exists("SELECT * FROM NO_PAPER_OFFER_ID_NO WHERE PIID='" + PIID +
"' AND YEAR='" + year + "' AND MONTH='" + month + "' AND SUBSTRING (OFFER_ID_SENVEN,5,1)='"+OFFER_TYPE_CODE+"'"))
            /*add offer_type_code juage 161012  end */
            {
                v2 = numYY(8, 3, "001", "select * from NO_PAPER_OFFER_ID_NO", "OFFER_ID_SENVEN", OFFER_TYPE_CODE);
                v1 = bc.numNOYMD(11, 2, "01", "select * from NO_PAPER_OFFER_DET WHERE SUBSTRING(OFFER_ID,1,8)='" + v2 + "'", "OFFER_ID",
                 v2 + "-");
                v1 = v1 + "-" + SAMPLE_CODE + SAMPLE_CODE_FIRST;
            }
            else
            {
                v2 = bc.getOnlyString(string.Format("SELECT OFFER_ID_SENVEN FROM NO_PAPER_OFFER_ID_NO WHERE PIID='{0}'  AND YEAR='" + year +
                    "' AND MONTH='" + month + "'  AND SUBSTRING (OFFER_ID_SENVEN,5,1)='" + OFFER_TYPE_CODE + "'", PIID));
                /*区分类别码 start 161010 1/2*/
                v3 = v2.Substring(0, 4);
                v4 = v2.Substring(5, 3);
                v2 = v2.Substring(0, 4) + OFFER_TYPE_CODE + v2.Substring(5, 3);//类别码要替换 161010
                /*区分类别码 end 161010 1/2*/
                //v1 = bc.numNOYMD(11, 2, "01", "select * from NO_PAPER_OFFER_DET WHERE SUBSTRING(OFFER_ID,1,8)='" + v2 + "'", "OFFER_ID",v2 + "-"); 之前的版本类别全是Z
                /*区分类别码 start 161010 2/2*/
                v1 = bc.numNOYMD(11, 2, "01", "select * from NO_PAPER_OFFER_DET WHERE SUBSTRING(OFFER_ID,1,4)='" + v3 + "' AND SUBSTRING(OFFER_ID,6,3)='" + v4 + "'", "OFFER_ID", v2 + "-");
                /*区分类别码 end 161010 2/2*/
                v1 = v1 + "-" + SAMPLE_CODE + SAMPLE_CODE_FIRST;
            }
            OFFER_ID = v1;
            OFFER_ID_SENVEN = v2;
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;

            }
        }
        #endregion
        #region save
        public void save(DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT AUDIT_STATUS FROM NO_PAPER_OFFER_MST WHERE NPID='" + NPID + "'");
            string vproject_Id = bc.getOnlyString("SELECT PROJECT_ID FROM NO_PAPER_OFFER_MST WHERE NPID='" + NPID + "'");
            string vOFFER_TYPE_CODE = bc.getOnlyString("SELECT SUBSTRING(OFFER_ID,5,1) FROM NO_PAPER_OFFER_DET WHERE NPID='" + NPID + "'");
            GETID_OFFER_ID(OFFER_TYPE_CODE);
            string NIKEY = bc.numYM(10, 4, "0001", "SELECT * FROM NO_PAPER_OFFER_ID_NO", "NIKEY", "NI");
            if (!bc.exists("SELECT NPID FROM NO_PAPER_OFFER_DET WHERE NPID='" + NPID + "'"))
            {
                SQlcommandE_MST(sqlt);
                SQlcommandE_DET(sqlo, dt);
                IFExecution_SUCCESS = true;
                OFFER_ID_SENVEN = numYY(8, 3, "001", "select * from NO_PAPER_OFFER_ID_NO", "OFFER_ID_SENVEN", OFFER_TYPE_CODE);
                if (!bc.exists("SELECT * FROM NO_PAPER_OFFER_ID_NO WHERE PIID='" + PIID +
                    "' AND YEAR='" + year + "' AND MONTH='" + month + "' AND SUBSTRING (OFFER_ID_SENVEN,5,1)='" + OFFER_TYPE_CODE + "'"))
                {
                    basec.getcoms(@"INSERT INTO NO_PAPER_OFFER_ID_NO(NIKEY,PIID,OFFER_ID_SENVEN,YEAR,MONTH,DATE)
VALUES ('"+NIKEY +"','" + PIID + "','" + OFFER_ID_SENVEN +
                            "','" + year + "','" + month + "','" + varDate + "')");
                }
            }
           else if (bc.exists("SELECT NPID FROM NO_PAPER_OFFER_DET WHERE NPID='" + NPID + "'") && v1 != "Y" && 
                PROJECT_ID ==vproject_Id && OFFER_TYPE_CODE==vOFFER_TYPE_CODE )//项目号且类别不变才修改原来的单据，否则要新增编号
            {
                //MessageBox.Show("EXISTS");
                SQlcommandE_MST(sqlth + " WHERE NPID='" + NPID + "'");
                SQlcommandE_DET(sqlo, dt);
                IFExecution_SUCCESS = true;
            }
            else
            {
                SQlcommandE_MST(sqlth);
                SQlcommandE_DET(sqlo, dt);
                IFExecution_SUCCESS = true;
                OFFER_ID_SENVEN = numYY(8, 3, "001", "select * from NO_PAPER_OFFER_ID_NO", "OFFER_ID_SENVEN", OFFER_TYPE_CODE);
                if (!bc.exists("SELECT * FROM NO_PAPER_OFFER_ID_NO WHERE PIID='" + PIID +
                    "' AND YEAR='" + year + "' AND MONTH='" + month + 
                    "' AND SUBSTRING (OFFER_ID_SENVEN,5,1)='" + OFFER_TYPE_CODE + "'"))
                { //编码时没有按OFFER_ID_SENVEN 排序导致写入重复编号

                    basec.getcoms(@"INSERT INTO NO_PAPER_OFFER_ID_NO(NIKEY,PIID,OFFER_ID_SENVEN,YEAR,MONTH,DATE)
VALUES ('" + NIKEY + "','" + PIID + "','" + OFFER_ID_SENVEN +
           "','" + year + "','" + month + "','" + varDate + "')");//加此判断条件，写入时判断这7码是吗存在系统160523
                }
            }
        }
        #endregion
        #region 编号 YY
        public string numYY(int digit, int wcodedigit, string wcode, string sql, string tbColumns, string prifix)
        {
            string year, month, day;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            day = DateTime.Now.ToString("dd");
            string P_str_Code, t, r, sql1, q = "";
            int P_int_Code, w, w1;//由于PIID为主键，如果不指定OFFER_ID_SENVEN为排序方式的话，默认是按PIID排序，实际用的是OFFER_ID_SENVEN的最后一行
            sql1 = sql + string.Format(" WHERE  YEAR='{0}' AND LEN({1})=8 AND MONTH={2} ORDER BY NIKEY,OFFER_ID_SENVEN ASC", year, tbColumns, month);
            DataTable dt = bc.getdt(sql1);

            if (dt.Rows.Count > 0)
            {
                P_str_Code = Convert.ToString(dt.Rows[(dt.Rows.Count - 1)][tbColumns]);
                w1 = digit - wcodedigit;
                P_int_Code = Convert.ToInt32(P_str_Code.Substring(w1, wcodedigit)) + 1;
                t = Convert.ToString(P_int_Code);
                w = wcodedigit - t.Length;
                if (w >= 0)
                {
                    while (w >= 1)
                    {
                        q = q + "0";
                        w = w - 1;

                    }
                    r = year + month + prifix + q + P_int_Code;
                }
                else
                {
                    r = "Exceed Limited";

                }

            }
            else
            {
                r = year + month + prifix + wcode;
            }
            return r;
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
            sqlcom.Parameters.Add("NPID", SqlDbType.VarChar, 20).Value = NPID;
            sqlcom.Parameters.Add("PROJECT_ID", SqlDbType.VarChar, 20).Value = PROJECT_ID;
            sqlcom.Parameters.Add("OFFER_DATE", SqlDbType.VarChar, 20).Value = OFFER_DATE;
            sqlcom.Parameters.Add("AUDIT_STATUS", SqlDbType.VarChar, 100).Value = AUDIT_STATUS;
            sqlcom.Parameters.Add("WATER_CODE", SqlDbType.VarChar, 100).Value = "";
            sqlcom.Parameters.Add("OFFER_TYPE_CODE", SqlDbType.VarChar, 20).Value = OFFER_TYPE_CODE;
            sqlcom.Parameters.Add("REMARK", SqlDbType.VarChar, 1000).Value = REMARK;
            sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region SQlcommandE_DET
        protected void SQlcommandE_DET(string sql, DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            basec.getcoms("DELETE NO_PAPER_OFFER_DET WHERE NPID='" + NPID + "'");
            foreach (DataRow dr in dt.Rows)
            {
                SqlConnection sqlcon = bc.getcon();
                sqlcon.Open();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                NPKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM NO_PAPER_OFFER_DET", "NPKEY", "NP");
                sqlcom.Parameters.Add("@NPKEY", SqlDbType.VarChar, 20).Value = NPKEY;
                sqlcom.Parameters.Add("@NPID", SqlDbType.VarChar, 20).Value = NPID;
                sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = dr["项次"].ToString();
                sqlcom.Parameters.Add("@COUNT", SqlDbType.VarChar, 20).Value = dr["数量"].ToString();
                sqlcom.Parameters.Add("@UNIT_PRICE", SqlDbType.VarChar, 20).Value = dr["报出价"].ToString();
                GETID_OFFER_ID(OFFER_TYPE_CODE);
                //MessageBox.Show(OFFER_ID + " -");
                sqlcom.Parameters.Add("@OFFER_ID", SqlDbType.VarChar, 20).Value = OFFER_ID;
                sqlcom.Parameters.Add("@PIID", SqlDbType.VarChar, 20).Value = PIID;
                sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = MAKERID;
                sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
            }

        }
        #endregion
        #region audit
        public List<string> audit()
        {
            try
            {
                list2 = new List<string>();
                SAMPLE_CODE = bc.getOnlyString("SELECT SAMPLE_CODE FROM EMPLOYEEINFO WHERE EMID='" + LOGIN_EMID + "'");
                SAMPLE_CODE_FIRST = SAMPLE_CODE.Substring(0, 1);
                DataTable dtx = bc.getdt(sql + " WHERE A.NPID='" + NPID + "'");
                if (dtx.Rows.Count > 0)
                {
                    if (AUDIT_STATUS == "N")
                    {
                        basec.getcoms(@"UPDATE NO_PAPER_OFFER_MST SET AUDIT_STATUS='Y',OFFER_ID='" + OFFER_ID + "-" + SAMPLE_CODE_FIRST +
                            "'   WHERE NPID='" + NPID + "'");
                        list2.Add("已审核");
                    }
                    else
                    {
                        basec.getcoms("UPDATE NO_PAPER_OFFER_MST SET AUDIT_STATUS='N' ,OFFER_ID='" + OFFER_ID.Substring(0, (OFFER_ID).Length - 2)
                            + "' WHERE NPID='" + NPID + "'");
                        list2.Add("未审核");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
            return list2;
        }
        #endregion
        #region RETURN_HAVE_ID_DT
        public DataTable RETURN_HAVE_ID_DT(DataTable dtt,string EMID,string POSITION)
        {
            DataTable dt = GetTableInfo_SEARCH();
            i = 1;
            foreach (DataRow dr1 in dtt.Rows)
            {
                if (!bc.exists(ccustomer_info.sqlsi + " WHERE B.CNAME='" + dr1["客户名称"].ToString() +
                          "' AND A.USER_MAKERID='" + EMID + "' ") && POSITION =="AE")
                {
                   
                }
                else
                {
                    DataRow dr = dt.NewRow();
                    dr["序号"] = i;
                    dr["项目号"] = dr1["项目号"].ToString();
                    dr["报价日期"] = dr1["报价日期"].ToString();
                    dr["项目名称"] = dr1["项目名称"].ToString();
                    dr["客户名称"] = dr1["客户名称"].ToString();
                    dr["品牌"] = dr1["品牌"].ToString();
                    dr["AE"] = dr1["AE"].ToString();
                    dr["平面设计"] = dr1["平面设计"].ToString();
                    dr["结构设计"] = dr1["结构设计"].ToString();
                    dr["报出价"] = dr1["报出价"].ToString();
                    dr["数量"] = dr1["数量"].ToString();
                    dr["报价编号"] = dr1["报价编号"].ToString();
                    dr["备注"] = dr1["备注"].ToString();
                    dt.Rows.Add(dr);
                    i = i + 1;
                }
            }
            return dt;
        }
        #endregion
        #region ExcelPrint
        public void ExcelPrint(DataTable dt, string BillName, string Printpath)
        {
            //int j;
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);
            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;

            for (i = 0; i < dt.Rows.Count; i++)
            {
                worksheet.Cells[3 + i, "A"] = dt.Rows[i]["项目号"].ToString();
                worksheet.Cells[3 + i, "B"] = dt.Rows[i]["审核状态"].ToString();
                worksheet.Cells[3 + i, "C"] = dt.Rows[i]["客户名称"].ToString();
                worksheet.Cells[3 + i, "D"] = dt.Rows[i]["AE"].ToString();
                worksheet.Cells[3 + i, "E"] = dt.Rows[i]["AE助理-1"].ToString();
                worksheet.Cells[3 + i, "F"] = dt.Rows[i]["AE助理-2"].ToString();
                worksheet.Cells[3 + i, "G"] = dt.Rows[i]["平面设计"].ToString();
                worksheet.Cells[3 + i, "H"] = dt.Rows[i]["平面设计助理-1"].ToString();
                worksheet.Cells[3 + i, "I"] = dt.Rows[i]["平面设计助理-2"].ToString();

                worksheet.Cells[3 + i, "J"] = dt.Rows[i]["结构设计"].ToString();
                worksheet.Cells[3 + i, "K"] = dt.Rows[i]["结构设计助理-1"].ToString();
                worksheet.Cells[3 + i, "L"] = dt.Rows[i]["结构设计助理-2"].ToString();

                worksheet.Cells[3 + i, "M"] = dt.Rows[i]["报价"].ToString();
                worksheet.Cells[3 + i, "N"] = dt.Rows[i]["审核"].ToString();

            }
            worksheet.get_Range(worksheet.Cells[3, "A"], worksheet.Cells[3 + i - 1, "N"]).Borders.LineStyle = 1;
            //workbook.Save();
            //bc.csharpExcelPrint(sfdg.FileName);
            /*application.Quit();
            worksheet = null;
            workbook = null;
            application = null;
            GC.Collect();*/
        }
        #endregion
        #region RETURN_PFID_NPID_DT
        public DataTable RETURN_PFID_NPID_DT(string PROJECT_ID)//避免加载过多出现死机，要根据具体项目号来查询160612
        {
            DataTable dtx = bc.getdt(string.Format(cprinting_offer.sqlte + "  WHERE C.PROJECT_ID='" + PROJECT_ID + "' ORDER BY B.OFFER_ID ASC"));//纸品报价
            DataTable dtx1 = new DataTable();
            dtx1.Columns.Add("项目号", typeof(string));
            dtx1.Columns.Add("项目名称", typeof(string));
            dtx1.Columns.Add("编号", typeof(string));
            dtx1.Columns.Add("报价编号", typeof(string));
            dtx1.Columns.Add("审核状态", typeof(string));
            dtx1.Columns.Add("客户名称", typeof(string));
            dtx1.Columns.Add("品牌", typeof(string));
            dtx1.Columns.Add("AE01", typeof(string));
            dtx1.Columns.Add("AE02", typeof(string));
            dtx1.Columns.Add("AE03", typeof(string));
            dtx1.Columns.Add("平面01", typeof(string));
            dtx1.Columns.Add("平面02", typeof(string));
            dtx1.Columns.Add("平面03", typeof(string));
            dtx1.Columns.Add("结构01", typeof(string));
            dtx1.Columns.Add("结构02", typeof(string));
            dtx1.Columns.Add("结构03", typeof(string));
            dtx1.Columns.Add("数量", typeof(string));
            dtx1.Columns.Add("报出价", typeof(string));

            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    DataRow dr1 = dtx1.NewRow();
                    dr1["编号"] = dr["编号"].ToString();
                    dr1["报价编号"] = dr["报价编号"].ToString();
                    dr1["项目号"] = dr["项目号"].ToString();
                    if (dr["审核状态"].ToString() == "未审核")
                    {
                        dr1["审核状态"] = "待报价";
                    }
                    else
                    {
                        dr1["审核状态"] = "已核价";
                    }
                    dr1["项目名称"] = dr["项目名称"].ToString();
                    dr1["客户名称"] = dr["客户名称"].ToString();
                    dr1["品牌"] = dr["品牌"].ToString();

                    dtx1.Rows.Add(dr1);
                }
            }
            dtx = bc.getdt(string.Format(this.sql + "  WHERE B.PROJECT_ID='" + PROJECT_ID + "' ORDER BY D.OFFER_ID ASC"));//非纸品报价
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    DataRow dr1 = dtx1.NewRow();
                    dr1["编号"] = dr["编号"].ToString();
                    dr1["报价编号"] = dr["报价编号"].ToString();
                    dr1["项目号"] = dr["项目号"].ToString();
                    dr1["项目名称"] = dr["项目名称"].ToString();
                    dr1["客户名称"] = dr["客户名称"].ToString();
                    dr1["品牌"] = dr["品牌"].ToString();
                    dr1["AE01"] = dr["AE"].ToString();
                    dr1["平面01"] = dr["平面设计"].ToString();
                    dr1["结构01"] = dr["结构设计"].ToString();
                    if (dr["审核状态"].ToString() == "未审核")
                    {
                        dr1["审核状态"] = "待报价";
                    }
                    else
                    {
                        dr1["审核状态"] = "已核价";
                    }
                    dr1["数量"] = dr["数量"].ToString();
                    dr1["报出价"] = dr["报出价"].ToString();
                    dtx1.Rows.Add(dr1);
                }
            }
            if (dtx1.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx1.Rows)
                {
                    DataTable dtx2 = bc.getdt(cproject_info.sql + string.Format(" WHERE A.PROJECT_ID='{0}'", dr["项目号"].ToString()));
                    if (dtx2.Rows.Count > 0)
                    {
                        dr["AE01"] = dtx2.Rows[0]["AE01"].ToString();
                        dr["AE02"] = dtx2.Rows[0]["AE02"].ToString();
                        dr["AE03"] = dtx2.Rows[0]["AE03"].ToString();
                        dr["平面01"] = dtx2.Rows[0]["平面01"].ToString();
                        dr["平面02"] = dtx2.Rows[0]["平面02"].ToString();
                        dr["平面03"] = dtx2.Rows[0]["平面03"].ToString();
                        dr["结构01"] = dtx2.Rows[0]["结构01"].ToString();
                        dr["结构02"] = dtx2.Rows[0]["结构02"].ToString();
                        dr["结构03"] = dtx2.Rows[0]["结构03"].ToString();
                    }
                    DataTable dt = cprinting_offer.RETURN_SEARCH(bc.getdt(cprinting_offer.sqlse + " WHERE C.PROJECT_ID ='" + dr["项目号"].ToString() + "'"));//取得报出价
                    dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "报价编号='" + dr["报价编号"].ToString() + "'");//cprinting_offer.RETURN_SEARCH是按项目号统计所以要加这句
                    if (dt.Rows.Count > 0)
                    {
                        dr["报出价"] = dt.Rows[0]["报出价"].ToString();
                        dr["数量"] = dt.Rows[0]["报价数量"].ToString();
                    }
                    
                }
            }
            return dtx1;
        }
        #endregion
        #region RETURN_PFID_NPID_DT_FROM_OFFER_ID
        public DataTable RETURN_PFID_NPID_DT_FROM_OFFER_ID(string OFFER_ID)
        {
            DataTable dtx = bc.getdt(string.Format(cprinting_offer.sqlte + "  WHERE B.OFFER_ID='" + OFFER_ID + "' ORDER BY B.OFFER_ID ASC"));
            DataTable dtx1 = new DataTable();
            dtx1.Columns.Add("项目号", typeof(string));
            dtx1.Columns.Add("项目名称", typeof(string));
            dtx1.Columns.Add("编号", typeof(string));
            dtx1.Columns.Add("报价编号", typeof(string));
            dtx1.Columns.Add("审核状态", typeof(string));
            dtx1.Columns.Add("客户名称", typeof(string));
            dtx1.Columns.Add("品牌", typeof(string));
            dtx1.Columns.Add("AE01", typeof(string));
            dtx1.Columns.Add("AE02", typeof(string));
            dtx1.Columns.Add("AE03", typeof(string));
            dtx1.Columns.Add("平面01", typeof(string));
            dtx1.Columns.Add("平面02", typeof(string));
            dtx1.Columns.Add("平面03", typeof(string));
            dtx1.Columns.Add("结构01", typeof(string));
            dtx1.Columns.Add("结构02", typeof(string));
            dtx1.Columns.Add("结构03", typeof(string));
            dtx1.Columns.Add("数量", typeof(string));
            dtx1.Columns.Add("报出价", typeof(string));

            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    DataRow dr1 = dtx1.NewRow();
                    dr1["编号"] = dr["编号"].ToString();
                    dr1["报价编号"] = dr["报价编号"].ToString();
                    dr1["项目号"] = dr["项目号"].ToString();
                    if (dr["审核状态"].ToString() == "未审核")
                    {
                        dr1["审核状态"] = "待报价";
                    }
                    else
                    {
                        dr1["审核状态"] = "已核价";
                    }
                    dr1["项目名称"] = dr["项目名称"].ToString();
                    dr1["客户名称"] = dr["客户名称"].ToString();
                    dr1["品牌"] = dr["品牌"].ToString();

                    dtx1.Rows.Add(dr1);
                }
            }
            dtx = bc.getdt(string.Format(this.sql + "  WHERE D.OFFER_ID='" + OFFER_ID + "' ORDER BY D.OFFER_ID ASC"));
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    DataRow dr1 = dtx1.NewRow();
                    dr1["编号"] = dr["编号"].ToString();
                    dr1["报价编号"] = dr["报价编号"].ToString();
                    dr1["项目号"] = dr["项目号"].ToString();
                    dr1["项目名称"] = dr["项目名称"].ToString();
                    dr1["客户名称"] = dr["客户名称"].ToString();
                    dr1["品牌"] = dr["品牌"].ToString();
                    dr1["AE01"] = dr["AE"].ToString();
                    dr1["平面01"] = dr["平面设计"].ToString();
                    dr1["结构01"] = dr["结构设计"].ToString();
                    if (dr["审核状态"].ToString() == "未审核")
                    {
                        dr1["审核状态"] = "待报价";
                    }
                    else
                    {
                        dr1["审核状态"] = "已核价";
                    }
                    dr1["数量"] = dr["数量"].ToString();
                    dr1["报出价"] = dr["报出价"].ToString();
                    dtx1.Rows.Add(dr1);
                }
            }
            if (dtx1.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx1.Rows)
                {
                    DataTable dtx2 = bc.getdt(cproject_info.sql + string.Format(" WHERE A.PROJECT_ID='{0}'", dr["项目号"].ToString()));
                    if (dtx2.Rows.Count > 0)
                    {
                        dr["AE01"] = dtx2.Rows[0]["AE01"].ToString();
                        dr["AE02"] = dtx2.Rows[0]["AE02"].ToString();
                        dr["AE03"] = dtx2.Rows[0]["AE03"].ToString();
                        dr["平面01"] = dtx2.Rows[0]["平面01"].ToString();
                        dr["平面02"] = dtx2.Rows[0]["平面02"].ToString();
                        dr["平面03"] = dtx2.Rows[0]["平面03"].ToString();
                        dr["结构01"] = dtx2.Rows[0]["结构01"].ToString();
                        dr["结构02"] = dtx2.Rows[0]["结构02"].ToString();
                        dr["结构03"] = dtx2.Rows[0]["结构03"].ToString();
                    }
                    DataTable dt = cprinting_offer.RETURN_SEARCH(bc.getdt(cprinting_offer.sqlse + " WHERE C.PROJECT_ID ='" + dr["项目号"].ToString() + "'"));//取得报出价
                    dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "报价编号='" + dr["报价编号"].ToString() + "'");//cprinting_offer.RETURN_SEARCH是按项目号统计所以要加这句
                    if (dt.Rows.Count > 0)
                    {
                        dr["报出价"] = dt.Rows[0]["报出价"].ToString();
                        dr["数量"] = dt.Rows[0]["报价数量"].ToString();
                    }
                }
            }
            return dtx1;
        }
        #endregion
        #region RETURN_PFID_NPID
        public string  RETURN_PFID_NPID(string OFFER_ID)
        {
            string RETURN_PFID_NPID = "";
            DataTable dtx = bc.getdt(string.Format(cprinting_offer.sqlte + "  WHERE B.OFFER_ID='" + OFFER_ID + "' ORDER BY B.OFFER_ID ASC"));
            if (dtx.Rows.Count > 0)
            {
                RETURN_PFID_NPID = dtx.Rows[0]["编号"].ToString();
            }
            dtx = bc.getdt(string.Format(this.sql + "  WHERE D.OFFER_ID='" + OFFER_ID + "' ORDER BY D.OFFER_ID ASC"));
            if (dtx.Rows.Count > 0)
            {
                RETURN_PFID_NPID = dtx.Rows[0]["编号"].ToString();
            }
            return RETURN_PFID_NPID;
        }
        #endregion
    
    }
}
