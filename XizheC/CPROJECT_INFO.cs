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
using Excel = Microsoft.Office.Interop.Excel;

namespace XizheC
{
    public class CPROJECT_INFO
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _BRAND;
        public string BRAND
        {
            set { _BRAND = value; }
            get { return _BRAND; }

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

        private  bool _IFExecutionSUCCESS;
        public  bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private string _PIID;
        public string PIID
        {
            set { _PIID = value; }
            get { return _PIID; }

        }
        private string _PROJECT_ID;
        public string PROJECT_ID
        {
            set { _PROJECT_ID = value; }
            get { return _PROJECT_ID; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        private string _CUID;
        public string CUID
        {
            set { _CUID = value; }
            get { return _CUID; }

        }
        private string _AE_MAKERID_ONE;
        public string AE_MAKERID_ONE
        {
            set { _AE_MAKERID_ONE = value; }
            get { return _AE_MAKERID_ONE; }

        }
        private string _AE_MAKERID_TWO;
        public string AE_MAKERID_TWO
        {
            set { _AE_MAKERID_TWO = value; }
            get { return _AE_MAKERID_TWO; }

        }
        private string _AE_MAKERID_THREE;
        public string AE_MAKERID_THREE
        {
            set { _AE_MAKERID_THREE = value; }
            get { return _AE_MAKERID_THREE; }

        }

        private string _PLANE_MAKERID_ONE;
        public string PLANE_MAKERID_ONE
        {
            set { _PLANE_MAKERID_ONE = value; }
            get { return _PLANE_MAKERID_ONE; }

        }
        private string _PLANE_MAKERID_TWO;
        public string PLANE_MAKERID_TWO
        {
            set { _PLANE_MAKERID_TWO = value; }
            get { return _PLANE_MAKERID_TWO; }

        }
        private string _PLANE_MAKERID_THREE;
        public string PLANE_MAKERID_THREE
        {
            set { _PLANE_MAKERID_THREE = value; }
            get { return _PLANE_MAKERID_THREE; }

        }

        private string _STRUCTURE_MAKERID_ONE;
        public string STRUCTURE_MAKERID_ONE
        {
            set { _STRUCTURE_MAKERID_ONE = value; }
            get { return _STRUCTURE_MAKERID_ONE; }

        }
        private string _STRUCTURE_MAKERID_TWO;
        public string STRUCTURE_MAKERID_TWO
        {
            set { _STRUCTURE_MAKERID_TWO = value; }
            get { return _STRUCTURE_MAKERID_TWO; }

        }
        private string _STRUCTURE_MAKERID_THREE;
        public string STRUCTURE_MAKERID_THREE
        {
            set { _STRUCTURE_MAKERID_THREE = value; }
            get { return _STRUCTURE_MAKERID_THREE; }

        }
        private string _OFFER_MAKERID;
        public string OFFER_MAKERID
        {
            set { _OFFER_MAKERID = value; }
            get { return _OFFER_MAKERID; }

        }
        private string _AUDIT_MAKERID;
        public string AUDIT_MAKERID
        {
            set { _AUDIT_MAKERID = value; }
            get { return _AUDIT_MAKERID; }

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
 

        private string _PIKEY;
        public string PIKEY
        {

            set { _PIKEY = value; }
            get { return _PIKEY; }

        }
  
        #endregion
        DataTable dt = new DataTable();
        int i;
        #region sql
        string setsql = @"
SELECT 
A.PIID AS 项目编号,
A.PROJECT_ID AS 项目号,
A.PROJECT_NAME AS 项目名称,
A.CUID AS 客户ID,
A.BRAND AS 品牌,
(SELECT CNAME FROM CUSTOMERINFO_MST WHERE CUID=A.CUID) AS 客户名称,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.AE_MAKERID_1) AS AE01工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.AE_MAKERID_1) AS AE01,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.AE_MAKERID_2) AS AE02工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.AE_MAKERID_2) AS AE02,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.AE_MAKERID_3) AS AE03工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.AE_MAKERID_3) AS AE03,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.PLANE_MAKERID_1) AS 平面01工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PLANE_MAKERID_1) AS 平面01,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.PLANE_MAKERID_2) AS 平面02工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PLANE_MAKERID_2) AS 平面02,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.PLANE_MAKERID_3)AS 平面03工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PLANE_MAKERID_3) AS 平面03,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_MAKERID_1) AS 结构01工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_MAKERID_1) AS 结构01,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_MAKERID_2)  AS 结构02工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_MAKERID_2) AS 结构02,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_MAKERID_3) AS 结构03工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_MAKERID_3) AS 结构03,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.OFFER_MAKERID)  AS 报价工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.OFFER_MAKERID) AS 报价,
(SELECT EMPLOYEE_ID FROM EMPLOYEEINFO WHERE EMID=A.AUDIT_MAKERID) AS 审核工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.AUDIT_MAKERID) AS 审核,
A.MakerID AS 制单人工号,
A.Date AS 制单日期
FROM PROJECT_INFO A


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PROJECT_INFO
(
PIID,
PROJECT_ID,
PROJECT_NAME,
CUID,
BRAND,
AE_MAKERID_1,
AE_MAKERID_2,
AE_MAKERID_3,
PLANE_MAKERID_1,
PLANE_MAKERID_2,
PLANE_MAKERID_3,
STRUCTURE_MAKERID_1,
STRUCTURE_MAKERID_2,
STRUCTURE_MAKERID_3,
MakerID,
Date,
YEAR,
MONTH
)
VALUES
(
@PIID,
@PROJECT_ID,
@PROJECT_NAME,
@CUID,
@BRAND,
@AE_MAKERID_1,
@AE_MAKERID_2,
@AE_MAKERID_3,
@PLANE_MAKERID_1,
@PLANE_MAKERID_2,
@PLANE_MAKERID_3,
@STRUCTURE_MAKERID_1,
@STRUCTURE_MAKERID_2,
@STRUCTURE_MAKERID_3,
@MakerID,
@Date,
@YEAR,
@MONTH

)
";
        string setsqlth = @"
UPDATE PROJECT_INFO SET 
PIID=@PIID,
PROJECT_NAME=@PROJECT_NAME,
CUID=@CUID,
BRAND=@BRAND,
AE_MAKERID_1=@AE_MAKERID_1,
AE_MAKERID_2=@AE_MAKERID_2,
AE_MAKERID_3=@AE_MAKERID_3,
PLANE_MAKERID_1=@PLANE_MAKERID_1,
PLANE_MAKERID_2=@PLANE_MAKERID_2,
PLANE_MAKERID_3=@PLANE_MAKERID_3,
STRUCTURE_MAKERID_1=@STRUCTURE_MAKERID_1,
STRUCTURE_MAKERID_2=@STRUCTURE_MAKERID_2,
STRUCTURE_MAKERID_3=@STRUCTURE_MAKERID_3,
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
        public CPROJECT_INFO()
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
            dt.Columns.Add("项目号", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("客户名称", typeof(string));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("AE助理-1", typeof(string));
            dt.Columns.Add("AE助理-2", typeof(string));
            dt.Columns.Add("平面设计", typeof(string));
            dt.Columns.Add("平面设计助理-1", typeof(string));
            dt.Columns.Add("平面设计助理-2", typeof(string));
            dt.Columns.Add("结构设计", typeof(string));
            dt.Columns.Add("结构设计助理-1", typeof(string));
            dt.Columns.Add("结构设计助理-2", typeof(string));
            dt.Columns.Add("报价", typeof(string));
            dt.Columns.Add("审核", typeof(string));
            return dt;
        }
        #endregion
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from PROJECT_INFO_NO", "PIID", "PI");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO PROJECT_INFO_NO(PIID,DATE,YEAR,MONTH) VALUES ('" + v1 + "','" + varDate + "','" + year +
                  "','" + month + "')");
              
            }
            return GETID;
        }
        public string GETID_PROJECT_ID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = numYM(11, 3, "001", "select * from PROJECT_INFO ", "PROJECT_ID", "DBXM");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;

            }
            return GETID;
        }
        #region 编号YM
        public string numYM(int digit, int wcodedigit, string wcode, string sql, string tbColumns, string prifix)
        {
            string year, month;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            string P_str_Code, t, r, sql1, q = "";
            int P_int_Code, w, w1;

            sql1 = sql + " WHERE YEAR='" + year + "' AND  MONTH='" + month + "' ORDER BY PROJECT_ID ASC";
            SqlDataReader sqlread = bc.getread(sql1);
            DataTable dt = bc.getdt(sql1);
            sqlread.Read();
            if (sqlread.HasRows)
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
                    r = prifix + year + month + q + P_int_Code;
                }
                else
                {
                    r = "Exceed Limited";

                }

            }
            else
            {
                r = prifix + year + month + wcode;
            }
            sqlread.Close();
            return r;
        }
        #endregion
        #region save
        public void save()
        {

            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string GET_PROJECT_NAME = bc.getOnlyString("SELECT PROJECT_NAME FROM PROJECT_INFO WHERE  PIID='" + PIID + "'");

            if (!bc.exists("SELECT PIID FROM PROJECT_INFO WHERE PIID='" + PIID + "'"))
            {
                if (bc.exists("SELECT * FROM PROJECT_INFO where PROJECT_NAME='" + PROJECT_NAME + "'"))
                {
                    ErrowInfo = string.Format("项目名称：{0}" + " 已经存在系统", PROJECT_NAME);
                    IFExecution_SUCCESS = false;
                }

                else
                {

                   
                    SQlcommandE(sqlt,dt);
                    IFExecution_SUCCESS = true;
                }
            }
            else if (GET_PROJECT_NAME != PROJECT_NAME)
            {
                if (bc.exists("SELECT * FROM PROJECT_INFO where PROJECT_NAME='" + PROJECT_NAME + "'"))
                {

                    ErrowInfo = string.Format("项目名称：{0}" + " 已经存在系统", PROJECT_NAME);
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    
                    SQlcommandE(sqlth + " WHERE PIID='" + PIID + "'",dt);
                    IFExecution_SUCCESS = true;
                }
            }

            else
            {
             
                SQlcommandE(sqlth + " WHERE PIID='" + PIID + "'",dt);
                IFExecution_SUCCESS = true;
            }
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql,DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcon.Open();
            sqlcom.Parameters.Add("PIID", SqlDbType.VarChar, 20).Value = PIID;
            sqlcom.Parameters.Add("PROJECT_ID", SqlDbType.VarChar, 20).Value = PROJECT_ID;
            sqlcom.Parameters.Add("PROJECT_NAME", SqlDbType.VarChar, 100).Value = PROJECT_NAME;
            sqlcom.Parameters.Add("CUID", SqlDbType.VarChar, 20).Value = CUID;
            sqlcom.Parameters.Add("BRAND", SqlDbType.VarChar, 20).Value = BRAND;
            sqlcom.Parameters.Add("AE_MAKERID_1", SqlDbType.VarChar, 20).Value = AE_MAKERID_ONE;
            sqlcom.Parameters.Add("AE_MAKERID_2", SqlDbType.VarChar, 20).Value = AE_MAKERID_TWO;
            sqlcom.Parameters.Add("AE_MAKERID_3", SqlDbType.VarChar, 20).Value = AE_MAKERID_THREE;
            sqlcom.Parameters.Add("PLANE_MAKERID_1", SqlDbType.VarChar, 20).Value = PLANE_MAKERID_ONE;
            sqlcom.Parameters.Add("PLANE_MAKERID_2", SqlDbType.VarChar, 20).Value = PLANE_MAKERID_TWO;
            sqlcom.Parameters.Add("PLANE_MAKERID_3", SqlDbType.VarChar, 20).Value = PLANE_MAKERID_THREE;
            sqlcom.Parameters.Add("STRUCTURE_MAKERID_1", SqlDbType.VarChar, 20).Value = STRUCTURE_MAKERID_ONE;
            sqlcom.Parameters.Add("STRUCTURE_MAKERID_2", SqlDbType.VarChar, 20).Value = STRUCTURE_MAKERID_TWO;
            sqlcom.Parameters.Add("STRUCTURE_MAKERID_3", SqlDbType.VarChar, 20).Value = STRUCTURE_MAKERID_THREE;
            sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region RETURN_DT
        public DataTable RETURN_DT(DataTable dtt)
        {
            DataTable dt = GetTableInfo_SEARCH();
            foreach (DataRow dr1 in dtt.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["项目号"] = dr1["项目号"].ToString();
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["客户名称"] = dr1["客户名称"].ToString();
                dr["AE"] = dr1["AE01"].ToString();
                dr["AE助理-1"] = dr1["AE02"].ToString();
                dr["AE助理-2"] = dr1["AE03"].ToString();
                dr["平面设计"] = dr1["平面01"].ToString();
                dr["平面设计助理-1"] = dr1["平面02"].ToString();
                dr["平面设计助理-2"] = dr1["平面03"].ToString();
                dr["结构设计"] = dr1["结构01"].ToString();
                dr["结构设计助理-1"] = dr1["结构02"].ToString();
                dr["结构设计助理-2"] = dr1["结构03"].ToString();
                dr["报价"] = dr1["报价"].ToString();
                dr["审核"] = dr1["审核"].ToString();
                dt.Rows.Add(dr);
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
                worksheet.Cells[3 + i, "B"] = dt.Rows[i]["项目名称"].ToString();
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
            worksheet .get_Range(worksheet .Cells [3,"A"],worksheet .Cells [3+i-1,"N"]).Borders.LineStyle = 1;
            //workbook.Save();
            //bc.csharpExcelPrint(sfdg.FileName);
            /*application.Quit();
            worksheet = null;
            workbook = null;
            application = null;
            GC.Collect();*/
        }
        #endregion
    
    }
}
