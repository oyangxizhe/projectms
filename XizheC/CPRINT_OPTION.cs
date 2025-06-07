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
    public class CPRINT_OPTION:IGETID 
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

        private string _PRINT_OPTION;
        public string PRINT_OPTION
        {
            set { _PRINT_OPTION = value; }
            get { return _PRINT_OPTION; }

        }
        private string _POID;
        public string POID
        {
            set { _POID = value; }
            get { return _POID; }
        }
        private string _DEBURRING;
        public string DEBURRING
        {
            set { _DEBURRING = value; }
            get { return _DEBURRING; }
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
A.PRINT_OPTION AS 印刷选项,
A.DEBURRING AS 修边,
A.TISSUE_A AS 面纸内耗1到300,
A.TISSUE_B AS 面纸内耗大于300,
A.BODY_PAPER_A AS 底纸内耗1到300,
A.BODY_PAPER_B AS 底纸内耗大于300,
A.NO_PRINTING_A AS 无印刷用纸表面处理损耗_固定值,
A.NO_PRINTING_B AS 无印刷用纸表面处理损耗_百分比,
A.POSITIVE_PRINTING_A AS 正面印刷纸张损耗_A,
A.POSITIVE_PRINTING_B AS 正面印刷纸张损耗_B,
A.POSITIVE_PRINTING_C AS 正面印刷纸张损耗_C,
A.POSITIVE_PRINTING_D AS 正面印刷纸张损耗_D,
A.POSITIVE_PRINTING_E AS 正面印刷纸张损耗_E,
A.POSITIVE_PRINTING_F AS 正面印刷纸张损耗_F,
A.POSITIVE_PRINTING_G AS 正面印刷纸张损耗_G,
A.POSITIVE_PRINTING_H AS 正面印刷纸张损耗_H,
A.POSITIVE_PRINTING_I AS 正面印刷纸张损耗_I,
A.POSITIVE_PRINTING_J AS 正面印刷纸张损耗_J,
A.OPPOSITE_PRINTING_A AS 反面印刷纸张损耗_A,
A.OPPOSITE_PRINTING_B AS 反面印刷纸张损耗_B,
A.OPPOSITE_PRINTING_C AS 反面印刷纸张损耗_C,
A.OPPOSITE_PRINTING_D AS 反面印刷纸张损耗_D,
A.OPPOSITE_PRINTING_E AS 反面印刷纸张损耗_E,
A.OPPOSITE_PRINTING_F AS 反面印刷纸张损耗_F,
A.OPPOSITE_PRINTING_G AS 反面印刷纸张损耗_G,
A.OPPOSITE_PRINTING_H AS 反面印刷纸张损耗_H,
A.OPPOSITE_PRINTING_I AS 反面印刷纸张损耗_I,
A.OPPOSITE_PRINTING_J AS 反面印刷纸张损耗_J,
B.ENAME AS 制单人,
A.Date AS 制单日期
FROM PRINT_OPTION A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_OPTION
(
POID,
PRINT_OPTION,
DEBURRING,
TISSUE_A,
TISSUE_B,
BODY_PAPER_A,
BODY_PAPER_B,
NO_PRINTING_A,
NO_PRINTING_B,
POSITIVE_PRINTING_A,
POSITIVE_PRINTING_B,
POSITIVE_PRINTING_C,
POSITIVE_PRINTING_D,
POSITIVE_PRINTING_E,
POSITIVE_PRINTING_F,
POSITIVE_PRINTING_G,
POSITIVE_PRINTING_H,
POSITIVE_PRINTING_I,
POSITIVE_PRINTING_J,
OPPOSITE_PRINTING_A,
OPPOSITE_PRINTING_B,
OPPOSITE_PRINTING_C,
OPPOSITE_PRINTING_D,
OPPOSITE_PRINTING_E,
OPPOSITE_PRINTING_F,
OPPOSITE_PRINTING_G,
OPPOSITE_PRINTING_H,
OPPOSITE_PRINTING_I,
OPPOSITE_PRINTING_J,
MakerID,
Date,
Year,
Month
)
VALUES
(
@POID,
@PRINT_OPTION,
@DEBURRING,
@TISSUE_A,
@TISSUE_B,
@BODY_PAPER_A,
@BODY_PAPER_B,
@NO_PRINTING_A,
@NO_PRINTING_B,
@POSITIVE_PRINTING_A,
@POSITIVE_PRINTING_B,
@POSITIVE_PRINTING_C,
@POSITIVE_PRINTING_D,
@POSITIVE_PRINTING_E,
@POSITIVE_PRINTING_F,
@POSITIVE_PRINTING_G,
@POSITIVE_PRINTING_H,
@POSITIVE_PRINTING_I,
@POSITIVE_PRINTING_J,
@OPPOSITE_PRINTING_A,
@OPPOSITE_PRINTING_B,
@OPPOSITE_PRINTING_C,
@OPPOSITE_PRINTING_D,
@OPPOSITE_PRINTING_E,
@OPPOSITE_PRINTING_F,
@OPPOSITE_PRINTING_G,
@OPPOSITE_PRINTING_H,
@OPPOSITE_PRINTING_I,
@OPPOSITE_PRINTING_J,
@MakerID,
@Date,
@Year,
@Month
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
      
        public CPRINT_OPTION()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM PRINT_OPTION", "POID", "PO");
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
            dt.Columns.Add("印刷选项", typeof(string));
            dt.Columns.Add("修边", typeof(string));
            dt.Columns.Add("面纸内耗1到300", typeof(string));
            dt.Columns.Add("面纸内耗大于300", typeof(string));
            dt.Columns.Add("底纸内耗1到300", typeof(string));
            dt.Columns.Add("底纸内耗大于300", typeof(string));
            dt.Columns.Add("无印刷用纸表面处理损耗_固定值", typeof(string));
            dt.Columns.Add("无印刷用纸表面处理损耗_百分比", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_A", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_B", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_C", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_D", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_E", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_F", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_G", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_H", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_I", typeof(string));
            dt.Columns.Add("正面印刷纸张损耗_J", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_A", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_B", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_C", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_D", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_E", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_F", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_G", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_H", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_I", typeof(string));
            dt.Columns.Add("反面印刷纸张损耗_J", typeof(string));
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
                dr["印刷选项"] = dr1["印刷选项"].ToString();
                dr["修边"] = dr1["修边"].ToString();
                dr["面纸内耗1到300"] = dr1["面纸内耗1到300"].ToString();
                dr["面纸内耗大于300"] = dr1["面纸内耗大于300"].ToString();
                dr["底纸内耗1到300"] = dr1["底纸内耗1到300"].ToString();
                dr["底纸内耗大于300"] = dr1["底纸内耗大于300"].ToString();
                dr["无印刷用纸表面处理损耗_固定值"] = dr1["无印刷用纸表面处理损耗_固定值"].ToString();
                dr["无印刷用纸表面处理损耗_百分比"] = dr1["无印刷用纸表面处理损耗_百分比"].ToString();
                dr["正面印刷纸张损耗_A"] = dr1["正面印刷纸张损耗_A"].ToString();
                dr["正面印刷纸张损耗_B"] = dr1["正面印刷纸张损耗_B"].ToString();
                dr["正面印刷纸张损耗_C"] = dr1["正面印刷纸张损耗_C"].ToString();
                dr["正面印刷纸张损耗_D"] = dr1["正面印刷纸张损耗_D"].ToString();
                dr["正面印刷纸张损耗_E"] = dr1["正面印刷纸张损耗_E"].ToString();
                dr["正面印刷纸张损耗_F"] = dr1["正面印刷纸张损耗_F"].ToString();
                dr["正面印刷纸张损耗_G"] = dr1["正面印刷纸张损耗_G"].ToString();
                dr["正面印刷纸张损耗_H"] = dr1["正面印刷纸张损耗_H"].ToString();
                dr["正面印刷纸张损耗_I"] = dr1["正面印刷纸张损耗_I"].ToString();
                dr["正面印刷纸张损耗_J"] = dr1["正面印刷纸张损耗_J"].ToString();
                dr["反面印刷纸张损耗_A"] = dr1["反面印刷纸张损耗_A"].ToString();
                dr["反面印刷纸张损耗_B"] = dr1["反面印刷纸张损耗_B"].ToString();
                dr["反面印刷纸张损耗_C"] = dr1["反面印刷纸张损耗_C"].ToString();
                dr["反面印刷纸张损耗_D"] = dr1["反面印刷纸张损耗_D"].ToString();
                dr["反面印刷纸张损耗_E"] = dr1["反面印刷纸张损耗_E"].ToString();
                dr["反面印刷纸张损耗_F"] = dr1["反面印刷纸张损耗_F"].ToString();
                dr["反面印刷纸张损耗_G"] = dr1["反面印刷纸张损耗_G"].ToString();
                dr["反面印刷纸张损耗_H"] = dr1["反面印刷纸张损耗_H"].ToString();
                dr["反面印刷纸张损耗_I"] = dr1["反面印刷纸张损耗_I"].ToString();
                dr["反面印刷纸张损耗_J"] = dr1["反面印刷纸张损耗_J"].ToString();
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
            basec.getcoms("DELETE PRINT_OPTION");
            SQlcommandE(sqlt ,dt);
            IFExecution_SUCCESS = true;
 
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql ,DataTable dt)
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
                if (dr["印刷选项"].ToString() != "")
                {
                    sqlcom.Parameters.Add("POID", SqlDbType.VarChar, 20).Value = GETID();
                    sqlcom.Parameters.Add("PRINT_OPTION", SqlDbType.VarChar, 20).Value = dr["印刷选项"].ToString();
                    sqlcom.Parameters.Add("DEBURRING", SqlDbType.VarChar, 20).Value = dr["修边"].ToString();
                    sqlcom.Parameters.Add("TISSUE_A", SqlDbType.VarChar, 20).Value = dr["面纸内耗1到300"].ToString();
                    sqlcom.Parameters.Add("TISSUE_B", SqlDbType.VarChar, 20).Value = dr["面纸内耗大于300"].ToString();
                    sqlcom.Parameters.Add("BODY_PAPER_A", SqlDbType.VarChar, 20).Value = dr["底纸内耗1到300"].ToString();
                    sqlcom.Parameters.Add("BODY_PAPER_B", SqlDbType.VarChar, 20).Value = dr["底纸内耗大于300"].ToString();
                    sqlcom.Parameters.Add("NO_PRINTING_A", SqlDbType.VarChar, 20).Value = dr["无印刷用纸表面处理损耗_固定值"].ToString();
                    sqlcom.Parameters.Add("NO_PRINTING_B", SqlDbType.VarChar, 20).Value = dr["无印刷用纸表面处理损耗_百分比"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_A", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_A"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_B", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_B"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_C", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_C"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_D", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_D"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_E", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_E"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_F", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_F"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_G", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_G"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_H", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_H"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_I", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_I"].ToString();
                    sqlcom.Parameters.Add("POSITIVE_PRINTING_J", SqlDbType.VarChar, 20).Value = dr["正面印刷纸张损耗_J"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_A", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_A"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_B", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_B"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_C", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_C"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_D", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_D"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_E", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_E"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_F", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_F"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_G", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_G"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_H", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_H"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_I", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_I"].ToString();
                    sqlcom.Parameters.Add("OPPOSITE_PRINTING_J", SqlDbType.VarChar, 20).Value = dr["反面印刷纸张损耗_J"].ToString();
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
