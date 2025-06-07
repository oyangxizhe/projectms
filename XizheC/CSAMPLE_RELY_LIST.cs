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
    public class CSAMPLE_RELY_LIST
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _CHARGE_AUDIT_STATUS;
        public string CHARGE_AUDIT_STATUS
        {
            set { _CHARGE_AUDIT_STATUS = value; }
            get { return _CHARGE_AUDIT_STATUS; }

        }
        private string _SMAL_POP;
        public string SMAL_POP
        {
            set { _SMAL_POP = value; }
            get { return _SMAL_POP; }

        }
        private string _DISPLAY_FRAME;
        public string DISPLAY_FRAME
        {
            set { _DISPLAY_FRAME = value; }
            get { return _DISPLAY_FRAME; }

        }
        private string _ALONE_DEPOSIT;
        public string ALONE_DEPOSIT
        {
            set { _ALONE_DEPOSIT = value; }
            get { return _ALONE_DEPOSIT; }

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
        private string _SRID;
        public string SRID
        {
            set { _SRID = value; }
            get { return _SRID; }

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
        private string _APPOINT_AUDIT_MAKERID;
        public string APPOINT_AUDIT_MAKERID
        {
            set { _APPOINT_AUDIT_MAKERID = value; }
            get { return _APPOINT_AUDIT_MAKERID; }

        }
        private string _IF_PROJECT_AUDIT;
        public string IF_PROJECT_AUDIT
        {
            set { _IF_PROJECT_AUDIT = value; }
            get { return _IF_PROJECT_AUDIT; }

        }

        private string _IF_PAPER_AUDIT;
        public string IF_PAPER_AUDIT
        {
            set { _IF_PAPER_AUDIT = value; }
            get { return _IF_PAPER_AUDIT; }

        }
        private string _IF_ACRYLIC_AUDIT;
        public string IF_ACRYLIC_AUDIT
        {
            set { _IF_ACRYLIC_AUDIT = value; }
            get { return _IF_ACRYLIC_AUDIT; }

        }

        private string _SAMPLE_ID;
        public string SAMPLE_ID
        {
            set { _SAMPLE_ID = value; }
            get { return _SAMPLE_ID; }

        }
        private string _NEED_COUNT;
        public string NEED_COUNT
        {
            set { _NEED_COUNT = value; }
            get { return _NEED_COUNT; }

        }
        private string _NEED_DATE;
        public string NEED_DATE
        {
            set { _NEED_DATE = value; }
            get { return _NEED_DATE; }

        }
        private string _GROUP_TYPE;
        public string GROUP_TYPE
        {
            set { _GROUP_TYPE = value; }
            get { return _GROUP_TYPE; }

        }
        private string _ORDER_DATE;
        public string ORDER_DATE
        {
            set { _ORDER_DATE = value; }
            get { return _ORDER_DATE; }

        }

        private string _QUALITY_LEVAL;
        public string QUALITY_LEVAL
        {
            set { _QUALITY_LEVAL = value; }
            get { return _QUALITY_LEVAL; }

        }

        private string _OWN_OR_PURCHASE;
        public string OWN_OR_PURCHASE
        {
            set { _OWN_OR_PURCHASE = value; }
            get { return _OWN_OR_PURCHASE; }

        }
        private string _PURCHASE_MAKERID;
        public string PURCHASE_MAKERID
        {
            set { _PURCHASE_MAKERID = value; }
            get { return _PURCHASE_MAKERID; }

        }
        private string _IF_WOOD_IRON_AUDIT;
        public string IF_WOOD_IRON_AUDIT
        {
            set { _IF_WOOD_IRON_AUDIT = value; }
            get { return _IF_WOOD_IRON_AUDIT; }

        }
        private string _IF_PURCHASE_AUDIT;
        public string IF_PURCHASE_AUDIT
        {
            set { _IF_PURCHASE_AUDIT = value; }
            get { return _IF_PURCHASE_AUDIT; }

        }
        private string _PROJECT_AUDIT_STATUS;
        public string PROJECT_AUDIT_STATUS
        {
            set { _PROJECT_AUDIT_STATUS = value; }
            get { return _PROJECT_AUDIT_STATUS; }

        }
        private string _PAPER_AUDIT_STATUS;
        public string PAPER_AUDIT_STATUS
        {
            set { _PAPER_AUDIT_STATUS = value; }
            get { return _PAPER_AUDIT_STATUS; }

        }
        private string _ACRYLIC_AUDIT_STATUS;
        public string ACRYLIC_AUDIT_STATUS
        {
            set { _ACRYLIC_AUDIT_STATUS = value; }
            get { return _ACRYLIC_AUDIT_STATUS; }

        }
        private string _WOOD_IRON_AUDIT_STATUS;
        public string WOOD_IRON_AUDIT_STATUS
        {
            set { _WOOD_IRON_AUDIT_STATUS = value; }
            get { return _WOOD_IRON_AUDIT_STATUS; }

        }

        private string _PURCHASE_AUDIT_STATUS;
        public string PURCHASE_AUDIT_STATUS
        {
            set { _PURCHASE_AUDIT_STATUS = value; }
            get { return _PURCHASE_AUDIT_STATUS; }

        }
        private string _PAPER_SELECT;
        public string PAPER_SELECT
        {
            set { _PAPER_SELECT = value; }
            get { return _PAPER_SELECT; }

        }
        private string _DISPLAY_VALUE;
        public string DISPLAY_VALUE
        {
            set { _DISPLAY_VALUE = value; }
            get { return _DISPLAY_VALUE; }

        }
        private string _PROJECT_MAKERID;
        public string PROJECT_MAKERID
        {
            set { _PROJECT_MAKERID = value; }
            get { return _PROJECT_MAKERID; }

        }
        private string _PAPER_MAKERID;
        public string PAPER_MAKERID
        {
            set { _PAPER_MAKERID = value; }
            get { return _PAPER_MAKERID; }

        }
        private string _OWN_REMARK;
        public string OWN_REMARK
        {
            set { _OWN_REMARK = value; }
            get { return _OWN_REMARK; }

        }

        private string _ACRYLIC_MAKERID;
        public string ACRYLIC_MAKERID
        {
            set { _ACRYLIC_MAKERID = value; }
            get { return _ACRYLIC_MAKERID; }

        }

        private string _WOOD_IRON_MAKERID;
        public string WOOD_IRON_MAKERID
        {
            set { _WOOD_IRON_MAKERID = value; }
            get { return _WOOD_IRON_MAKERID; }

        }

        private string _PURCHASE_AUDIT_MAKERID;
        public string PURCHASE_AUDIT_MAKERID
        {
            set { _PURCHASE_AUDIT_MAKERID = value; }
            get { return _PURCHASE_AUDIT_MAKERID; }

        }

        private string _OWN_MAKERID;
        public string OWN_MAKERID
        {
            set { _OWN_MAKERID = value; }
            get { return _OWN_MAKERID; }

        }
        private string _PURCHASE_REMARK;
        public string PURCHASE_REMARK
        {
            set { _PURCHASE_REMARK = value; }
            get { return _PURCHASE_REMARK; }

        }
        private string _OTHER;
        public string OTHER
        {
            set { _OTHER = value; }
            get { return _OTHER; }

        }
        private string _DISPLAY_TYPE;
        public string DISPLAY_TYPE
        {
            set { _DISPLAY_TYPE = value; }
            get { return _DISPLAY_TYPE; }

        }
        #endregion
        CMATERIAL_PRICE cmaterial_price = new CMATERIAL_PRICE();
        int i;
        #region sql
        string setsql = @"
SELECT 
C.PROJECT_ID  AS 项目号,
C.PROJECT_NAME AS 项目名称,
A.SRID AS 打样编号,
A.SAMPLE_ID AS 打样单号,
A.NEED_COUNT AS 需求数量,
A.NEED_DATE AS 需求日期,
A.GROUP_TYPE AS 组别,
A.ORDER_DATE AS 下单日期,
B.MATERIAL_TYPE AS 加工内容,
B.SN AS 项次,
B.TECHNOLOGY AS 工艺,
CASE WHEN A.QUALITY_LEVAL='H' THEN '品质高'
WHEN A.QUALITY_LEVAL='M' THEN '品质中'
WHEN A.QUALITY_LEVAL='LOW' THEN '品质低'
ELSE ''
END AS 品质级别,
CASE WHEN A.OWN_OR_PURCHASE='OWN' THEN '自购'
WHEN A.OWN_OR_PURCHASE='PURCHASE' THEN '采购'
WHEN A.OWN_OR_PURCHASE='OWN_AND_PURCHASE' THEN '自购与采购'
ELSE ''
END 
AS 自购或采购,
(SELECT EMPLOYEE_ID FROM EmployeeInfo WHERE EMID=A.OWN_MAKERID ) AS 自购工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.OWN_MAKERID)  AS 自购,
(SELECT EMPLOYEE_ID FROM EmployeeInfo WHERE EMID=A.PURCHASE_MAKERID ) AS 采购工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PURCHASE_MAKERID)  AS 采购,
A.OWN_REMARK AS 自购说明,
A.OTHER AS 其他事项,
CASE WHEN A.DISPLAY_TYPE='PAPER' THEN '纸品'
WHEN A.DISPLAY_TYPE='METAL' THEN '金属'
WHEN A.DISPLAY_TYPE='WOOD' THEN '木器'
WHEN A.DISPLAY_TYPE='PLASTIC' THEN '塑料'
ELSE ''
END AS 陈列类型,
A.DISPLAY_VALUE AS 陈列数值,
CASE WHEN A.SMAL_POP='Y' AND A.DISPLAY_TYPE='PAPER' THEN '已选'
WHEN A.SMAL_POP='N' AND A.DISPLAY_TYPE='PAPER' THEN '未选'
ELSE ''
END AS 小POP,
CASE WHEN A.DISPLAY_FRAME='Y' AND A.DISPLAY_TYPE='PAPER' THEN '已选'
WHEN A.DISPLAY_FRAME='N' AND A.DISPLAY_TYPE='PAPER' THEN '未选'
ELSE ''
END AS 陈列架,
CASE WHEN A.ALONE_DEPOSIT='Y' AND A.DISPLAY_TYPE='PAPER' THEN '已选'
WHEN A.ALONE_DEPOSIT='N' AND A.DISPLAY_TYPE='PAPER' THEN '未选'
ELSE ''
END AS 堆头,
CASE WHEN A.IF_PAPER_AUDIT='Y' THEN '是'
ELSE '否'
END AS 是否需纸品签核,
CASE WHEN A.IF_ACRYLIC_AUDIT='Y' THEN '是'
ELSE '否'
END AS 是否需亚克力签核,
CASE WHEN A.IF_WOOD_IRON_AUDIT='Y' THEN '是'
ELSE '否'
END AS 是否需木铁签核,
CASE WHEN A.IF_PURCHASE_AUDIT='Y' THEN  '是'
ELSE '否'
END AS 是否需采购签核,
CASE WHEN A.PAPER_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.PAPER_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 纸品签核状态,
CASE WHEN A.ACRYLIC_AUDIT_STATUS='Y' THEN  '已签核'
WHEN A.ACRYLIC_AUDIT_STATUS='N' THEN  '未签核'
ELSE ''
END AS 亚克力签核状态,
CASE WHEN A.WOOD_IRON_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.WOOD_IRON_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 木铁签核状态,
CASE WHEN A.PURCHASE_AUDIT_STATUS='Y' THEN  '已签核'
WHEN A.PURCHASE_AUDIT_STATUS='N' THEN  '未签核'
ELSE ''
END AS 采购签核状态,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.PAPER_MAKERID) AS 纸品工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PAPER_MAKERID)  AS 纸品,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.ACRYLIC_MAKERID) AS 亚克力工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.ACRYLIC_MAKERID)  AS 亚克力,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_MAKERID)  AS 木铁工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_MAKERID)  AS 木铁,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.PURCHASE_AUDIT_MAKERID)AS 采购签核工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PURCHASE_AUDIT_MAKERID)  AS 采购签核,
CASE WHEN A.CHARGE_AUDIT_STATUS='Y' THEN '已完成'
ELSE '进行中' 
END
AS 项目状态,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.CHARGE_MAKERID)  AS 主管审核工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.CHARGE_MAKERID)  AS 主管审核,
A.MakerID AS 制单人工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.MAKERID)  AS 制单人,
A.Date AS 制单日期
FROM SAMPLE_RELY_LIST A
LEFT JOIN SAMPLE_TECHNOLOGY B ON A.SRID=B.SRID
LEFT JOIN PROJECT_INFO C ON SUBSTRING(A.SAMPLE_ID,1,LEN(A.SAMPLE_ID)-3)=C.PROJECT_ID 



";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO SAMPLE_RELY_LIST
(
SRID,
SAMPLE_ID,
NEED_COUNT,
NEED_DATE,
GROUP_TYPE,
ORDER_DATE,
QUALITY_LEVAL,
OWN_OR_PURCHASE,
OWN_MAKERID,
OWN_REMARK,
PURCHASE_MAKERID,
OTHER,
DISPLAY_TYPE,
SMAL_POP,
DISPLAY_FRAME,
ALONE_DEPOSIT,
DISPLAY_VALUE,
IF_PAPER_AUDIT,
IF_ACRYLIC_AUDIT,
IF_WOOD_IRON_AUDIT,
IF_PURCHASE_AUDIT,
PAPER_AUDIT_STATUS,
ACRYLIC_AUDIT_STATUS,
WOOD_IRON_AUDIT_STATUS,
PURCHASE_AUDIT_STATUS,
PAPER_MAKERID,
ACRYLIC_MAKERID,
WOOD_IRON_MAKERID,
PURCHASE_AUDIT_MAKERID,
CHARGE_AUDIT_STATUS,
MakerID,
Date,
YEAR,
MONTH
)
VALUES
(
@SRID,
@SAMPLE_ID,
@NEED_COUNT,
@NEED_DATE,
@GROUP_TYPE,
@ORDER_DATE,
@QUALITY_LEVAL,
@OWN_OR_PURCHASE,
@OWN_MAKERID,
@OWN_REMARK,
@PURCHASE_MAKERID,
@OTHER,
@DISPLAY_TYPE,
@SMAL_POP,
@DISPLAY_FRAME,
@ALONE_DEPOSIT,
@DISPLAY_VALUE,
@IF_PAPER_AUDIT,
@IF_ACRYLIC_AUDIT,
@IF_WOOD_IRON_AUDIT,
@IF_PURCHASE_AUDIT,
@PAPER_AUDIT_STATUS,
@ACRYLIC_AUDIT_STATUS,
@WOOD_IRON_AUDIT_STATUS,
@PURCHASE_AUDIT_STATUS,
@PAPER_MAKERID,
@ACRYLIC_MAKERID,
@WOOD_IRON_MAKERID,
@PURCHASE_AUDIT_MAKERID,
@CHARGE_AUDIT_STATUS,
@MakerID,
@Date,
@YEAR,
@MONTH
)
";
        string setsqlth = @"
INSERT INTO SAMPLE_TECHNOLOGY
(
SRID,
MATERIAL_TYPE,
TECHNOLOGY,
MakerID,
Date,
YEAR,
MONTH,
DAY
)
VALUES
(
@SRID,
@MATERIAL_TYPE,
@TECHNOLOGY,
@MakerID,
@Date,
@YEAR,
@MONTH,
@DAY
)

";

        string setsqlf = @"
UPDATE SAMPLE_RELY_LIST SET
SRID=@SRID,
SAMPLE_ID=@SAMPLE_ID,
NEED_COUNT=@NEED_COUNT,
NEED_DATE=@NEED_DATE,
GROUP_TYPE=@GROUP_TYPE,
ORDER_DATE=@ORDER_DATE,
QUALITY_LEVAL=@QUALITY_LEVAL,
OWN_OR_PURCHASE=@OWN_OR_PURCHASE,
OWN_MAKERID=@OWN_MAKERID,
OWN_REMARK=@OWN_REMARK,
PURCHASE_MAKERID=@PURCHASE_MAKERID,
OTHER=@OTHER,
DISPLAY_TYPE=@DISPLAY_TYPE,
SMAL_POP=@SMAL_POP,
DISPLAY_FRAME=@DISPLAY_FRAME,
ALONE_DEPOSIT=@ALONE_DEPOSIT,
DISPLAY_VALUE=@DISPLAY_VALUE,
IF_PAPER_AUDIT=@IF_PAPER_AUDIT,
IF_ACRYLIC_AUDIT=@IF_ACRYLIC_AUDIT,
IF_WOOD_IRON_AUDIT=@IF_WOOD_IRON_AUDIT,
IF_PURCHASE_AUDIT=@IF_PURCHASE_AUDIT,
PAPER_AUDIT_STATUS=@PAPER_AUDIT_STATUS,
ACRYLIC_AUDIT_STATUS=@ACRYLIC_AUDIT_STATUS,
WOOD_IRON_AUDIT_STATUS=@WOOD_IRON_AUDIT_STATUS,
PURCHASE_AUDIT_STATUS=@PURCHASE_AUDIT_STATUS,
PAPER_MAKERID=@PAPER_MAKERID,
ACRYLIC_MAKERID=@ACRYLIC_MAKERID,
WOOD_IRON_MAKERID=@WOOD_IRON_MAKERID,
PURCHASE_AUDIT_MAKERID=@PURCHASE_AUDIT_MAKERID,
CHARGE_AUDIT_STATUS=@CHARGE_AUDIT_STATUS,
Date=@Date,
YEAR=@YEAR,
MONTH=@MONTH

";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        public CSAMPLE_RELY_LIST()
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
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("打样单号", typeof(string));
            dt.Columns.Add("需求数量", typeof(string));
            dt.Columns.Add("需求日期", typeof(string));
            dt.Columns.Add("组别", typeof(string));
            dt.Columns.Add("下单日期", typeof(string));
            dt.Columns.Add("加工内容", typeof(string));
            dt.Columns.Add("工艺", typeof(string));
            dt.Columns.Add("品质级别", typeof(string));
            dt.Columns.Add("自购或采购", typeof(string));
            dt.Columns.Add("自购", typeof(string));
            dt.Columns.Add("采购", typeof(string));
            dt.Columns.Add("自购说明", typeof(string));
            dt.Columns.Add("其他事项", typeof(string));
            dt.Columns.Add("陈列类型", typeof(string));
            dt.Columns.Add("陈列数值", typeof(string));
            dt.Columns.Add("小POP", typeof(string));
            dt.Columns.Add("陈列架", typeof(string));
            dt.Columns.Add("堆头", typeof(string));
            dt.Columns.Add("样板计费", typeof(string));
            dt.Columns.Add("指定签核人", typeof(string));
            dt.Columns.Add("项目状态", typeof(string));
            dt.Columns.Add("审核确认", typeof(string));
            dt.Columns.Add("审核批注", typeof(string));
            dt.Columns.Add("审核日期", typeof(string));
            dt.Columns.Add("是否需纸品签核", typeof(string));
            dt.Columns.Add("是否需亚克力签核", typeof(string));
            dt.Columns.Add("是否需木铁签核", typeof(string));
            dt.Columns.Add("是否需采购签核", typeof(string));
            dt.Columns.Add("纸品签核状态", typeof(string));
            dt.Columns.Add("亚克力签核状态", typeof(string));
            dt.Columns.Add("木铁签核状态", typeof(string));
            dt.Columns.Add("采购签核状态", typeof(string));
            dt.Columns.Add("纸品", typeof(string));
            dt.Columns.Add("亚克力", typeof(string));
            dt.Columns.Add("木铁", typeof(string));
            dt.Columns.Add("采购签核", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            return dt;
        }

        #endregion
        #region GetTableInfo_SEARCH
        public DataTable GetTableInfo_SEARCH()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("打样单号", typeof(string));
            dt.Columns.Add("组别", typeof(string));
            dt.Columns.Add("需求数量", typeof(string));
            dt.Columns.Add("需求日期", typeof(string));
            dt.Columns.Add("下单日期", typeof(string));
            dt.Columns.Add("签核状态", typeof(string));
            dt.Columns.Add("项目状态", typeof(string));
            dt.Columns.Add("主管审核", typeof(string));
            dt.Columns.Add("打样金额", typeof(string));
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
            string v1 = bc.numYM(10, 4, "0001", "select * from SAMPLE_RELY_LIST_NO", "SRID", "SR");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO SAMPLE_RELY_LIST_NO(SRID,DATE,YEAR,MONTH) VALUES ('" + v1 + "','" + varDate + "','" + year +
                  "','" + month + "')");
              
            }
            return GETID;
        }
        #endregion
        #region GETID_SAMPLE_ID
        public string GETID_SAMPLE_ID(string PROJECT_ID)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numNOYMD(14, 2, "01", "select * from SAMPLE_RELY_LIST WHERE SUBSTRING(SAMPLE_ID,1,11)='" + PROJECT_ID + "' ", "SAMPLE_ID",
                PROJECT_ID+"-"," ORDER BY SAMPLE_ID ASC");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;

            }
            return GETID;
        }
        #endregion
        #region save
        public void save(CheckedListBox clb1,CheckedListBox clb2,CheckedListBox clb3,CheckedListBox clb4,CheckedListBox clb5)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT CHARGE_AUDIT_STATUS FROM SAMPLE_RELY_LIST WHERE SRID='" + SRID + "'");
            if (string.IsNullOrEmpty(SAMPLE_ID))
            {
                ErrowInfo = "打样单号为空执行失败";
                IFExecution_SUCCESS = false;
            }
           else if (!bc.exists("SELECT SRID FROM SAMPLE_RELY_LIST WHERE SRID='" + SRID + "'"))
            {
                SQlcommandE(sqlt);
                TECHNOLOGY(clb1, clb2, clb3, clb4, clb5);
                IFExecution_SUCCESS = true;
            }
            else if (bc.exists("SELECT SRID FROM SAMPLE_RELY_LIST WHERE SRID='" + SRID + "'") && v1 != "Y")
            {
                SQlcommandE(sqlf + " WHERE SRID='" + SRID + "'");
                basec.getcoms("DELETE SAMPLE_TECHNOLOGY WHERE SRID='"+SRID +"'");
                TECHNOLOGY(clb1, clb2, clb3, clb4, clb5);
                IFExecution_SUCCESS = true;
            }
            else
            {
                SQlcommandE(sqlt);
                TECHNOLOGY(clb1, clb2, clb3, clb4, clb5);
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
            sqlcom.Parameters.Add("SRID", SqlDbType.VarChar, 20).Value = SRID;
           
           
                sqlcom.Parameters.Add("SAMPLE_ID", SqlDbType.VarChar, 20).Value = SAMPLE_ID;
            
            sqlcom.Parameters.Add("NEED_COUNT", SqlDbType.VarChar, 20).Value = NEED_COUNT;
            sqlcom.Parameters.Add("NEED_DATE", SqlDbType.VarChar, 20).Value = NEED_DATE;
            sqlcom.Parameters.Add("GROUP_TYPE", SqlDbType.VarChar, 20).Value = GROUP_TYPE;
            sqlcom.Parameters.Add("ORDER_DATE", SqlDbType.VarChar, 20).Value = ORDER_DATE;
            sqlcom.Parameters.Add("QUALITY_LEVAL", SqlDbType.VarChar, 20).Value = QUALITY_LEVAL;
            sqlcom.Parameters.Add("OWN_OR_PURCHASE", SqlDbType.VarChar, 20).Value = OWN_OR_PURCHASE;
            sqlcom.Parameters.Add("OWN_MAKERID", SqlDbType.VarChar, 20).Value = OWN_MAKERID;
            sqlcom.Parameters.Add("OWN_REMARK", SqlDbType.VarChar, 1000).Value = OWN_REMARK;
            sqlcom.Parameters.Add("PURCHASE_MAKERID", SqlDbType.VarChar, 20).Value = PURCHASE_MAKERID;
            sqlcom.Parameters.Add("OTHER", SqlDbType.VarChar, 1000).Value = OTHER;
            sqlcom.Parameters.Add("DISPLAY_TYPE", SqlDbType.VarChar, 20).Value = DISPLAY_TYPE;
            sqlcom.Parameters.Add("SMAL_POP", SqlDbType.VarChar, 20).Value = SMAL_POP;
            sqlcom.Parameters.Add("DISPLAY_FRAME", SqlDbType.VarChar, 20).Value = DISPLAY_FRAME;
            sqlcom.Parameters.Add("ALONE_DEPOSIT", SqlDbType.VarChar, 20).Value = ALONE_DEPOSIT;
            sqlcom.Parameters.Add("DISPLAY_VALUE", SqlDbType.VarChar, 20).Value = DISPLAY_VALUE;
            sqlcom.Parameters.Add("IF_PAPER_AUDIT", SqlDbType.VarChar, 20).Value = IF_PAPER_AUDIT;
            sqlcom.Parameters.Add("IF_ACRYLIC_AUDIT", SqlDbType.VarChar, 20).Value = IF_ACRYLIC_AUDIT;
            sqlcom.Parameters.Add("IF_WOOD_IRON_AUDIT", SqlDbType.VarChar, 20).Value = IF_WOOD_IRON_AUDIT;
            sqlcom.Parameters.Add("IF_PURCHASE_AUDIT", SqlDbType.VarChar, 20).Value = IF_PURCHASE_AUDIT;
            sqlcom.Parameters.Add("PAPER_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = PAPER_AUDIT_STATUS;
            sqlcom.Parameters.Add("ACRYLIC_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = ACRYLIC_AUDIT_STATUS;
            sqlcom.Parameters.Add("WOOD_IRON_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = WOOD_IRON_AUDIT_STATUS;
            sqlcom.Parameters.Add("PURCHASE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = PURCHASE_AUDIT_STATUS;
            sqlcom.Parameters.Add("PAPER_MAKERID", SqlDbType.VarChar, 20).Value = PAPER_MAKERID;
            sqlcom.Parameters.Add("ACRYLIC_MAKERID", SqlDbType.VarChar, 20).Value = ACRYLIC_MAKERID;
            sqlcom.Parameters.Add("WOOD_IRON_MAKERID", SqlDbType.VarChar, 20).Value = WOOD_IRON_MAKERID;
            sqlcom.Parameters.Add("PURCHASE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = PURCHASE_AUDIT_MAKERID;
            sqlcom.Parameters.Add("CHARGE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = CHARGE_AUDIT_STATUS;
            sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region  TECHNOLOGY
        private void TECHNOLOGY(CheckedListBox clb1, CheckedListBox clb2, CheckedListBox clb3, CheckedListBox clb4, CheckedListBox clb5)
        {

            for (i = 0; i < clb1.CheckedItems.Count; i++)
            {
                SQlcommandE_TECHNOLOGY(sqlth, "画面", clb1.CheckedItems[i].ToString());
            }
            for (i = 0; i < clb2.CheckedItems.Count; i++)
            {
                SQlcommandE_TECHNOLOGY(sqlth, "纸品", clb2.CheckedItems[i].ToString());
            }
            for (i = 0; i < clb3.CheckedItems.Count; i++)
            {
                SQlcommandE_TECHNOLOGY(sqlth, "金属", clb3.CheckedItems[i].ToString());
            }
            for (i = 0; i < clb4.CheckedItems.Count; i++)
            {
                SQlcommandE_TECHNOLOGY(sqlth, "亚克力", clb4.CheckedItems[i].ToString());
            }
            for (i = 0; i < clb5.CheckedItems.Count; i++)
            {
                SQlcommandE_TECHNOLOGY(sqlth, "木", clb5.CheckedItems[i].ToString());
            }

        }
        #endregion
        #region  SQlcommandE_TECHNOLOGY
        protected void SQlcommandE_TECHNOLOGY(string sql,string MATERIAL_TYPE,string TECHANOLOGY)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            sqlcon.Open();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@SRID", SqlDbType.VarChar, 20).Value = SRID;
            sqlcom.Parameters.Add("@MATERIAL_TYPE", SqlDbType.VarChar, 20).Value = MATERIAL_TYPE;
            sqlcom.Parameters.Add("@TECHNOLOGY", SqlDbType.VarChar, 20).Value = TECHANOLOGY;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region RETURN_DT
        public DataTable RETURN_DT(DataTable dtt)
        {
            DataTable dt = GetTableInfo();
            DataTable dtx = new DataTable();
            foreach (DataRow dr1 in dtt.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["打样单号"] = dr1["打样单号"].ToString();
                dr["需求数量"] = dr1["需求数量"].ToString();
                dr["需求日期"] = dr1["需求日期"].ToString();
                dr["组别"] = dr1["组别"].ToString();
                dr["下单日期"] = dr1["下单日期"].ToString();
                dr["加工内容"] = dr1["加工内容"].ToString();
                dr["工艺"] = dr1["工艺"].ToString();
                dr["品质级别"] = dr1["品质级别"].ToString();
                dr["自购或采购"] = dr1["自购或采购"].ToString();
                dr["自购"] = dr1["自购"].ToString();
                dr["采购"] = dr1["采购"].ToString();
                dr["自购说明"] = dr1["自购说明"].ToString();
                dr["其他事项"] = dr1["其他事项"].ToString();
                dr["陈列类型"] = dr1["陈列类型"].ToString();
                dr["陈列数值"] = dr1["陈列数值"].ToString();
                dr["小POP"] = dr1["小POP"].ToString();
                dr["陈列架"] = dr1["陈列架"].ToString();
                dr["堆头"] = dr1["堆头"].ToString();
                dr["项目状态"] = dr1["项目状态"].ToString();
                dr["是否需纸品签核"] = dr1["是否需纸品签核"].ToString();
                dr["是否需亚克力签核"] = dr1["是否需亚克力签核"].ToString();
                dr["是否需木铁签核"] = dr1["是否需木铁签核"].ToString();
                dr["是否需采购签核"] = dr1["是否需采购签核"].ToString();
                dr["纸品签核状态"] = dr1["纸品签核状态"].ToString();
                dr["亚克力签核状态"] = dr1["亚克力签核状态"].ToString();
                dr["木铁签核状态"] = dr1["木铁签核状态"].ToString();
                dr["纸品"] = dr1["纸品"].ToString();
                dr["亚克力"] = dr1["亚克力"].ToString();
                dr["木铁"] = dr1["木铁"].ToString();
                dr["采购签核"] = dr1["采购签核"].ToString();
                dr["制单人"] = dr1["制单人"].ToString();
                dt.Rows.Add(dr);
            }
            
            foreach (DataRow dr in dt.Rows)
            {
                
                if (dr["陈列类型"].ToString() == "纸品")
                {
                    dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "纸"));
                    decimal d1 = 0, d2 = 0, d3 = 0;
                    if (dtx.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                        {
                            d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                        {
                            d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                        {
                            d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                        }

                    }
                    if (dr["小POP"].ToString() == "已选" && dr["陈列架"].ToString() == "未选" && dr["堆头"].ToString() == "未选")
                    {
                        if (d1 != 0)
                        {
                            dr["样板计费"] = d1.ToString("0.00");

                        }
                        //MessageBox.Show(dt.Rows[0]["样板计费"].ToString());
                    }
                    else if (dr["小POP"].ToString() == "未选" && dr["陈列架"].ToString() == "已选" && dr["堆头"].ToString() == "未选")
                    {
                        if (d2 != 0)
                        {
                            dr["样板计费"] = d2.ToString("0.00");
                        }
                    }
                    else if (dr["小POP"].ToString() == "未选" && dr["陈列架"].ToString() == "未选" && dr["堆头"].ToString() == "已选")
                    {
                        d3 = d3 * decimal.Parse(dr["陈列数值"].ToString());
                        if (d3 != 0)
                        {
                            dr["样板计费"] = d3.ToString("0.00");
                        }
                    }
                    else if (dr["小POP"].ToString() == "已选" && dr["陈列架"].ToString() == "已选" && dr["堆头"].ToString() == "未选")
                    {
                        d3 = d1 + d2;
                        if (d3 != 0)
                        {
                            dr["样板计费"] = d3.ToString("0.00");
                        }
                    }
                    else if (dr["小POP"].ToString() == "已选" && dr["陈列架"].ToString() == "未选" && dr["堆头"].ToString() == "已选")
                    {
                        d3 = d1 + d3 * decimal.Parse(dr["陈列数值"].ToString());
                        if (d3 != 0)
                        {
                            dr["样板计费"] = d3.ToString("0.00");
                        }
                    }
                    else if (dr["小POP"].ToString() == "未选" && dr["陈列架"].ToString() == "已选" && dr["堆头"].ToString() == "已选")
                    {
                        d3 = d2 + d3 * decimal.Parse(dr["陈列数值"].ToString());
                        if (d3 != 0)
                        {
                            dr["样板计费"] = d3.ToString("0.00");
                        }
                    }
                    else if (dr["小POP"].ToString() == "已选" && dr["陈列架"].ToString() == "已选" && dr["堆头"].ToString() == "已选")
                    {
                        d3 = d1 + d2 + d3 * decimal.Parse(dr["陈列数值"].ToString());
                        if (d3 != 0)
                        {
                            dr["样板计费"] = d3.ToString("0.00");
                        }
                    }
                }
                else if (dr["陈列类型"].ToString() == "金属")
                {
                    dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "金属"));
                    decimal d1 = 0, d2 = 0, d3 = 0;
                    if (dtx.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                        {
                            d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                        {
                            d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                        {
                            d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                        }

                    }
                    if (!string.IsNullOrEmpty(dr["陈列数值"].ToString()))
                    {
                        d2 = d2 * decimal.Parse(dr["陈列数值"].ToString());
                    }
                
                    if (d2 <= d1)
                    {
                        d2 = d1;
                    }
                    else if (d2 >= d3)
                    {
                        d2 = d3;
                    }
                    if (d2 != 0)
                    {
                        dr["样板计费"] = d2.ToString("0.00");
                    }
                }
                else if (dr["陈列类型"].ToString() == "木器")
                {
                    dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "木"));
                    decimal d1 = 0, d2 = 0, d3 = 0;
                    if (dtx.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                        {
                            d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                        {
                            d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                        {
                            d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                        }

                    }

                    if (!string.IsNullOrEmpty(dr["陈列数值"].ToString()))
                    {
                        d2 = d2 * decimal.Parse(dr["陈列数值"].ToString());
                    }
                    if (d2 <= d1)
                    {
                        d2 = d1;
                    }
                    else if (d2 >= d3)
                    {
                        d2 = d3;
                    }
                    if (d2 != 0)
                    {
                        dr["样板计费"] = d2.ToString("0.00");
                    }
                }
                else if (dr["陈列类型"].ToString() == "塑料")
                {
                    dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "亚克力"));
                    decimal d1 = 0, d2 = 0, d3 = 0;
                    if (dtx.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                        {
                            d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                        {
                            d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                        {
                            d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                        }

                    }

                    if (!string.IsNullOrEmpty(dr["陈列数值"].ToString()))
                    {
                        d2 = d2 * decimal.Parse(dr["陈列数值"].ToString());
                    }
                    if (d2 <= d1)
                    {
                        d2 = d1;
                    }
                    else if (d2 >= d3)
                    {
                        d2 = d3;
                    }
                    if (d2 != 0)
                    {
                        dr["样板计费"] = d2.ToString("0.00");
                    }
                }   
                       
                    
            }
        
            return dt;
        }
        #endregion
        #region RETURN_SEARCH
        public DataTable RETURN_DT_SEARCH(DataTable dtt)
        {
            DataTable dt = GetTableInfo_SEARCH();
            DataTable dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtt, "打样单号");
            i = 1;
            if (dtx.Rows.Count > 0)
            {
                
                foreach (DataRow dr2 in dtx.Rows)
                {
                    DataTable dtx1 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "打样单号='"+dr2["VALUE"].ToString()+"'");
                    if (dtx1.Rows.Count > 0)
                    {
                        DataRow dr = dt.NewRow();
                        dr["序号"] = i;
                        dr["项目名称"] = dtx1.Rows [0]["项目名称"].ToString();
                        dr["打样单号"] = dtx1.Rows [0]["打样单号"].ToString();
                        dr["组别"] = dtx1.Rows [0]["组别"].ToString();
                        dr["需求数量"] = dtx1.Rows [0]["需求数量"].ToString();
                        dr["需求日期"] = dtx1.Rows [0]["需求日期"].ToString();
                        dr["下单日期"] = dtx1.Rows [0]["下单日期"].ToString();
                        if (JUAGE_IF_AUDIT_END(dtx1.Rows [0]["打样编号"].ToString()))
                        {
                            dr["签核状态"] = "待签核";
                        }
                        else
                        {
                            dr["签核状态"] = "签核完成";
                        }
                        dr["项目状态"] = dtx1.Rows [0]["项目状态"].ToString();
                        dr["主管审核"] = dtx1.Rows [0]["主管审核"].ToString();
                        DataTable dtx4 =this.RETURN_DT(bc.getdt(this.sql + " WHERE A.SAMPLE_ID='" + dr2["VALUE"].ToString() +
                                          "'"));

                        dr["打样金额"] = dtx4.Rows[0]["样板计费"].ToString();
                        dt.Rows.Add(dr);
                        i = i + 1;
                    }
          
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
            Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
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
            Excel.CheckBoxes ckbs = (Excel.CheckBoxes)worksheet.CheckBoxes(Type.Missing);

           /* Excel.CheckBox cbts = (Excel.CheckBox)ckbs.Item(1);//自采
            MessageBox.Show(cbts.Value + " " + cbts.Caption);
            Excel.CheckBox cbts1 = (Excel.CheckBox)ckbs.Item(2);//采购
            MessageBox.Show(cbts1.Value + " " + cbts1.Caption);
            Excel.CheckBox cbts3 = (Excel.CheckBox)ckbs.Item(3);//小POP
            MessageBox.Show(cbts3.Value + " " + cbts3.Caption);
            Excel.CheckBox cbts4 = (Excel.CheckBox)ckbs.Item(4);//陈列架
            MessageBox.Show(cbts4.Value + " " + cbts4.Caption);

            Excel.CheckBox cbts5 = (Excel.CheckBox)ckbs.Item(5);//堆头
            MessageBox.Show(cbts5.Value + " " + cbts5.Caption);
            Excel.CheckBox cbts6 = (Excel.CheckBox)ckbs.Item(6);//纸品
            MessageBox.Show(cbts6.Value + " " + cbts6.Caption);

            Excel.CheckBox cbts7 = (Excel.CheckBox)ckbs.Item(7);//金属
            MessageBox.Show(cbts7.Value + " " + cbts7.Caption);
            Excel.CheckBox cbts8 = (Excel.CheckBox)ckbs.Item(8);//木器
            MessageBox.Show(cbts8.Value + " " + cbts8.Caption);
            Excel.CheckBox cbts9 = (Excel.CheckBox)ckbs.Item(9);//塑料
            MessageBox.Show(cbts9.Value + " " + cbts9.Caption);*/
            worksheet.Cells[6,"E"]= dt.Rows[0]["打样单号"].ToString();
            worksheet.Cells[6, "W"] = dt.Rows[0]["项目名称"].ToString();
            worksheet.Cells[7, "E"] = dt.Rows[0]["需求数量"].ToString();
            worksheet.Cells[7, "W"] = dt.Rows[0]["需求日期"].ToString();
            worksheet.Cells[8, "E"] = dt.Rows[0]["组别"].ToString();
            worksheet.Cells[8, "W"] = dt.Rows[0]["下单日期"].ToString();

            /*worksheet.Cells[6 + 0, 7] = dt.Rows[0]["加工内容"].ToString();
            worksheet.Cells[6 + 0, 8] = dt.Rows[0]["工艺"].ToString();

            worksheet.Cells[6 + 0, 9] = dt.Rows[0]["品质级别"].ToString();*/

            if (dt.Rows[0]["自购或采购"].ToString() == "自购")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(1);
                cbt.Value = true;
              
            }
            else if (dt.Rows[0]["自购或采购"].ToString() == "采购")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(2);
                cbt.Value = true;

            }
            else if (dt.Rows[0]["自购或采购"].ToString() == "自购与采购")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(1);
                cbt.Value = true;
                cbt = (Excel.CheckBox)ckbs.Item(2);
                cbt.Value = true;

            }
            //worksheet.Cells[6 + 0, 10] = dt.Rows[0]["自购或采购"].ToString();
            worksheet.Cells[17, "I"] = dt.Rows[0]["自购"].ToString();
            worksheet.Cells[17, "S"] = dt.Rows[0]["采购"].ToString();
            worksheet.Cells[18, "D"] = dt.Rows[0]["自购说明"].ToString();
            worksheet.Cells[26, "D"] = dt.Rows[0]["其他事项"].ToString();

            worksheet.Cells[34, "E"] = dt.Rows[0]["制单人"].ToString();
            worksheet.Cells[35, "E"] = dt.Rows[0]["纸品"].ToString();
            worksheet.Cells[35, "M"] = dt.Rows[0]["亚克力"].ToString();
            worksheet.Cells[35, "U"] = dt.Rows[0]["木铁"].ToString();
            worksheet.Cells[35, "AC"] = dt.Rows[0]["采购签核"].ToString();
   
            if (dt.Rows[0]["小POP"].ToString() == "已选")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(3);
                cbt.Value = true;
            }
            if (dt.Rows[0]["陈列架"].ToString() == "已选")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(4);
                cbt.Value = true;
            }
            if (dt.Rows[0]["堆头"].ToString() == "已选")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(5);
                cbt.Value = true;
                worksheet.Cells[38, "AA"] = dt.Rows[0]["陈列数值"].ToString();
            }
     
            if (dt.Rows[0]["陈列类型"].ToString() == "纸品")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(6);
                cbt.Value = true;
            }
            else  if (dt.Rows[0]["陈列类型"].ToString() == "金属")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(7);
                cbt.Value = true;
                worksheet.Cells[39, "H"] = dt.Rows[0]["陈列数值"].ToString();

            }
            else if (dt.Rows[0]["陈列类型"].ToString() == "塑料")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(9);
                cbt.Value = true;
                worksheet.Cells[40, "H"] = dt.Rows[0]["陈列数值"].ToString();

            }
            else if (dt.Rows[0]["陈列类型"].ToString() == "木器")
            {
                Excel.CheckBox cbt = (Excel.CheckBox)ckbs.Item(8);
                cbt.Value = true;
                worksheet.Cells[41, "H"] = dt.Rows[0]["陈列数值"].ToString();

            }
            //MessageBox.Show(dt.Rows[0]["样板计费"].ToString());
            worksheet.Cells[41, "Z"] = dt.Rows[0]["样板计费"].ToString();
            //worksheet.get_Range("A41", "AA41").Font.Bold = true; //设置字体为粗体
            DataTable dtx = bc.RETURN_NOHAVE_REPEAT_DT(dt, "加工内容");
            if (dtx.Rows.Count > 0)
            {
                for (i = 0; i < dtx.Rows.Count; i++)
                {
                    int n = 0, j = 0;
                    worksheet.Cells[i + 10, "D"] = dtx.Rows[i]["VALUE"].ToString();

                    DataTable dtx1 = bc.GET_DT_TO_DV_TO_DT(dt, "", "加工内容='" + dtx.Rows[i]["VALUE"].ToString() + "'");
                    if (dtx1.Rows.Count > 0)
                    {
                        n = j + 7;
                        for (j = 0; j < dtx1.Rows.Count; j++)
                        {
                            worksheet.Cells[i + 10, n] = dtx1.Rows[j]["工艺"].ToString();
                            n = n + 5;
                        }
                    }
                }
            }
     
            //workbook.Save();
            //bc.csharpExcelPrint(sfdg.FileName);
            /*application.Quit();
            worksheet = null;
            workbook = null;
            application = null;
            GC.Collect();*/
        }
        #endregion
        #region JUAGE_IF_AUDIT_END SRID
        public bool JUAGE_IF_AUDIT_END(string SRID)
        {
            bool b = false;
            DataTable dt = new DataTable();
            dt=bc.getdt(sql+string .Format (" WHERE A.SRID='{0}'",SRID ));
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["是否需纸品签核"].ToString() == "是" && dt.Rows[0]["纸品签核状态"].ToString() == "未签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需亚克力签核"].ToString() == "是" && dt.Rows[0]["亚克力签核状态"].ToString() == "未签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁签核"].ToString() == "是" && dt.Rows[0]["木铁签核状态"].ToString() == "未签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需采购签核"].ToString() == "是" && dt.Rows[0]["采购签核状态"].ToString() == "未签核")
                {
                    b = true;
                }
            }
            else
            {
                b = true;
            }
            return b;

        }
        #endregion
    }
}
