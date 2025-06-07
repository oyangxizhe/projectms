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
using System.Net;
using System.IO;


namespace XizheC
{
    public class CPN_PRODUCTION_INSTRUCTIONS
    {
        basec bc = new basec();
        CNOTICE_LIST cnotice_list = new CNOTICE_LIST();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }
        }
        private string _SUBMIT_MAKERID;
        public string SUBMIT_MAKERID
        {
            set { _SUBMIT_MAKERID = value; }
            get { return _SUBMIT_MAKERID; }

        }
        private string _PIID;
        public string PIID
        {
            set { _PIID = value; }
            get { return _PIID; }
        }
        private string _IF_AUDIT_PRICE;
        public string IF_AUDIT_PRICE
        {
            set { _IF_AUDIT_PRICE = value; }
            get { return _IF_AUDIT_PRICE; }
        }
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }

        }
        private string _IF_SUBMIT;
        public string IF_SUBMIT
        {
            set { _IF_SUBMIT = value; }
            get { return _IF_SUBMIT; }

        }
        private string _ORDER_ID;
        public string ORDER_ID
        {
            set { _ORDER_ID = value; }
            get { return _ORDER_ID; }

        }
        private string _PRODUCTION_COUNT;
        public string PRODUCTION_COUNT
        {
            set { _PRODUCTION_COUNT = value; }
            get { return _PRODUCTION_COUNT; }

        }
        private string _HAVE_TAX_UNIT_PRICE;
        public string HAVE_TAX_UNIT_PRICE
        {
            set { _HAVE_TAX_UNIT_PRICE = value; }
            get { return _HAVE_TAX_UNIT_PRICE; }

        }
        private string _MATTERS_NEEDING_ATTENTION;
        public string MATTERS_NEEDING_ATTENTION
        {
            set { _MATTERS_NEEDING_ATTENTION = value; }
            get { return _MATTERS_NEEDING_ATTENTION; }

        }
        private string _PACKING_METHOD;
        public string PACKING_METHOD
        {
            set { _PACKING_METHOD = value; }
            get { return _PACKING_METHOD; }

        }
        private string _OUTSIDE_BOX_MATERIAL;
        public string OUTSIDE_BOX_MATERIAL
        {
            set { _OUTSIDE_BOX_MATERIAL = value; }
            get { return _OUTSIDE_BOX_MATERIAL; }

        }
        private string _INSTRUCTION_REQUIRE;
        public string INSTRUCTION_REQUIRE
        {
            set { _INSTRUCTION_REQUIRE = value; }
            get { return _INSTRUCTION_REQUIRE; }

        }
        private string _INSTRUCTION_SIZE;
        public string INSTRUCTION_SIZE
        {
            set { _INSTRUCTION_SIZE = value; }
            get { return _INSTRUCTION_SIZE; }

        }
        private string _OUTSIDE_BOX_HEIGHT;
        public string OUTSIDE_BOX_HEIGHT
        {
            set { _OUTSIDE_BOX_HEIGHT = value; }
            get { return _OUTSIDE_BOX_HEIGHT; }

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
        private string _OUTSIDE_BOX_LONG;
        public string OUTSIDE_BOX_LONG
        {
            set { _OUTSIDE_BOX_LONG = value; }
            get { return _OUTSIDE_BOX_LONG; }

        }
        private string _OUTSIDE_BOX_WEIGHT;
        public string OUTSIDE_BOX_WEIGHT
        {
            set { _OUTSIDE_BOX_WEIGHT = value; }
            get { return _OUTSIDE_BOX_WEIGHT; }

        }
        private string _PAPER_PRODUCTION_AUDIT_STATUS;
        public string PAPER_PRODUCTION_AUDIT_STATUS
        {
            set { _PAPER_PRODUCTION_AUDIT_STATUS = value; }
            get { return _PAPER_PRODUCTION_AUDIT_STATUS; }

        }
        private string _WOOD_IRON_PRODUCTION_AUDIT_MAKERID;
        public string WOOD_IRON_PRODUCTION_AUDIT_MAKERID
        {
            set { _WOOD_IRON_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _WOOD_IRON_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID;
        public string APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID
        {
            set { _APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID;
        public string APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID
        {
            set { _APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID;
        public string APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID
        {
            set { _APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _APPOINT_PAPER_PLAN_AUDIT_MAKERID;
        public string APPOINT_PAPER_PLAN_AUDIT_MAKERID
        {
            set { _APPOINT_PAPER_PLAN_AUDIT_MAKERID = value; }
            get { return _APPOINT_PAPER_PLAN_AUDIT_MAKERID; }

        }
        private string _APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID;
        public string APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID
        {
            set { _APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID = value; }
            get { return _APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID; }
        }
         private string _APPOINT_STRUCTURE_AUDIT_MAKERID;
        public string APPOINT_STRUCTURE_AUDIT_MAKERID
        {
            set { _APPOINT_STRUCTURE_AUDIT_MAKERID = value; }
            get { return _APPOINT_STRUCTURE_AUDIT_MAKERID; }

        }
        private string _APPOINT_PLANE_AUDIT_MAKERID;
        public string APPOINT_PLANE_AUDIT_MAKERID
        {
            set { _APPOINT_PLANE_AUDIT_MAKERID = value; }
            get { return _APPOINT_PLANE_AUDIT_MAKERID; }

        }
        private string _APPOINT_PAPER_PURCHASE_AUDIT_MAKERID;
        public string APPOINT_PAPER_PURCHASE_AUDIT_MAKERID
        {
            set { _APPOINT_PAPER_PURCHASE_AUDIT_MAKERID = value; }
            get { return _APPOINT_PAPER_PURCHASE_AUDIT_MAKERID; }

        }
        private string _APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID;
        public string APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID
        {
            set { _APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID = value; }
            get { return _APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private string _PNID;
        public string PNID
        {
            set { _PNID = value; }
            get { return _PNID; }

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
        private string _WAREID;
        public string WAREID
        {
            set { _WAREID = value; }
            get { return _WAREID; }

        }
        private string _EDIT_DATE;
        public string EDIT_DATE
        {
            set { _EDIT_DATE = value; }
            get { return _EDIT_DATE; }

        }

        private string _EDIT_TIMES;
        public string EDIT_TIMES
        {
            set { _EDIT_TIMES = value; }
            get { return _EDIT_TIMES; }

        }
        private string _ORDER_TYPE;
        public string ORDER_TYPE
        {
            set { _ORDER_TYPE = value; }
            get { return _ORDER_TYPE; }

        }
        private string _ORDER_DATE;
        public string ORDER_DATE
        {
            set { _ORDER_DATE = value; }
            get { return _ORDER_DATE; }

        }


        private string _PAPER_PLAN_AUDIT_STATUS;
        public string PAPER_PLAN_AUDIT_STATUS
        {
            set { _PAPER_PLAN_AUDIT_STATUS = value; }
            get { return _PAPER_PLAN_AUDIT_STATUS; }

        }
        private string _PAPER_PRODUCTION_AUDIT_MAKERID;
        public string PAPER_PRODUCTION_AUDIT_MAKERID
        {
            set { _PAPER_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _PAPER_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _PAPER_PLAN_AUDIT_MAKERID;
        public string PAPER_PLAN_AUDIT_MAKERID
        {
            set { _PAPER_PLAN_AUDIT_MAKERID = value; }
            get { return _PAPER_PLAN_AUDIT_MAKERID; }

        }
        private string _WOOD_IRON_PLAN_AUDIT_MAKERID;
        public string WOOD_IRON_PLAN_AUDIT_MAKERID
        {
            set { _WOOD_IRON_PLAN_AUDIT_MAKERID = value; }
            get { return _WOOD_IRON_PLAN_AUDIT_MAKERID; }

        }
        private string _ACRYLIC_PRODUCTION_AUDIT_MAKERID;
        public string ACRYLIC_PRODUCTION_AUDIT_MAKERID
        {
            set { _ACRYLIC_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _ACRYLIC_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _OUTSIDE_BOX_WIDTH;
        public string OUTSIDE_BOX_WIDTH
        {
            set { _OUTSIDE_BOX_WIDTH = value; }
            get { return _OUTSIDE_BOX_WIDTH; }

        }
        private string _IF_PAPER_PRODUCTION_AUDIT;
        public string IF_PAPER_PRODUCTION_AUDIT
        {
            set { _IF_PAPER_PRODUCTION_AUDIT = value; }
            get { return _IF_PAPER_PRODUCTION_AUDIT; }

        }
        private string _IF_WOOD_IRON_PRODUCTION_AUDIT;
        public string IF_WOOD_IRON_PRODUCTION_AUDIT
        {
            set { _IF_WOOD_IRON_PRODUCTION_AUDIT = value; }
            get { return _IF_WOOD_IRON_PRODUCTION_AUDIT; }

        }
        private string _IF_WOOD_IRON_AUDIT;
        public string IF_WOOD_IRON_AUDIT
        {
            set { _IF_WOOD_IRON_AUDIT = value; }
            get { return _IF_WOOD_IRON_AUDIT; }
        }
        private string _IF_PAPER_PLAN_AUDIT;
        public string IF_PAPER_PLAN_AUDIT
        {
            set { _IF_PAPER_PLAN_AUDIT = value; }
            get { return _IF_PAPER_PLAN_AUDIT; }
        }
        private string _IF_STRUCTURE_AUDIT;
        public string IF_STRUCTURE_AUDIT
        {
            set { _IF_STRUCTURE_AUDIT = value; }
            get { return _IF_STRUCTURE_AUDIT; }
        }
        private string _IF_PLANE_AUDIT;
        public string IF_PLANE_AUDIT
        {
            set { _IF_PLANE_AUDIT = value; }
            get { return _IF_PLANE_AUDIT; }
        }
        private string _IF_ACRYLIC_PRODUCTION_AUDIT;
        public string IF_ACRYLIC_PRODUCTION_AUDIT
        {
            set { _IF_ACRYLIC_PRODUCTION_AUDIT = value; }
            get { return _IF_ACRYLIC_PRODUCTION_AUDIT; }

        }
        private string _IF_PAPER_PURCHASE_AUDIT;
        public string IF_PAPER_PURCHASE_AUDIT
        {
            set { _IF_PAPER_PURCHASE_AUDIT = value; }
            get { return _IF_PAPER_PURCHASE_AUDIT; }
        }
        private string _IF_WOOD_IRON_PLAN_AUDIT;
        public string IF_WOOD_IRON_PLAN_AUDIT
        {
            set { _IF_WOOD_IRON_PLAN_AUDIT = value; }
            get { return _IF_WOOD_IRON_PLAN_AUDIT; }

        }
        private string _IF_WOOD_IRON_PURCHASE_AUDIT;
        public string IF_WOOD_IRON_PURCHASE_AUDIT
        {
            set { _IF_WOOD_IRON_PURCHASE_AUDIT = value; }
            get { return _IF_WOOD_IRON_PURCHASE_AUDIT; }

        }
        private string _WOOD_IRON_PRODUCTION_AUDIT_STATUS;
        public string WOOD_IRON_PRODUCTION_AUDIT_STATUS
        {
            set { _WOOD_IRON_PRODUCTION_AUDIT_STATUS = value; }
            get { return _WOOD_IRON_PRODUCTION_AUDIT_STATUS; }

        }

        private string _ACRYLIC_PRODUCTION_AUDIT_STATUS;
        public string ACRYLIC_PRODUCTION_AUDIT_STATUS
        {
            set { _ACRYLIC_PRODUCTION_AUDIT_STATUS = value; }
            get { return _ACRYLIC_PRODUCTION_AUDIT_STATUS; }

        }
        private string _WOOD_IRON_PLAN_AUDIT_STATUS;
        public string WOOD_IRON_PLAN_AUDIT_STATUS
        {
            set { _WOOD_IRON_PLAN_AUDIT_STATUS = value; }
            get { return _WOOD_IRON_PLAN_AUDIT_STATUS; }

        }
   
        private string _DELIVERY_DATE;
        public string DELIVERY_DATE
        {
            set { _DELIVERY_DATE = value; }
            get { return _DELIVERY_DATE; }

        }
        private string _DELIVERY_PLACE;
        public string DELIVERY_PLACE
        {
            set { _DELIVERY_PLACE = value; }
            get { return _DELIVERY_PLACE; }

        }
        private string _DELIVERY_BATCH;
        public string DELIVERY_BATCH
        {
            set { _DELIVERY_BATCH = value; }
            get { return _DELIVERY_BATCH; }

        }
         private string _INSTURCTION_REQUIRE;
        public string  INSTURCTION_REQUIRE
        {
            set { _INSTURCTION_REQUIRE = value; }
            get { return _INSTURCTION_REQUIRE; }

        }
        #endregion
        CMATERIAL_PRICE cmaterial_price = new CMATERIAL_PRICE();
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        CNO_PAPER_OFFER cno_paper_offer = new CNO_PAPER_OFFER();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        int i;
        #region sql
        string setsql = @"
SELECT
A.PNID AS 编号,
CASE WHEN SUBSTRING(A.PFID,1,1)='P' THEN (SELECT OFFER_ID FROM PRINTING_OFFER_MST WHERE PFID=A.PFID)
ELSE  (SELECT OFFER_ID FROM NO_PAPER_OFFER_DET WHERE NPKEY=A.PFID)
END
AS 报价编号,
A.ORDER_ID AS 订单编号,
A.IF_AUDIT_PRICE AS 报价,
A.PRODUCTION_COUNT AS 生产数量,
A.HAVE_TAX_UNIT_PRICE AS 含税单价,
CASE WHEN A.HAVE_TAX_UNIT_PRICE IS NOT NULL AND A.HAVE_TAX_UNIT_PRICE<>'' THEN A.HAVE_TAX_UNIT_PRICE*A.PRODUCTION_COUNT
ELSE 0
END  AS 订单金额,
A.EDIT_TIMES AS 修改次数,
A.WAREID AS 品号,
A.ORDER_DATE AS 下单日期,
A.EDIT_DATE AS 修改日期,
A.DELIVERY_DATE AS 交货日期,
A.DELIVERY_BATCH AS 交货批次,
A.DELIVERY_PLACE AS 交货地点,
CASE WHEN A.AUDIT_STATUS='SEND' THEN '已送签'
WHEN A.AUDIT_STATUS='END' THEN '签核完毕'
ELSE '未送签'
END AS 签核状态,
CASE WHEN A.IF_PAPER_PRODUCTION_AUDIT='Y' THEN '是'
WHEN A.IF_PAPER_PRODUCTION_AUDIT='N' THEN '否'
ELSE ''
END AS 是否需纸品生产签核,
CASE WHEN A.IF_WOOD_IRON_PRODUCTION_AUDIT='Y' THEN  '是'
WHEN A.IF_WOOD_IRON_PRODUCTION_AUDIT='N' THEN  '否'
ELSE ''
END AS 是否需木铁生产签核,
CASE WHEN A.IF_ACRYLIC_PRODUCTION_AUDIT='Y' THEN '是'
WHEN A.IF_ACRYLIC_PRODUCTION_AUDIT='N' THEN '否'
ELSE ''
END AS 是否需亚克力生产签核,
CASE WHEN A.IF_PAPER_PLAN_AUDIT='Y' THEN '是'
WHEN A.IF_PAPER_PLAN_AUDIT='N' THEN '否'
ELSE ''
END AS 是否需纸品计划签核,
CASE WHEN A.IF_WOOD_IRON_PLAN_AUDIT='Y' THEN  '是'
WHEN A.IF_WOOD_IRON_PLAN_AUDIT='N' THEN  '否'
ELSE ''
END AS 是否需木铁计划签核,
CASE WHEN A.IF_STRUCTURE_AUDIT='Y' THEN  '是'
WHEN A.IF_STRUCTURE_AUDIT='N' THEN  '否'
ELSE ''
END AS 是否需结构设计签核,
CASE WHEN A.IF_PLANE_AUDIT='Y' THEN '是'
WHEN A.IF_PLANE_AUDIT='N' THEN '否'
ELSE ''
END AS 是否需平面设计签核,
CASE WHEN A.IF_PAPER_PURCHASE_AUDIT='Y' THEN '是'
WHEN A.IF_PAPER_PURCHASE_AUDIT='N' THEN '否'
ELSE ''
END AS 是否需纸品采购签核,
CASE WHEN A.IF_WOOD_IRON_PURCHASE_AUDIT='Y' THEN  '是'
WHEN A.IF_WOOD_IRON_PURCHASE_AUDIT='N' THEN  '否'
ELSE ''
END AS 是否需木铁采购签核,
CASE WHEN A.PAPER_PRODUCTION_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.PAPER_PRODUCTION_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 纸品生产签核状态,
CASE WHEN A.WOOD_IRON_PRODUCTION_AUDIT_STATUS='Y' THEN  '已签核'
WHEN A.WOOD_IRON_PRODUCTION_AUDIT_STATUS='N' THEN  '未签核'
ELSE ''
END AS 木铁生产签核状态,
CASE WHEN A.ACRYLIC_PRODUCTION_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.ACRYLIC_PRODUCTION_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 亚克力生产签核状态,
CASE WHEN A.PAPER_PLAN_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.PAPER_PLAN_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 纸品计划签核状态,
CASE WHEN A.WOOD_IRON_PLAN_AUDIT_STATUS='Y' THEN  '已签核'
WHEN A.WOOD_IRON_PLAN_AUDIT_STATUS='N' THEN  '未签核'
ELSE ''
END AS 木铁计划签核状态,
CASE WHEN A.STRUCTURE_AUDIT_STATUS='Y' THEN  '已签核'
WHEN A.STRUCTURE_AUDIT_STATUS='N' THEN  '未签核'
ELSE ''
END AS 结构设计签核状态,
CASE WHEN A.PLANE_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.PLANE_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 平面设计签核状态,
CASE WHEN A.PAPER_PURCHASE_AUDIT_STATUS='Y' THEN '已签核'
WHEN A.PAPER_PURCHASE_AUDIT_STATUS='N' THEN '未签核'
ELSE ''
END AS 纸品采购签核状态,
CASE WHEN A.WOOD_IRON_PURCHASE_AUDIT_STATUS='Y' THEN  '已签核'
WHEN A.WOOD_IRON_PURCHASE_AUDIT_STATUS='N' THEN  '未签核'
ELSE ''
END AS 木铁采购签核状态,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.PAPER_PRODUCTION_AUDIT_MAKERID) AS 纸品生产工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PAPER_PRODUCTION_AUDIT_MAKERID)  AS 纸品生产,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PRODUCTION_AUDIT_MAKERID) AS 木铁生产工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PRODUCTION_AUDIT_MAKERID)  AS 木铁生产,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.ACRYLIC_PRODUCTION_AUDIT_MAKERID)  AS 亚克力生产工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.ACRYLIC_PRODUCTION_AUDIT_MAKERID)  AS 亚克力生产,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.PAPER_PLAN_AUDIT_MAKERID)  AS 纸品计划工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PAPER_PLAN_AUDIT_MAKERID)  AS 纸品计划,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PLAN_AUDIT_MAKERID)AS 木铁计划工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PLAN_AUDIT_MAKERID)  AS 木铁计划,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_AUDIT_MAKERID) AS 结构设计工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.STRUCTURE_AUDIT_MAKERID)  AS 结构设计,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.PLANE_AUDIT_MAKERID)  AS 平面设计工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PLANE_AUDIT_MAKERID)  AS 平面设计,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.PAPER_PURCHASE_AUDIT_MAKERID)  AS 纸品采购工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.PAPER_PURCHASE_AUDIT_MAKERID)  AS 纸品采购,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PURCHASE_AUDIT_MAKERID)AS 木铁采购工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PURCHASE_AUDIT_MAKERID)  AS 木铁采购,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID) AS 指定纸品生产工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID)  AS 指定纸品生产,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID) AS 指定木铁生产工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID)  AS 指定木铁生产,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID)  AS 指定亚克力生产工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID)  AS 指定亚克力生产,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PAPER_PLAN_AUDIT_MAKERID)  AS 指定纸品计划工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PAPER_PLAN_AUDIT_MAKERID)  AS 指定纸品计划,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID)AS 指定木铁计划工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID)  AS 指定木铁计划,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_STRUCTURE_AUDIT_MAKERID) AS 指定结构设计工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_STRUCTURE_AUDIT_MAKERID)  AS 指定结构设计,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PLANE_AUDIT_MAKERID)  AS 指定平面设计工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PLANE_AUDIT_MAKERID)  AS 指定平面设计,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PAPER_PURCHASE_AUDIT_MAKERID)  AS 指定纸品采购工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_PAPER_PURCHASE_AUDIT_MAKERID)  AS 指定纸品采购,
(SELECT EMPLOYEE_ID  FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID)AS 指定木铁采购工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID)  AS 指定木铁采购,
CASE WHEN A.IF_SUBMIT='Y' THEN '已提交'
ELSE '未提交' 
END
AS 是否提交,
A.MakerID AS 制单人编号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.MAKERID)  AS 制单人,
A.Date AS 制单日期,
CASE WHEN C.CName IS NULL THEN A.CNAME 
ELSE 
C.CNAME 
END 
AS 客户名称,
CASE WHEN B.AE_MAKERID_1 IS NULL THEN A.AE 
ELSE 
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.AE_MAKERID_1) 
END AS AE01,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.AE_MAKERID_2) AS AE02,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.AE_MAKERID_3) AS AE03,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.PLANE_MAKERID_1) AS 平面01,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.PLANE_MAKERID_2) AS 平面02,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.PLANE_MAKERID_3) AS 平面03,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.STRUCTURE_MAKERID_1) AS 结构01,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.STRUCTURE_MAKERID_2) AS 结构02,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.STRUCTURE_MAKERID_3) AS 结构03,
A.ORDER_TYPE AS 订单类型,
A.PACKING_METHOD AS 包装方式,
A.OUTSIDE_BOX_MATERIAL AS 外箱材质,
A.OUTSIDE_BOX_LONG AS 长,
A.OUTSIDE_BOX_WIDTH AS 宽,
A.OUTSIDE_BOX_HEIGHT AS 高,
A.OUTSIDE_BOX_WEIGHT AS 外箱重量,
A.INSTRUCTION_SIZE AS 说明书尺寸,
A.INSTRUCTION_REQUIRE AS 说明书要求,
A.MATTERS_NEEDING_ATTENTION AS 生产注意事项,
B.PROJECT_ID AS 项目号,
B.PROJECT_NAME AS 项目名称,
CASE WHEN B.BRAND IS NULL THEN A.BRAND 
ELSE 
B.BRAND
END  
AS 品牌,
A.SUBMIT_MAKERID AS 提交人工号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.SUBMIT_MAKERID) AS 提交人,
A.CNAME AS 导入的客户,
A.BRAND AS 导入的品牌,
A.AE AS 导入的AE
FROM PN_PRODUCTION_INSTRUCTIONS A
LEFT JOIN PROJECT_INFO B ON A.PIID=B.PIID
LEFT JOIN CUSTOMERINFO_MST C ON B.CUID=C.CUID






";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PN_PRODUCTION_INSTRUCTIONS
(
PNID,
PFID,
PIID,
ORDER_DATE,
EDIT_DATE,
EDIT_TIMES,
ORDER_ID,
IF_AUDIT_PRICE,
WAREID,
PRODUCTION_COUNT,
HAVE_TAX_UNIT_PRICE,
DELIVERY_DATE,
DELIVERY_BATCH,
DELIVERY_PLACE,
ORDER_TYPE,
PACKING_METHOD,
OUTSIDE_BOX_MATERIAL,
OUTSIDE_BOX_LONG,
OUTSIDE_BOX_WIDTH,
OUTSIDE_BOX_HEIGHT,
OUTSIDE_BOX_WEIGHT,
INSTRUCTION_SIZE,
INSTRUCTION_REQUIRE,
MATTERS_NEEDING_ATTENTION,
IF_PAPER_PRODUCTION_AUDIT,
IF_WOOD_IRON_PRODUCTION_AUDIT,
IF_ACRYLIC_PRODUCTION_AUDIT,
IF_PAPER_PLAN_AUDIT,
IF_WOOD_IRON_PLAN_AUDIT,
IF_STRUCTURE_AUDIT,
IF_PLANE_AUDIT,
IF_PAPER_PURCHASE_AUDIT,
IF_WOOD_IRON_PURCHASE_AUDIT,
APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID,
APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID,
APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID,
APPOINT_PAPER_PLAN_AUDIT_MAKERID,
APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID,
APPOINT_STRUCTURE_AUDIT_MAKERID,
APPOINT_PLANE_AUDIT_MAKERID,
APPOINT_PAPER_PURCHASE_AUDIT_MAKERID,
APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID,
PAPER_PRODUCTION_AUDIT_STATUS,
WOOD_IRON_PRODUCTION_AUDIT_STATUS,
ACRYLIC_PRODUCTION_AUDIT_STATUS,
PAPER_PLAN_AUDIT_STATUS,
WOOD_IRON_PLAN_AUDIT_STATUS,
STRUCTURE_AUDIT_STATUS,
PLANE_AUDIT_STATUS,
PAPER_PURCHASE_AUDIT_STATUS,
WOOD_IRON_PURCHASE_AUDIT_STATUS,
MakerID,
Date

)
VALUES
(
@PNID,
@PFID,
@PIID,
@ORDER_DATE,
@EDIT_DATE,
@EDIT_TIMES,
@ORDER_ID,
@IF_AUDIT_PRICE,
@WAREID,
@PRODUCTION_COUNT,
@HAVE_TAX_UNIT_PRICE,
@DELIVERY_DATE,
@DELIVERY_BATCH,
@DELIVERY_PLACE,
@ORDER_TYPE,
@PACKING_METHOD,
@OUTSIDE_BOX_MATERIAL,
@OUTSIDE_BOX_LONG,
@OUTSIDE_BOX_WIDTH,
@OUTSIDE_BOX_HEIGHT,
@OUTSIDE_BOX_WEIGHT,
@INSTRUCTION_SIZE,
@INSTRUCTION_REQUIRE,
@MATTERS_NEEDING_ATTENTION,
@IF_PAPER_PRODUCTION_AUDIT,
@IF_WOOD_IRON_PRODUCTION_AUDIT,
@IF_ACRYLIC_PRODUCTION_AUDIT,
@IF_PAPER_PLAN_AUDIT,
@IF_WOOD_IRON_PLAN_AUDIT,
@IF_STRUCTURE_AUDIT,
@IF_PLANE_AUDIT,
@IF_PAPER_PURCHASE_AUDIT,
@IF_WOOD_IRON_PURCHASE_AUDIT,
@APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID,
@APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID,
@APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID,
@APPOINT_PAPER_PLAN_AUDIT_MAKERID,
@APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID,
@APPOINT_STRUCTURE_AUDIT_MAKERID,
@APPOINT_PLANE_AUDIT_MAKERID,
@APPOINT_PAPER_PURCHASE_AUDIT_MAKERID,
@APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID,
@PAPER_PRODUCTION_AUDIT_STATUS,
@WOOD_IRON_PRODUCTION_AUDIT_STATUS,
@ACRYLIC_PRODUCTION_AUDIT_STATUS,
@PAPER_PLAN_AUDIT_STATUS,
@WOOD_IRON_PLAN_AUDIT_STATUS,
@STRUCTURE_AUDIT_STATUS,
@PLANE_AUDIT_STATUS,
@PAPER_PURCHASE_AUDIT_STATUS,
@WOOD_IRON_PURCHASE_AUDIT_STATUS,
@MakerID,
@Date

)
";
        string setsqlth = @"


";

        string setsqlf = @"
UPDATE PN_PRODUCTION_INSTRUCTIONS SET
PNID=@PNID,
PFID=@PFID,
PIID=@PIID,
ORDER_ID=@ORDER_ID,
ORDER_TYPE=@ORDER_TYPE,
EDIT_DATE=@EDIT_DATE,
EDIT_TIMES=@EDIT_TIMES,
IF_AUDIT_PRICE=@IF_AUDIT_PRICE,
WAREID=@WAREID,
PRODUCTION_COUNT=@PRODUCTION_COUNT,
HAVE_TAX_UNIT_PRICE=@HAVE_TAX_UNIT_PRICE,
DELIVERY_BATCH=@DELIVERY_BATCH,
DELIVERY_PLACE=@DELIVERY_PLACE,
PACKING_METHOD=@PACKING_METHOD,
OUTSIDE_BOX_MATERIAL=@OUTSIDE_BOX_MATERIAL,
OUTSIDE_BOX_LONG=@OUTSIDE_BOX_LONG,
OUTSIDE_BOX_WIDTH=@OUTSIDE_BOX_WIDTH,
OUTSIDE_BOX_HEIGHT=@OUTSIDE_BOX_HEIGHT,
OUTSIDE_BOX_WEIGHT=@OUTSIDE_BOX_WEIGHT,
INSTRUCTION_SIZE=@INSTRUCTION_SIZE,
INSTRUCTION_REQUIRE=@INSTRUCTION_REQUIRE,
MATTERS_NEEDING_ATTENTION=@MATTERS_NEEDING_ATTENTION,
IF_PAPER_PRODUCTION_AUDIT=@IF_PAPER_PRODUCTION_AUDIT,
IF_WOOD_IRON_PRODUCTION_AUDIT=@IF_WOOD_IRON_PRODUCTION_AUDIT,
IF_ACRYLIC_PRODUCTION_AUDIT=@IF_ACRYLIC_PRODUCTION_AUDIT,
IF_PAPER_PLAN_AUDIT=@IF_PAPER_PLAN_AUDIT,
IF_WOOD_IRON_PLAN_AUDIT=@IF_WOOD_IRON_PLAN_AUDIT,
IF_STRUCTURE_AUDIT=@IF_STRUCTURE_AUDIT,
IF_PLANE_AUDIT=@IF_PLANE_AUDIT,
IF_PAPER_PURCHASE_AUDIT=@IF_PAPER_PURCHASE_AUDIT,
IF_WOOD_IRON_PURCHASE_AUDIT=@IF_WOOD_IRON_PURCHASE_AUDIT,
APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID=@APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID,
APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID=@APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID,
APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID=@APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID,
APPOINT_PAPER_PLAN_AUDIT_MAKERID=@APPOINT_PAPER_PLAN_AUDIT_MAKERID,
APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID=@APPOINT_PAPER_PLAN_AUDIT_MAKERID,
APPOINT_STRUCTURE_AUDIT_MAKERID=@APPOINT_STRUCTURE_AUDIT_MAKERID,
APPOINT_PLANE_AUDIT_MAKERID=@APPOINT_PLANE_AUDIT_MAKERID,
APPOINT_PAPER_PURCHASE_AUDIT_MAKERID=@APPOINT_PAPER_PURCHASE_AUDIT_MAKERID,
APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID=@APPOINT_PAPER_PURCHASE_AUDIT_MAKERID,
Date=@Date
";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        public CPN_PRODUCTION_INSTRUCTIONS()
        {
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
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("订单编号", typeof(string));
            dt.Columns.Add("品号", typeof(string));
            dt.Columns.Add("客户名称", typeof(string));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("结构设计", typeof(string));
            dt.Columns.Add("平面设计", typeof(string));
            dt.Columns.Add("生产数量", typeof(string));
            dt.Columns.Add("下单日期", typeof(string));
            dt.Columns.Add("交货日期", typeof(string));
            dt.Columns.Add("报价", typeof(string));
            dt.Columns.Add("会签", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            return dt;
        }

        #endregion
        #region GetTableInfo_SEARCH
        public DataTable GetTableInfo_SEARCH_o()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("订单编号", typeof(string));
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
            string v1 = bc.numYM(10, 4, "0001", "select * from PN_PRODUCTION_INSTRUCTIONS_NO", "PNID", "PN");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO PN_PRODUCTION_INSTRUCTIONS_NO(PNID,DATE,YEAR,MONTH) VALUES ('" + v1 + "','" + varDate + "','" + year +
                  "','" + month + "')");
              
            }
            return GETID;
        }
        #endregion
        #region GETID_ORDER_ID
        public string GETID_ORDER_ID(string YYMM,string ORDER_TYPE)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = "";
            if (ORDER_TYPE.Length == 2)
            {
                  v1 = numNOYMD(3, 3, "001", "select * from PN_PRODUCTION_INSTRUCTIONS WHERE SUBSTRING(ORDER_ID,1,4)='" + YYMM +
                      "' ", "ORDER_ID", YYMM + ORDER_TYPE, " ORDER BY RIGHT(ORDER_ID,3) ASC");
            }
            if (ORDER_TYPE.Length == 3)
            {
                v1 = numNOYMD(3, 3, "001", "select * from PN_PRODUCTION_INSTRUCTIONS WHERE SUBSTRING(ORDER_ID,1,4)='" + YYMM +
                    "' ", "ORDER_ID", YYMM + ORDER_TYPE, " ORDER BY RIGHT(ORDER_ID,3) ASC");
            }
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #endregion
        #region 编号NOYMD
        public string numNOYMD(int digit, int wcodedigit, string wcode, string sql, string tbColumns, string prifix, string sort)
        {

            string year, month, day;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            day = DateTime.Now.ToString("dd");
            string P_str_Code, t, r, q = "";
            int P_int_Code, w, w1;
            DataTable dt = bc.getdt(sql + sort);
            if (dt.Rows.Count > 0)
            {
                P_str_Code = Convert.ToString(dt.Rows[(dt.Rows.Count - 1)][tbColumns]);
                w1 = P_str_Code.Length - 3;
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
                    r = prifix + q + P_int_Code;
                }
                else
                {
                    r = "Exceed Limited";

                }

            }
            else
            {
                r = prifix + wcode;
            }
            return r;
        }
        #endregion
        #region save
        public void save(string YYMM,string ORDER_TYPE)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT IF_SUBMIT FROM PN_PRODUCTION_INSTRUCTIONS WHERE PNID='" + PNID + "'");
            string vorder_id = bc.getOnlyString("SELECT ORDER_ID FROM PN_PRODUCTION_INSTRUCTIONS WHERE PNID='" + PNID + "'");
            string vorder_type = "";
          
      
            if (!bc.exists("SELECT PNID FROM PN_PRODUCTION_INSTRUCTIONS WHERE PNID='" + PNID + "'"))
            {
                ORDER_ID =GETID_ORDER_ID(YYMM,ORDER_TYPE);
                SQlcommandE(sqlt);
                IFExecution_SUCCESS = true;
            }
            else
            {
                ORDER_ID = vorder_id;//未修改订单类型，订单编号不变
                if (vorder_id.Length > 0)
                {
                    if (vorder_id.Length == 9)
                    {
                        vorder_type = vorder_id.Substring(4, 2);
                    }
                    else if (vorder_id.Length == 10)
                    {
                        vorder_type = vorder_id.Substring(4, 3);
                    }
                    if (vorder_type != ORDER_TYPE)//如果修改了订单类型，那么要把订单编号里的订单类型替换为新的订单类型
                    {
                        ORDER_ID = ORDER_ID.Substring(0, 4) + ORDER_TYPE  + ORDER_ID.Substring(ORDER_ID.Length - 3, 3);//此为新的订单编号
                      
                    }
                }
                SQlcommandE(sqlf + " WHERE PNID='" + PNID + "'");
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
            sqlcom.Parameters.Add("@PNID", SqlDbType.VarChar, 20).Value = PNID;
            sqlcom.Parameters.Add("@PFID", SqlDbType.VarChar, 20).Value = PFID;
            sqlcom.Parameters.Add("@PIID", SqlDbType.VarChar, 20).Value = PIID;
            sqlcom.Parameters.Add("@ORDER_DATE", SqlDbType.VarChar, 20).Value = ORDER_DATE;
            sqlcom.Parameters.Add("@EDIT_DATE", SqlDbType.VarChar, 20).Value = EDIT_DATE;
            sqlcom.Parameters.Add("@EDIT_TIMES", SqlDbType.VarChar, 20).Value = EDIT_TIMES;
            sqlcom.Parameters.Add("@ORDER_ID", SqlDbType.VarChar, 20).Value = ORDER_ID;
            sqlcom.Parameters.Add("@IF_AUDIT_PRICE", SqlDbType.VarChar, 20).Value = IF_AUDIT_PRICE;
            sqlcom.Parameters.Add("@WAREID", SqlDbType.VarChar, 20).Value = WAREID;
            sqlcom.Parameters.Add("@PRODUCTION_COUNT", SqlDbType.VarChar, 20).Value = PRODUCTION_COUNT;
            sqlcom.Parameters.Add("@HAVE_TAX_UNIT_PRICE", SqlDbType.VarChar, 20).Value = HAVE_TAX_UNIT_PRICE;
            sqlcom.Parameters.Add("@DELIVERY_DATE", SqlDbType.VarChar, 20).Value = DELIVERY_DATE;
            sqlcom.Parameters.Add("@DELIVERY_BATCH", SqlDbType.VarChar, 20).Value = DELIVERY_BATCH;
            sqlcom.Parameters.Add("@DELIVERY_PLACE", SqlDbType.VarChar, 20).Value = DELIVERY_PLACE;
            sqlcom.Parameters.Add("@ORDER_TYPE", SqlDbType.VarChar, 20).Value = ORDER_TYPE;
            sqlcom.Parameters.Add("@PACKING_METHOD", SqlDbType.VarChar, 20).Value = PACKING_METHOD;
            sqlcom.Parameters.Add("@OUTSIDE_BOX_MATERIAL", SqlDbType.VarChar, 20).Value = OUTSIDE_BOX_MATERIAL;
            sqlcom.Parameters.Add("@OUTSIDE_BOX_LONG", SqlDbType.VarChar, 20).Value = OUTSIDE_BOX_LONG;
            sqlcom.Parameters.Add("@OUTSIDE_BOX_WIDTH", SqlDbType.VarChar, 20).Value = OUTSIDE_BOX_WIDTH;
            sqlcom.Parameters.Add("@OUTSIDE_BOX_HEIGHT", SqlDbType.VarChar, 20).Value = OUTSIDE_BOX_HEIGHT;
            sqlcom.Parameters.Add("@OUTSIDE_BOX_WEIGHT", SqlDbType.VarChar, 20).Value = OUTSIDE_BOX_WEIGHT;
            sqlcom.Parameters.Add("@INSTRUCTION_SIZE", SqlDbType.VarChar, 20).Value = INSTRUCTION_SIZE;
            sqlcom.Parameters.Add("@INSTRUCTION_REQUIRE", SqlDbType.VarChar, 20).Value = INSTRUCTION_REQUIRE;
            sqlcom.Parameters.Add("@MATTERS_NEEDING_ATTENTION", SqlDbType.VarChar, 1000).Value = MATTERS_NEEDING_ATTENTION;
            sqlcom.Parameters.Add("@IF_PAPER_PRODUCTION_AUDIT", SqlDbType.VarChar, 20).Value = IF_PAPER_PRODUCTION_AUDIT;
            sqlcom.Parameters.Add("@IF_WOOD_IRON_PRODUCTION_AUDIT", SqlDbType.VarChar, 20).Value = IF_WOOD_IRON_PRODUCTION_AUDIT;
            sqlcom.Parameters.Add("@IF_ACRYLIC_PRODUCTION_AUDIT", SqlDbType.VarChar, 20).Value = IF_ACRYLIC_PRODUCTION_AUDIT;
            sqlcom.Parameters.Add("@IF_PAPER_PLAN_AUDIT", SqlDbType.VarChar, 20).Value = IF_PAPER_PLAN_AUDIT;
            sqlcom.Parameters.Add("@IF_WOOD_IRON_PLAN_AUDIT", SqlDbType.VarChar, 20).Value = IF_WOOD_IRON_PLAN_AUDIT;
            sqlcom.Parameters.Add("@IF_STRUCTURE_AUDIT", SqlDbType.VarChar, 20).Value = IF_STRUCTURE_AUDIT;
            sqlcom.Parameters.Add("@IF_PLANE_AUDIT", SqlDbType.VarChar, 20).Value = IF_PLANE_AUDIT;
            sqlcom.Parameters.Add("@IF_PAPER_PURCHASE_AUDIT", SqlDbType.VarChar, 20).Value = IF_PAPER_PURCHASE_AUDIT;
            sqlcom.Parameters.Add("@IF_WOOD_IRON_PURCHASE_AUDIT", SqlDbType.VarChar, 20).Value = IF_WOOD_IRON_PURCHASE_AUDIT;
            sqlcom.Parameters.Add("@APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value=APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value=APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value=APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_PAPER_PLAN_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value=APPOINT_PAPER_PLAN_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value=APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_STRUCTURE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = APPOINT_STRUCTURE_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_PLANE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = APPOINT_PLANE_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_PAPER_PURCHASE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = APPOINT_PAPER_PURCHASE_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID;
            sqlcom.Parameters.Add("@PAPER_PRODUCTION_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = PAPER_PRODUCTION_AUDIT_STATUS;
            sqlcom.Parameters.Add("@WOOD_IRON_PRODUCTION_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = WOOD_IRON_PRODUCTION_AUDIT_STATUS;
            sqlcom.Parameters.Add("@ACRYLIC_PRODUCTION_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = ACRYLIC_PRODUCTION_AUDIT_STATUS;
            sqlcom.Parameters.Add("@PAPER_PLAN_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = PAPER_PLAN_AUDIT_STATUS;
            sqlcom.Parameters.Add("@WOOD_IRON_PLAN_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = WOOD_IRON_PLAN_AUDIT_STATUS;
            sqlcom.Parameters.Add("@STRUCTURE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = "N";
            sqlcom.Parameters.Add("@PLANE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = "N";
            sqlcom.Parameters.Add("@PAPER_PURCHASE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = "N";
            sqlcom.Parameters.Add("@WOOD_IRON_PURCHASE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = "N";
            sqlcom.Parameters.Add("@MakerID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region  RETURN_DT_SEARCH
        public DataTable RETURN_DT_SEARCH(DataTable dtt, string EMID,string POSITION)
        {
            DataTable dt = GetTableInfo_SEARCH();
            i = 1;
            if (dtt.Rows.Count > 0)
            {
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
                        dr["订单编号"] = dr1["订单编号"].ToString();
                        dr["品号"] = dr1["品号"].ToString();
                        dr["客户名称"] = dr1["客户名称"].ToString();
                        dr["AE"] = dr1["AE01"].ToString();
                        dr["结构设计"] = dr1["结构01"].ToString();
                        dr["平面设计"] = dr1["平面01"].ToString();
                        dr["生产数量"] = dr1["生产数量"].ToString();
                        dr["下单日期"] = dr1["下单日期"].ToString();
                        dr["交货日期"] = dr1["交货日期"].ToString();
                        dr["报价"] = dr1["报价"].ToString();
                        dr["制单人"] = dr1["制单人"].ToString();
                        if (JUAGE_IF_AUDIT_END(dr1["编号"].ToString()))
                        {
                            dr["会签"] = "已会签";
                        }
                        else
                        {
                            dr["会签"] = "待会签";
                        }
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
            worksheet.Cells[4, "C"] = dt.Rows[0]["下单日期"].ToString();
            worksheet.Cells[4, "F"] = dt.Rows[0]["修改日期"].ToString();
            worksheet.Cells[4, "I"] = dt.Rows[0]["修改次数"].ToString();
            worksheet.Cells[6, "I"] = dt.Rows[0]["报价"].ToString();
            worksheet.Cells[7, "C"] = dt.Rows[0]["订单编号"].ToString();
            worksheet.Cells[7, "F"] = dt.Rows[0]["项目号"].ToString();
            worksheet.Cells[7, "I"] = dt.Rows[0]["报价编号"].ToString();
            worksheet.Cells[8, "C"] = dt.Rows[0]["品号"].ToString();
            worksheet.Cells[8, "I"] = dt.Rows[0]["生产数量"].ToString();
            worksheet.Cells[9, "C"] = dt.Rows[0]["客户名称"].ToString();
            worksheet.Cells[9, "I"] = dt.Rows[0]["品牌"].ToString();
            worksheet.Cells[10, "C"] = dt.Rows[0]["交货日期"].ToString();
            worksheet.Cells[10, "F"] = dt.Rows[0]["交货批次"].ToString();
            worksheet.Cells[10, "I"] = dt.Rows[0]["交货地点"].ToString();
            worksheet.Cells[11, "C"] = dt.Rows[0]["AE01"].ToString();
            worksheet.Cells[11, "F"] = dt.Rows[0]["平面01"].ToString();
            worksheet.Cells[11, "I"] = dt.Rows[0]["结构01"].ToString();
            worksheet.Cells[12, "C"] = dt.Rows[0]["AE02"].ToString();
            worksheet.Cells[12, "F"] = dt.Rows[0]["平面02"].ToString();
            worksheet.Cells[12, "I"] = dt.Rows[0]["结构02"].ToString();
            worksheet.Cells[13, "C"] = dt.Rows[0]["AE03"].ToString();
            worksheet.Cells[13, "F"] = dt.Rows[0]["平面03"].ToString();
            worksheet.Cells[13, "I"] = dt.Rows[0]["结构03"].ToString();
            worksheet.Cells[16, "C"] = dt.Rows[0]["包装方式"].ToString();
            worksheet.Cells[16, "G"] = dt.Rows[0]["外箱材质"].ToString();
            worksheet.Cells[17, "D"] = dt.Rows[0]["长"].ToString();
            worksheet.Cells[17, "F"] = dt.Rows[0]["宽"].ToString();
            worksheet.Cells[17, "H"] = dt.Rows[0]["高"].ToString();
            worksheet.Cells[18, "C"] = dt.Rows[0]["外箱重量"].ToString();
            worksheet.Cells[18, "G"] = dt.Rows[0]["说明书要求"].ToString();
            worksheet.Cells[18, "I"] = dt.Rows[0]["说明书尺寸"].ToString();
            worksheet.Cells[21, "B"] = dt.Rows[0]["生产注意事项"].ToString();
            worksheet.Cells[39, "B"] = dt.Rows[0]["纸品生产"].ToString();
            worksheet.Cells[39, "C"] = dt.Rows[0]["木铁生产"].ToString();
            worksheet.Cells[39, "D"] = dt.Rows[0]["亚克力生产"].ToString();
            worksheet.Cells[39, "E"] = dt.Rows[0]["纸品计划"].ToString();
            worksheet.Cells[39, "F"] = dt.Rows[0]["木铁计划"].ToString();
            worksheet.Cells[39, "G"] = dt.Rows[0]["结构设计"].ToString();
            worksheet.Cells[39, "H"] = dt.Rows[0]["平面设计"].ToString();
            worksheet.Cells[39, "I"] = dt.Rows[0]["纸品采购"].ToString();
            worksheet.Cells[39, "J"] = dt.Rows[0]["木铁采购"].ToString();
            DataTable  dtx = bc.getdt("SELECT * FROM WAREFILE WHERE WAREID='"+dt.Rows [0]["编号"].ToString ()+"'");
            if (dtx.Rows.Count > 0)
            {
              
               /*将服务器上的图片下载到本地 start*/
               WebClient wclient = new WebClient();
               wclient.DownloadFile(dtx.Rows[0]["PATH"].ToString(), "d:\\" + dtx.Rows[0]["NEW_FILE_NAME"].ToString());
               /*将服务器上的图片下载到本地 end*/
               worksheet.Shapes.AddPicture("d:\\" + dtx.Rows[0]["NEW_FILE_NAME"].ToString(), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 360, 345, 150, 150);
               //worksheet.Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1, "123456", "Red", 15, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 150, 200);
               /*删除本地的临时图片文件 start*/
               if (File.Exists("d:\\" + dtx.Rows[0]["NEW_FILE_NAME"].ToString()))
               {
                   File.Delete("d:\\" + dtx.Rows[0]["NEW_FILE_NAME"].ToString());
               }
                /*删除本地的临时图片文件 end*/
            }
      
             //下面为第二张图片
            dtx = bc.getdt("SELECT * FROM WAREFILE WHERE WAREID='"+dt.Rows [0]["编号"].ToString ()+"'");
            if (dtx.Rows.Count > 0 && dtx.Rows.Count >=4)
            {
              
               /*将服务器上的图片下载到本地 start*/
               WebClient wclient = new WebClient();
               wclient.DownloadFile(dtx.Rows[2]["PATH"].ToString(), "d:\\" + dtx.Rows[2]["NEW_FILE_NAME"].ToString());
               /*将服务器上的图片下载到本地 end*/
               worksheet.Shapes.AddPicture("d:\\" + dtx.Rows[2]["NEW_FILE_NAME"].ToString(), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 360, 500, 150, 150);
               //worksheet.Shapes.AddTextEffect(Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1, "123456", "Red", 15, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 150, 200);
               /*删除本地的临时图片文件 start*/
               if (File.Exists("d:\\" + dtx.Rows[2]["NEW_FILE_NAME"].ToString()))
               {
                   File.Delete("d:\\" + dtx.Rows[2]["NEW_FILE_NAME"].ToString());
               }
                /*删除本地的临时图片文件 end*/
            }
      
           dt = bc.getdt(cnotice_list .sql );
           if (dt.Rows.Count > 0)
           {
               foreach (DataRow dr in dt.Rows)
               {
                 
                   if (dr["员工姓名"].ToString() == "黄范燮")
                   {
                       worksheet.Cells[42, "B"] = dr["员工姓名"].ToString();
                   }
                   if (dr["员工姓名"].ToString() == "孙加平")
                   {
                       worksheet.Cells[42, "C"] = dr["员工姓名"].ToString();
                   }
                   if (dr["员工姓名"].ToString() == "张峰")
                   {
                       worksheet.Cells[42, "D"] = dr["员工姓名"].ToString();
                   }
                   if (dr["员工姓名"].ToString() == "vicky")
                   {
                       worksheet.Cells[42, "E"] = dr["员工姓名"].ToString();
                   }
                   if (dr["员工姓名"].ToString() == "马小龙")
                   {
                       worksheet.Cells[42, "F"] = dr["员工姓名"].ToString();
                   }
               }
           }
        
        }
        #endregion
        #region JUAGE_IF_AUDIT_END 
        public bool JUAGE_IF_AUDIT_END(string PNID)
        {
            bool b = true;
            DataTable dt = new DataTable();
            dt = bc.getdt(sql + string.Format(" WHERE A.PNID='{0}'", PNID));
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["是否需纸品生产签核"].ToString() == "是" && dt.Rows[0]["纸品生产签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需木铁生产签核"].ToString() == "是" && dt.Rows[0]["木铁生产签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需亚克力生产签核"].ToString() == "是" && dt.Rows[0]["亚克力生产签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需纸品计划签核"].ToString() == "是" && dt.Rows[0]["纸品计划签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需木铁计划签核"].ToString() == "是" && dt.Rows[0]["木铁计划签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需平面设计签核"].ToString() == "是" && dt.Rows[0]["平面设计签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需结构设计签核"].ToString() == "是" && dt.Rows[0]["结构设计签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需纸品采购签核"].ToString() == "是" && dt.Rows[0]["纸品采购签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需木铁采购签核"].ToString() == "是" && dt.Rows[0]["木铁采购签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
            }
            else
            {
                b = false;
            }
            return b;

        }
        #endregion
        #region JUAGE_IF_EXISTS_AUDIT_RECORD 
        public bool JUAGE_IF_EXISTS_AUDIT_RECORD(string PNID)//此方法判断实际审核状态
        {
            bool b = false;
            DataTable dt = new DataTable();
            dt = bc.getdt(sql + string.Format(" WHERE A.PNID='{0}'", PNID));
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["是否需纸品生产签核"].ToString() == "是" && dt.Rows[0]["纸品生产签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁生产签核"].ToString() == "是" && dt.Rows[0]["木铁生产签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需亚克力生产签核"].ToString() == "是" && dt.Rows[0]["亚克力生产签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需纸品计划签核"].ToString() == "是" && dt.Rows[0]["纸品计划签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁计划签核"].ToString() == "是" && dt.Rows[0]["木铁计划签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需平面设计签核"].ToString() == "是" && dt.Rows[0]["平面设计签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需结构设计签核"].ToString() == "是" && dt.Rows[0]["结构设计签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需纸品采购签核"].ToString() == "是" && dt.Rows[0]["纸品采购签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁采购签核"].ToString() == "是" && dt.Rows[0]["木铁采购签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
            }
        
            return b;

        }
        #endregion
        #region JUAGE_IF_EXISTS_AUDIT
        public bool JUAGE_IF_EXISTS_AUDIT(string PNID)//此方法判断原始审核数据，含重新送签
        {
            bool b = false;
            DataTable dt = new DataTable();
            dt = bc.getdt(sql + string.Format(" WHERE A.PNID='{0}'", PNID));
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["是否需纸品生产签核"].ToString() == "是" && dt.Rows[0]["纸品生产签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁生产签核"].ToString() == "是" && dt.Rows[0]["木铁生产签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需亚克力生产签核"].ToString() == "是" && dt.Rows[0]["亚克力生产签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需纸品计划签核"].ToString() == "是" && dt.Rows[0]["纸品计划签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁计划签核"].ToString() == "是" && dt.Rows[0]["木铁计划签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需平面设计签核"].ToString() == "是" && dt.Rows[0]["平面设计签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需结构设计签核"].ToString() == "是" && dt.Rows[0]["结构设计签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需纸品采购签核"].ToString() == "是" && dt.Rows[0]["纸品采购签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
                else if (dt.Rows[0]["是否需木铁采购签核"].ToString() == "是" && dt.Rows[0]["木铁采购签核状态"].ToString() == "已签核")
                {
                    b = true;
                }
            }

            return b;

        }
        #endregion
        #region RETURN_OFFER_ID_DT
        public DataTable RETURN_OFFER_ID_DT(string OFFER_ID)
        {
            
            DataTable dtx = new DataTable();
            dtx = bc.getdt(string.Format(cprinting_offer.sql + " WHERE B.OFFER_ID='{0}' ORDER BY B.OFFER_ID ASC", OFFER_ID ));
            DataTable dtx1 = new DataTable();
            dtx1.Columns.Add("报价编号", typeof(string));
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    DataRow dr1 = dtx1.NewRow();
                    dr1["报价编号"] = dr["报价编号"].ToString();
                    dtx1.Rows.Add(dr1);
                }
            }
            dtx = bc.getdt(string.Format(cno_paper_offer.sql + " WHERE D.OFFER_ID='{0}' ORDER BY D.OFFER_ID ASC", OFFER_ID ));
            if (dtx.Rows.Count > 0)
            {

                foreach (DataRow dr in dtx.Rows)
                {
                    DataRow dr1 = dtx1.NewRow();
                    dr1["报价编号"] = dr["报价编号"].ToString();
                    dtx1.Rows.Add(dr1);
                }
            }
            return dtx1;
        }
        #endregion
        #region RETURN_OFFER_ID_IF_EXISTS
        public bool RETURN_OFFER_ID_IF_EXISTS(string OFFER_ID)
        {
            bool b = false;
            DataTable dt = RETURN_OFFER_ID_DT(OFFER_ID);
            if (dt.Rows.Count > 0)
            {
                b = true;
            }
            return b;
        }
        #endregion
    }
}
