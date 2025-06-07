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

namespace XizheC
{
    public class CAUDIT_LIST
    {
        basec bc = new basec();
        #region nature
        private string _USID;
        public string USID
        {
            set { _USID = value; }
            get { return _USID; }

        }
        private string _sql;
        public string sql
        {
            set { _sql = value; }
            get { return _sql; }

        }
        private string _UNAME;
        public string UNAME
        {
            set { _UNAME = value; }
            get { return _UNAME; }

        }
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _ENAME;
        public string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        private string _AUDIT_LIST;
        public string AUDIT_LIST
        {
            set { _AUDIT_LIST = value; }
            get { return _AUDIT_LIST; }

        }
        #endregion
        DataTable dt = new DataTable();
        string setsql = @"

SELECT 
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
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.WOOD_IRON_PURCHASE_AUDIT_MAKERID)  AS 木铁采购 FROM AUDIT_LIST A

";
        public CAUDIT_LIST()
        {
            sql = setsql;
        }
 
    
    }
}
