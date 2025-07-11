﻿using System;
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
    public class CPRINT_PACK_MATERIAL:IGETID 
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private string _PPKEY;
        public string PPKEY
        {
            set { _PPKEY = value; }
            get { return _PPKEY; }

        }
        private string _PROJECT_ID;
        public string PROJECT_ID
        {
            set { _PROJECT_ID = value; }
            get { return _PROJECT_ID; }
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

        private string _PROJECT;
        public string PROJECT
        {
            set { _PROJECT = value; }
            get { return _PROJECT; }

        }
        private string _DPID;
        public string DPID
        {
            set { _DPID = value; }
            get { return _DPID; }

        }
        private string _STARTING_PRICE;
        public string STARTING_PRICE
        {
            set { _STARTING_PRICE = value; }
            get { return _STARTING_PRICE; }

        }
        private string _STARTING_PRICE_UNIT;
        public string STARTING_PRICE_UNIT
        {
            set { _STARTING_PRICE_UNIT = value; }
            get { return _STARTING_PRICE_UNIT; }

        }
        private string _UNIT_PRICE_UNIT;
        public string UNIT_PRICE_UNIT
        {
            set { _UNIT_PRICE_UNIT = value; }
            get { return _UNIT_PRICE_UNIT; }

        }
        private string _MAX_PRICE;
        public string MAX_PRICE
        {
            set { _MAX_PRICE = value; }
            get { return _MAX_PRICE; }

        }
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }

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
A.PFID AS 编号,
A.PROJECT AS 项目,
A.PACK_MATERIAL_LENGTH AS 长,
A.PACK_MATERIAL_WIDTH AS 宽,
A.PACK_MATERIAL_HEIGHT AS 高,
A.COUNT AS 数量,
A.PRICE AS 单价,
A.BOX_TYPE AS 箱形,
A.MATERIAL AS 材质,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM PRINT_PACK_MATERIAL A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID
LEFT JOIN PRINTING_OFFER_MST C ON A.PFID=C.PFID


";


        string setsqlo = @"



";

        string setsqlt = @"

INSERT INTO PRINT_PACK_MATERIAL
(
PPKEY,
SN,
PFID,
PROJECT,
PACK_MATERIAL_LENGTH,
PACK_MATERIAL_WIDTH,
PACK_MATERIAL_HEIGHT,
BOX_TYPE,
MATERIAL,
COUNT,
PRICE,
MakerID,
Date,
Year,
Month,
day
)
VALUES
(
@PPKEY,
@SN,
@PFID,
@PROJECT,
@PACK_MATERIAL_LENGTH,
@PACK_MATERIAL_WIDTH,
@PACK_MATERIAL_HEIGHT,
@BOX_TYPE,
@MATERIAL,
@COUNT,
@PRICE,
@MakerID,
@Date,
@Year,
@Month,
@day
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
        CPACK_MATERIAL cPACK_MATERIAL = new CPACK_MATERIAL();
        
        public CPRINT_PACK_MATERIAL()
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
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM PRINT_PACK_MATERIAL", "PPID", "PP");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region save
        public void save(DataGridView dgv)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            basec.getcoms("DELETE PRINT_PACK_MATERIAL WHERE PFID='" + PFID + "'");
            SQlcommandE(sqlt, dgv);
            IFExecution_SUCCESS = true;
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql,DataGridView dgv)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcon.Open();
                PPKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM PRINT_PACK_MATERIAL", "PPKEY", "PP");
                sqlcom.Parameters.Add("PPKEY", SqlDbType.VarChar, 20).Value = PPKEY;
                sqlcom.Parameters.Add("SN", SqlDbType.VarChar, 20).Value = dgv["序号", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                sqlcom.Parameters.Add("PROJECT", SqlDbType.VarChar, 20).Value = dgv["项目", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("PACK_MATERIAL_LENGTH", SqlDbType.VarChar, 20).Value = dgv["长", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("PACK_MATERIAL_WIDTH", SqlDbType.VarChar, 20).Value = dgv["宽", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("PACK_MATERIAL_HEIGHT", SqlDbType.VarChar, 20).Value = dgv["高", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("BOX_TYPE", SqlDbType.VarChar, 20).Value = dgv["箱形", i].FormattedValue.ToString();
                sqlcom.Parameters.Add("MATERIAL", SqlDbType.VarChar, 20).Value = dgv["材质", i].FormattedValue.ToString();
                if (!string.IsNullOrEmpty(dgv["数量", i].FormattedValue.ToString()))
                {
                    sqlcom.Parameters.Add("COUNT", SqlDbType.VarChar, 20).Value = dgv["数量", i].FormattedValue.ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("COUNT", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                DataTable dtx1 = bc.getdt(@"
SELECT * FROM PACK_MATERIAL A WHERE A.PACK_MATERIAL='" + dgv["材质", i].FormattedValue.ToString() + "'AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='" +
bc.RETURN_CUSTOMER_TYPE(PROJECT_ID) + "' ");
             
                decimal d1 = 0, d2 = 0, d3 = 0;
                if (dtx1.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))
                    {
                        d1 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                    {

                        d2 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                    }

                    d3 = d1 / (1 + d2 / 100);
                    //此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10
                }
               // MessageBox.Show(d3.ToString ());
                if (i == 0 || i == 1 || i == 2 || i == 3 || i == 4 || i == 5 || i == 6)
                {
                    if (d3 != 0)
                    {
                        sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = d3;//前7行的单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10

                    }
                    else
                    {
                        sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;

                    }
                }
                else if (!string.IsNullOrEmpty(dgv["单价", i].FormattedValue.ToString()))
                {
                    sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = dgv["单价", i].FormattedValue.ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("PRICE", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
                sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
            }
          
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("项目", typeof(string));
            dt.Columns.Add("长", typeof(string));
            dt.Columns.Add("宽", typeof(string));
            dt.Columns.Add("高", typeof(string));
            dt.Columns.Add("箱形", typeof(string));
            dt.Columns.Add("材质", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("单价", typeof(string));
            return dt;
        }
        #endregion
        #region RETURN_NO_FREE_KEY_DT
        public DataTable RETURN_NO_FREE_KEY_DT(DataTable dt)
        {
            DataTable dtt = GetTableInfo();
            DataTable dtx = bc.getdt(cPACK_MATERIAL.sql);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtx, "", string.Format("项目='{0}'", dr1["项目"].ToString()));
                    if (dtt1.Rows.Count > 0)
                    {
                        DataRow dr = dtt.NewRow();
                        dr["项目"] = dr1["项目"].ToString();
                        dr["长"] = dr1["长"].ToString();
                        dr["宽"] = dr1["宽"].ToString();
                        dr["高"] = dr1["高"].ToString();
                        dr["箱形"] = dr1["箱形"].ToString();
                        dr["材质"] = dr1["材质"].ToString();
                        dr["数量"] = dr1["数量"].ToString();
                        dr["单价"] = dr1["单价"].ToString();
                        dtt.Rows.Add(dr);
                    }
                    else
                    {
                      
                    }
                }
            }

            return dtt;
        }
        #endregion
        #region RETURN_FREE_KEY_DT
        public DataTable RETURN_FREE_KEY_DT(DataTable dt)
        {
            DataTable dtt = GetTableInfo();
            DataTable dtx = bc.getdt(cPACK_MATERIAL.sql);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtx, "", string.Format("项目='{0}'", dr1["项目"].ToString()));
                    if (dtt1.Rows.Count > 0)
                    {
                    
                    }
                    else
                    {
                        DataRow dr = dtt.NewRow();
                        dr["项目"] = dr1["项目"].ToString();
                        dr["长"] = dr1["长"].ToString();
                        dr["宽"] = dr1["宽"].ToString();
                        dr["高"] = dr1["高"].ToString();
                        dr["箱形"] = dr1["箱形"].ToString();
                        dr["材质"] = dr1["材质"].ToString();
                        dr["数量"] = dr1["数量"].ToString();
                        dr["单价"] = dr1["单价"].ToString();
                        dtt.Rows.Add(dr);
                    }
                }
            }

            return dtt;
        }
        #endregion
    }
}
