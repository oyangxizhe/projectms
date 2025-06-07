using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using XizheC;

namespace CSPSS.BASE_INFO
{
    public partial class AUDIT_LIST : Form
    {
        DataTable dt = new DataTable();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private static bool _IF_DOUBLE_CLICK;
        public static bool IF_DOUBLE_CLICK
        {
            set { _IF_DOUBLE_CLICK = value; }
            get { return _IF_DOUBLE_CLICK; }

        }
        private static string _EMPLOYEE_ID;
        public static string EMPLOYEE_ID
        {
            set { _EMPLOYEE_ID = value; }
            get { return _EMPLOYEE_ID; }

        }
        private static string _ENAME;
        public static string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        basec bc = new basec();

        StringBuilder sqb = new StringBuilder();
        protected int M_int_judge, i;
        protected int select;
        CAUDIT_LIST caudit_list = new CAUDIT_LIST();
        public AUDIT_LIST()
        {
            InitializeComponent();
        }
        #region double_click
        private void dgvEmployeeInfo_DoubleClick(object sender, EventArgs e)
        {
            
        }
        #endregion

 

        private void bind()
        {
           
          
            dt = basec.getdts(caudit_list .sql );

            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["纸品生产"].ToString()))
                {
                    comboBox1.Text = dt.Rows[0]["纸品生产"].ToString() + "-" + dt.Rows[0]["纸品生产工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["木铁生产工号"].ToString()))
                {
                    comboBox2.Text = dt.Rows[0]["木铁生产"].ToString() + "-" + dt.Rows[0]["木铁生产工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["亚克力生产工号"].ToString()))
                {
                    comboBox3.Text = dt.Rows[0]["亚克力生产"].ToString() + "-" + dt.Rows[0]["亚克力生产工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["纸品计划工号"].ToString()))
                {
                    comboBox4.Text = dt.Rows[0]["纸品计划"].ToString() + "-" + dt.Rows[0]["纸品计划工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["木铁计划工号"].ToString()))
                {
                    comboBox5.Text = dt.Rows[0]["木铁计划"].ToString() + "-" + dt.Rows[0]["木铁计划工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["结构设计工号"].ToString()))
                {
                    comboBox6.Text = dt.Rows[0]["结构设计"].ToString() + "-" + dt.Rows[0]["结构设计工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["平面设计工号"].ToString()))
                {

                    comboBox7.Text = dt.Rows[0]["平面设计"].ToString() + "-" + dt.Rows[0]["平面设计工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["纸品采购工号"].ToString()))
                {
                    comboBox8.Text = dt.Rows[0]["纸品采购"].ToString() + "-" + dt.Rows[0]["纸品采购工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["木铁采购工号"].ToString()))
                {
                    comboBox9.Text = dt.Rows[0]["木铁采购"].ToString() + "-" + dt.Rows[0]["木铁采购工号"].ToString();
                }

            }
            else
            {
                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
                comboBox4.Text = "";
                comboBox5.Text = "";
                comboBox6.Text = "";
                comboBox7.Text = "";
                comboBox8.Text = "";
                comboBox9.Text = "";
            }
      
            hint.Location = new Point(256, 136);
            hint.ForeColor = Color.Red;
           
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
        }
     
    
        #region save
        private void save()
        {
           
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss");
            string EMID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='"+comboBox1.Text  +"'");
            string varMakerID = LOGIN.EMID;
            string sqlo = @"
INSERT INTO AUDIT_LIST(
PAPER_PRODUCTION_AUDIT_MAKERID,
WOOD_IRON_PRODUCTION_AUDIT_MAKERID,
ACRYLIC_PRODUCTION_AUDIT_MAKERID,
PAPER_PLAN_AUDIT_MAKERID,
WOOD_IRON_PLAN_AUDIT_MAKERID,
STRUCTURE_AUDIT_MAKERID,
PLANE_AUDIT_MAKERID,
PAPER_PURCHASE_AUDIT_MAKERID,
WOOD_IRON_PURCHASE_AUDIT_MAKERID,
MakerID,
Date)
VALUES
(
@PAPER_PRODUCTION_AUDIT_MAKERID,
@WOOD_IRON_PRODUCTION_AUDIT_MAKERID,
@ACRYLIC_PRODUCTION_AUDIT_MAKERID,
@PAPER_PLAN_AUDIT_MAKERID,
@WOOD_IRON_PLAN_AUDIT_MAKERID,
@STRUCTURE_AUDIT_MAKERID,
@PLANE_AUDIT_MAKERID,
@PAPER_PURCHASE_AUDIT_MAKERID,
@WOOD_IRON_PURCHASE_AUDIT_MAKERID,
@MakerID,
@Date

)
";
            string v1 = "", v2 = "", v3 = "", v4 = "", v5 = "", v6 = "", v7 = "", v8 = "", v9 = "";
            if (comboBox1.Text != "")
            {
                v1 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox1.Text, '-') + "'");
            }
            if (comboBox2.Text != "")
            {
                v2 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox2.Text, '-') + "'");
            }
            if (comboBox3.Text != "")
            {
                v3 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox3.Text, '-') + "'");
            }
            if (comboBox4.Text != "")
            {
                v4 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox4.Text, '-') + "'");
            }
            if (comboBox5.Text != "")
            {
                v5 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox5.Text, '-') + "'");
            }
            if (comboBox6.Text != "")
            {
                v6 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox6.Text, '-') + "'");
            }
            if (comboBox7.Text != "")
            {
                v7 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox7.Text, '-') + "'");
            }
            if (comboBox8.Text != "")
            {
                v8 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox8.Text, '-') + "'");
            }
            if (comboBox9.Text != "")
            {
                v9 = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" +
              bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox9.Text, '-') + "'");
            }
            if (v1 == "" && v2 == "" && v3 == "" && v4 == "" && v5 == "" && v6 == "" && v7 == "" && v8 == "" && v9 == "")
            {
                hint.Text = "至少有一项不为空才能保存";
            }
            else
            {
                basec.getcoms("DELETE AUDIT_LIST");
                SqlConnection sqlcon = bc.getcon();
                sqlcon.Open();
                SqlCommand sqlcom = new SqlCommand(sqlo, sqlcon);
                sqlcom.Parameters.Add("@PAPER_PRODUCTION_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v1;
                sqlcom.Parameters.Add("@WOOD_IRON_PRODUCTION_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v2;
                sqlcom.Parameters.Add("@ACRYLIC_PRODUCTION_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v3;
                sqlcom.Parameters.Add("@PAPER_PLAN_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v4;
                sqlcom.Parameters.Add("@WOOD_IRON_PLAN_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v5;
                sqlcom.Parameters.Add("@STRUCTURE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v6;
                sqlcom.Parameters.Add("@PLANE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v7;
                sqlcom.Parameters.Add("@PAPER_PURCHASE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v8;
                sqlcom.Parameters.Add("@WOOD_IRON_PURCHASE_AUDIT_MAKERID", SqlDbType.VarChar, 20).Value = v9;
                sqlcom.Parameters.Add("@MakerID", SqlDbType.VarChar, 20).Value = varMakerID;
                sqlcom.Parameters.Add("@Date", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
                IFExecution_SUCCESS = true;
                bind();
            }
        }
        #endregion
        #region juage()
        private bool juage()
        {

            bool b = false;
           if (comboBox1 .Text!="" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", 
                bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox1.Text, '-'))))
            {
                b = true;
                hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox1.Text, '-') + " 不存在系统";
            }
           else if (comboBox2 .Text!="" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
            bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox2.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox2.Text, '-') + " 不存在系统";
           }
           else if (comboBox3.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox3.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox3.Text, '-') + " 不存在系统";
           }
           else if (comboBox4.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox4.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox4.Text, '-') + " 不存在系统";
           }
           else if (comboBox5.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox5.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox5.Text, '-') + " 不存在系统";
           }
           else if (comboBox6.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox6.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox6.Text, '-') + " 不存在系统";
           }
           else if (comboBox7.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox7.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox7.Text, '-') + " 不存在系统";
           }
           else if (comboBox8.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox8.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox8.Text, '-') + " 不存在系统";
           }
           else if (comboBox9.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",
     bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox9.Text, '-'))))
           {
               b = true;
               hint.Text = "员工工号 " + bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox9.Text, '-') + " 不存在系统";
           }
            return b;

        }
        #endregion
        public void ClearText()
        {
           
         
        
        }


        private void add()
        {
            ClearText();
            
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            
            if (juage())
            {

            }
            else
            {
                save();
                if (IFExecution_SUCCESS)
                {
                    add();
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {


           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
       
        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = ENAME + "-" + EMPLOYEE_ID;
                
            }
        }
        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox2.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox2.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox2.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox2.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox3_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox3.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox3.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox3.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox3.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox4_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox4.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox4.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox4.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox4.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox5_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox5.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox5.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox5.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox5.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox6_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox6.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox6.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox6.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox6.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox7_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox7.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox7.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox7.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox7.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox8_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox8.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox8.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox8.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox8.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void comboBox9_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.AUDIT_LIST_USE();
            FRM.ShowDialog();
            this.comboBox9.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox9.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox9.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox9.Text = ENAME + "-" + EMPLOYEE_ID;

            }
        }

        private void AUDIT_LIST_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            bind();

        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                  
                        basec.getcoms("DELETE AUDIT_LIST" );
                        bind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

     

   
  
    }
}
