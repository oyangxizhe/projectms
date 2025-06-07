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

namespace CSPSS.OFFER_MANAGE
{
    public partial class PROJECT_INFOT : Form
    {
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dt1 = new DataTable();
        basec bc=new basec ();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _EDIT;
        public string EDIT
        {
            set { _EDIT = value; }
            get { return _EDIT; }

        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private static string _CUID;
        public static string CUID
        {
            set { _CUID = value; }
            get { return _CUID; }
        }
        private static string _CNAME;
        public static string CNAME
        {
            set { _CNAME = value; }
            get { return _CNAME; }
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

        private static string _CO_WAREID;
        public static string CO_WAREID
        {
            set { _CO_WAREID = value; }
            get { return _CO_WAREID; }

        }
        CEMPLOYEE_INFO cemployee_info = new CEMPLOYEE_INFO();
        private  delegate bool dele(string a1,string a2);
        private delegate void delex();
        PROJECT_INFO F1 = new PROJECT_INFO();
        protected int M_int_judge, i;
        protected int select;
        CPROJECT_INFO cPROJECT_INFO = new CPROJECT_INFO();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
          public PROJECT_INFOT(PROJECT_INFO  FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        public PROJECT_INFOT()
        {
            InitializeComponent();
        }
      
        private void PROJECT_INFOT_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            label1.Font = new Font("宋体",9,FontStyle.Bold);
            label2.Font = new Font("宋体", 9, FontStyle.Bold);
            label3.Font = new Font("宋体", 9, FontStyle.Bold);
            label9.Font = new Font("宋体", 9, FontStyle.Bold);
            label5.Font = new Font("宋体", 9, FontStyle.Bold);
            textBox2.BackColor = CCOLOR.CUSTOMER_YELLOW;
            comboBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
            comboBox3.BackColor = CCOLOR.CUSTOMER_YELLOW;
            comboBox2.BackColor = CCOLOR.CUSTOMER_YELLOW;
            bind();
            right();
        }

        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("项目新增", LOGIN.USID);
            btnAdd.Visible = false;
            btnSave.Visible = false;
            label15.Visible = false;
            label17.Visible = false;
            if (dtx.Rows.Count > 0)
            {
                if (dtx.Rows[0]["新增权限"].ToString() == "有权限")
                {
                    btnAdd.Visible = true;
                    btnSave.Visible = true;
                    label15.Visible = true;
                    label17.Visible = true;
                }
                if (dtx.Rows[0]["修改权限"].ToString() == "有权限")
                {
                    btnSave.Visible = true;
                    label15.Visible = true;
                    EDIT = "有权限";
                }
            }
        }
        #endregion
        public void ClearText()
        {
            textBox2.Text = "";
            comboBox1.Text = "";
            comboBox11.Text = "";
            textBox1.Text = "";
        }
    
        #region bind
        private void bind()
        {
        
            textBox2.Focus();
         
            label40.Text = "";
            label46.Text = "";
            label43.Text = "";
            label42.Text = "";
            label48.Text = "";
            label44.Text = "";
            label41.Text = "";
            label45.Text = "";
            label47.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox8.Text = "";
            comboBox9.Text = "";
            comboBox10.Text = "";

            hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
            
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {

                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
            DataTable dtx = basec.getdts(cPROJECT_INFO.sql +" WHERE A.PIID='"+IDO  +"'");
            if (dtx.Rows.Count > 0)
            {

                textBox1.Text = dtx.Rows[0]["项目号"].ToString();
                textBox2.Text = dtx.Rows[0]["项目名称"].ToString();
                comboBox1.Text = dtx.Rows[0]["客户名称"].ToString();

                comboBox2.Text = dtx.Rows[0]["品牌"].ToString();

                comboBox3.Text = dtx.Rows[0]["AE01工号"].ToString();
                comboBox4.Text = dtx.Rows[0]["AE02工号"].ToString();

                comboBox5.Text = dtx.Rows[0]["AE03工号"].ToString();

                comboBox6.Text = dtx.Rows[0]["平面01工号"].ToString();
                comboBox7.Text = dtx.Rows[0]["平面02工号"].ToString();
                comboBox8.Text = dtx.Rows[0]["平面03工号"].ToString();

                comboBox9.Text = dtx.Rows[0]["结构01工号"].ToString();
                comboBox10.Text = dtx.Rows[0]["结构02工号"].ToString();
                comboBox11.Text = dtx.Rows[0]["结构03工号"].ToString();

                label40.Text = dtx.Rows[0]["AE01"].ToString();
                label46.Text = dtx.Rows[0]["AE02"].ToString();
                label43.Text = dtx.Rows[0]["AE03"].ToString();

                label42.Text = dtx.Rows[0]["平面01"].ToString();
                label48.Text = dtx.Rows[0]["平面02"].ToString();
                label45.Text = dtx.Rows[0]["平面03"].ToString();

                label41.Text = dtx.Rows[0]["结构01"].ToString();
                label44.Text = dtx.Rows[0]["结构02"].ToString();
                label47.Text = dtx.Rows[0]["结构03"].ToString();
            }
     
            this.Text = "项目信息编辑";
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            dtx = bc.getdt("SELECT A.CNAME AS CNAME FROM CUSTOMERINFO_MST A LEFT JOIN CUSTOMERINFO_DET B ON A.CUID=B.CUID  WHERE B.USER_MAKERID='"+LOGIN .EMID +"'");
            dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "CNAME");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Items.Clear();
                foreach (DataRow dr in dtx.Rows)
                {

                    string suggestWord = dr["VALUE"].ToString();
                    inputInfoSource.Add(suggestWord);
                    comboBox1.Items.Add(suggestWord);
                }
            }
         
        
            this.comboBox1 .AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox1 .AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.comboBox1 .AutoCompleteCustomSource = inputInfoSource;

        }
        #endregion
  

        private void btnEdit_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
            M_int_judge = 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
     
            try
            {
                btnSave.Focus();
                if (juage())
                {
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    save();
                    if (IFExecution_SUCCESS == true)
                    {
                        //add();
                        BASE_INFO.SUN_SCREEN FRM = new CSPSS.BASE_INFO.SUN_SCREEN(this);
                        FRM.PROJECT_ID = textBox1.Text;
                        FRM.PROJECT_NAME = textBox2.Text;
                        FRM.Show();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void add()
        {
            ClearText();
            IDO = cPROJECT_INFO.GETID();
            IFExecution_SUCCESS = false;
            bind();
            ADD_OR_UPDATE = "ADD";
        }
        private void save()
        {

            btnSave.Focus();
            //dgvfoucs();
            cPROJECT_INFO.EMID = LOGIN.EMID;
            cPROJECT_INFO.PIID = IDO;
            cPROJECT_INFO.PROJECT_ID = cPROJECT_INFO.GETID_PROJECT_ID();
            cPROJECT_INFO.PROJECT_NAME = textBox2.Text;
            cPROJECT_INFO.CUID = bc.getOnlyString("SELECT CUID FROM CUSTOMERINFO_MST WHERE CNAME='" + comboBox1.Text + "'");
            cPROJECT_INFO.BRAND = comboBox2.Text;
            cPROJECT_INFO.AE_MAKERID_ONE = bc.RETURN_EMID(comboBox3.Text);
            cPROJECT_INFO.AE_MAKERID_TWO = bc.RETURN_EMID(comboBox4.Text);
            cPROJECT_INFO.AE_MAKERID_THREE = bc.RETURN_EMID(comboBox5.Text);
            cPROJECT_INFO.PLANE_MAKERID_ONE = bc.RETURN_EMID(comboBox6.Text);
            cPROJECT_INFO.PLANE_MAKERID_TWO = bc.RETURN_EMID(comboBox7.Text);
            cPROJECT_INFO.PLANE_MAKERID_THREE = bc.RETURN_EMID(comboBox8.Text);
            cPROJECT_INFO.STRUCTURE_MAKERID_ONE = bc.RETURN_EMID(comboBox9.Text);
            cPROJECT_INFO.STRUCTURE_MAKERID_TWO = bc.RETURN_EMID(comboBox10.Text);
            cPROJECT_INFO.STRUCTURE_MAKERID_THREE = bc.RETURN_EMID(comboBox11.Text);
            cPROJECT_INFO.save();
            IFExecution_SUCCESS = cPROJECT_INFO.IFExecution_SUCCESS;
            hint.Text = cPROJECT_INFO.ErrowInfo;
            if (IFExecution_SUCCESS)
            {

                bind();
            }
            F1.bind();
        }
        private bool juage()
        {
            
           bool b = false;
           if (IDO ==null )
           {
               hint.Text = "编号不能为空";
               b = true;
           }
           if (bc.exists(cPROJECT_INFO.sql + " WHERE A.PIID='" + IDO + "'") && EDIT != "有权限")
           {
               hint.Text = "本账号无修改权限！";
               b = true;
           }
           else if (textBox2 .Text== "")
           {
               hint.Text = "项目名称不能为空";
               b = true;
           }
           else if (comboBox1 .Text == "")
           {
               hint.Text = "客户不能为空";
               b = true;
           }
           else if (!bc.exists ("SELECT * FROM CUSTOMERINFO_MST WHERE CNAME='"+comboBox1 .Text +"'"))
           {
               hint.Text = "客户不存在系统中";
               b = true;
           }
        
           else if (juage3())
           {

               b = true;
           }
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE PIID='{0}'",bc.RETURN_PIID(textBox2 .Text ))))
            {
                hint.Text = string.Format("尺寸 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
            return b;
        }

        #region juage3()

        private bool juage3()
        {
            bool b = false;
                 if (comboBox2.Text  == "")
                {

                    b = true;
                    hint.Text = string.Format("品牌不能为空");

                }
               else  if (!bc.exists(string.Format("SELECT * FROM CUSTOMERINFO_DET WHERE BRAND='{0}'", comboBox2.Text )))
                {

                    b = true;
                    hint.Text = string.Format("品牌不存在系统中");

                }
               else  if (comboBox3.Text  == "")
                {

                    b = true;
                    hint.Text = string.Format("AE工号不能为空");

                }
                else if (!bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", comboBox3.Text )))
                {

                    b = true;
                    hint.Text = string.Format("AE工号不存在系统中");

                }
                else if (comboBox4.Text!= "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",comboBox4.Text )))
                {
                  
                    b = true;
                    hint.Text = string.Format("AE助理-1工号不存在系统中");

                }
                else if (comboBox5.Text != "" &&
                    !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", comboBox5.Text )))
                {

                    b = true;
                    hint.Text = string.Format("AE助理-2工号不存在系统中");

                }
            
                else if (comboBox6.Text !="" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", comboBox6.Text )))
                {

                    b = true;
                    hint.Text = string.Format("平面设计工号不存在系统中");

                }
                else if (comboBox7.Text  != "" &&
                    !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", comboBox7.Text )))
                {

                    b = true;
                    hint.Text = string.Format("平面设计助理-1工号不存在系统中");

                }
                else if (comboBox8.Text != "" &&
              !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", comboBox8.Text )))
                {

                    b = true;
                    hint.Text = string.Format("平面设计助理-2工号不存在系统中");

                }
          
                else if (comboBox9.Text !="" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",comboBox9.Text )))
                {

                    b = true;
                    hint.Text = string.Format("结构设计工号不存在系统中");

                }
                else if (comboBox10.Text != "" &&
                    !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",comboBox10.Text )))
                {

                    b = true;
                    hint.Text = string.Format("结构设计助理-1工号不存在系统中");

                }
                else if (comboBox11.Text  != "" &&
              !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",comboBox11.Text )))
                {

                    b = true;
                    hint.Text = string.Format("结构设计助理-2工号不存在系统中");

                }
            

            return b;
        }
        #endregion
        private void btnDel_Click(object sender, EventArgs e)
        {
           
           /* if (bc.exists(string.Format("SELECT * FROM WORKORDER_MST WHERE PIID='{0}'", bc.RETURN_PIID(textBox2.Text))))
            {
                hint.Text = string.Format("尺寸 {0} 已经在工单中使用不允许删除", textBox2.Text);
             
            }
            else
            {
               
            }*/
                 if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                  
                    basec.getcoms("DELETE PROJECT_INFO_DET WHERE PIID='"+textBox1 .Text +"'");
                    basec.getcoms("DELETE PROJECT_INFO_MST WHERE PIID='" +textBox1 .Text + "'");
                    bind();
                    ClearText();
                    textBox1.Text = "";
                 
                }
         
            try
            {
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && ((!(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)))
            {
                SendKeys.SendWait("{Tab}");
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");

                return true;
            }
            if (keyData == (Keys.F7))
            {

                //double_info();

                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
    


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }


        private void btnSearch_Click(object sender, EventArgs e)
        {
            bind();
            
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }
        private void comboBox1_DropDown(object sender, EventArgs e)
        {
     
        }
        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                dt = bc.getdt(ccustomer_info.sql + " WHERE B.CNAME='"+comboBox1 .Text +"'");
                dt = bc.RETURN_NOHAVE_REPEAT_DT(dt, "品牌");
                if (dt.Rows.Count > 0)
                {
                    comboBox2.Items.Clear();
                    comboBox2.Items.Add("");
                    foreach (DataRow dr in dt.Rows)
                    {
                        comboBox2.Items.Add(dr["VALUE"].ToString());


                    }

                }
            }
        }
        private void comboBox3_DropDown(object sender, EventArgs e)
        {
            try
            {
                BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
                FRM.PROJECT_INFO_USE();
                FRM.IDO = cemployee_info.GETID();
                FRM.POSITION = "AE";
                FRM.ShowDialog();
                this.comboBox3.IntegralHeight = false;//使组合框不调整大小以显示其所有项
                this.comboBox3.DroppedDown = false;//使组合框不显示其下拉部分
                this.comboBox3.IntegralHeight = true;//恢复默认值
                if (IF_DOUBLE_CLICK)
                {
                    comboBox3.Text = EMPLOYEE_ID;
                    label40.Text = ENAME;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void comboBox4_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "AE";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox4.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox4.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox4.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox4.Text = EMPLOYEE_ID;
                label46.Text = ENAME;
            }
        }

        private void comboBox5_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "AE";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox5.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox5.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox5.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox5.Text = EMPLOYEE_ID;
                label43.Text = ENAME;
            }
        }

        private void comboBox6_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "平面";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox6.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox6.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox6.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox6.Text = EMPLOYEE_ID;
                label42.Text = ENAME;
            }
        }

        private void comboBox7_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "平面";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox7.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox7.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox7.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox7.Text = EMPLOYEE_ID;
                label48.Text = ENAME;
            }
        }

        private void comboBox8_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "平面";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox8.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox8.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox8.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox8.Text = EMPLOYEE_ID;
                label45.Text = ENAME;
            }
        }

        private void comboBox9_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "结构";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox9.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox9.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox9.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox9.Text = EMPLOYEE_ID;
                label41.Text = ENAME;
            }
        }

        private void comboBox10_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "结构";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox10.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox10.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox10.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox10.Text = EMPLOYEE_ID;
                label44.Text = ENAME;
            }
        }

        private void comboBox11_DropDown(object sender, EventArgs e)
        {
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.PROJECT_INFO_USE();
            FRM.POSITION = "结构";
            FRM.IDO = cemployee_info.GETID();
            FRM.ShowDialog();
            this.comboBox11.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox11.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox11.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox11.Text = EMPLOYEE_ID;
                label47.Text = ENAME;
            }
        }
    }
}
