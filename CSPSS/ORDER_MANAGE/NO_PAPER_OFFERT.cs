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
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CSPSS.ORDER_MANAGE
{
    public partial class NO_PAPER_OFFERT : Form
    {
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dt1 = new DataTable();
        basec bc=new basec ();
        #region nature
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private static string _GET_PROJECT_ID;
        public static string GET_PROJECT_ID
        {
            set { _GET_PROJECT_ID = value; }
            get { return _GET_PROJECT_ID; }
        }
        private string _EDIT;
        public string EDIT
        {
            set { _EDIT = value; }
            get { return _EDIT; }

        }
        private string _OFFER_ID;
        public string OFFER_ID
        {
            set { _OFFER_ID = value; }
            get { return _OFFER_ID; }

        }
        private string _AUDIT_STATUS;
        public string AUDIT_STATUS
        {
            set { _AUDIT_STATUS = value; }
            get { return _AUDIT_STATUS; }

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
        private string _SAMPLE_CODE;
        public string SAMPLE_CODE
        {
            set { _SAMPLE_CODE = value; }
            get { return _SAMPLE_CODE; }

        }
        private string _SAMPLE_CODE_FIRST;
        public string SAMPLE_CODE_FIRST
        {
            set { _SAMPLE_CODE_FIRST = value; }
            get { return _SAMPLE_CODE_FIRST; }

        }

        #endregion
        CEMPLOYEE_INFO cemployee_info = new CEMPLOYEE_INFO();
        private  delegate bool dele(string a1,string a2);
        private delegate void delex();
        NO_PAPER_OFFER F1 = new NO_PAPER_OFFER();
        protected int M_int_judge, i;
        protected int select;
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CNO_PAPER_OFFER cNO_PAPER_OFFER = new CNO_PAPER_OFFER();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        XizheC.CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        StringBuilder sqb = new StringBuilder();
          public NO_PAPER_OFFERT(NO_PAPER_OFFER  FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        public NO_PAPER_OFFERT()
        {
            InitializeComponent();
        }
      
        private void NO_PAPER_OFFERT_Load(object sender, EventArgs e)
        {
            //comboBox1.Text = "DBXM1604002";
            //IDO = "NP16080001";
            //comboBox4.Text = "J";
            textBox10.Multiline = true;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
          this.Icon = Resource1.xz_200X200;
            label1.Font = new Font("宋体",9,FontStyle.Bold);
            label2.Font = new Font("宋体", 9, FontStyle.Bold);
            label3.Font = new Font("宋体", 9, FontStyle.Bold);
            label9.Font = new Font("宋体", 9, FontStyle.Bold);
            label5.Font = new Font("宋体", 9, FontStyle.Bold);
            comboBox1.BackColor  = CCOLOR.CUSTOMER_YELLOW;
            comboBox4.BackColor = CCOLOR.CUSTOMER_YELLOW;
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;
            label40.Text = "";
            label41.Text = "";
            label42.Text = "";
            try
            {
                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
           
            //right();
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
            comboBox1.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox9.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox10.Text = "";
            label40.Text = "";
            label41.Text = "";
            label42.Text = "";
            comboBox4.Text = "";
        }
    
        #region bind
        private void bind()
        {
           
            textBox1.Focus();
            label40.Text = "";
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
            //this.Text = "编辑";
           
         dtx = basec.getdts(cNO_PAPER_OFFER.sql  + " where A.NPID='" + IDO + "' ORDER BY  A.NPID ASC ");
         if (dtx.Rows.Count > 0)
         {
            
             comboBox1.Text = dtx.Rows[0]["项目号"].ToString();
             comboBox4.Text = dtx.Rows[0]["报价编号"].ToString().Substring(4, 1);
             textBox1.Text = dtx.Rows[0]["项目名称"].ToString();
             dateTimePicker1.Text = dtx.Rows[0]["报价日期"].ToString();
             textBox9.Text = dtx.Rows[0]["客户名称"].ToString();
             textBox2.Text = dtx.Rows[0]["品牌"].ToString();
             textBox10.Text = dtx.Rows[0]["备注"].ToString();
             if (dtx.Rows[0]["审核状态"].ToString() == "已审核")
             {
                 label13.Text = "已审核";
                 pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/audit.png"));
             }
             else
             {
                 label13.Text = "未审核";
                 pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/61.png"));
             }
             dt = cNO_PAPER_OFFER.GetTableInfo();
             int j = 1;
             foreach (DataRow dr1 in dtx.Rows)
             {
                 DataRow dr = dt.NewRow();
                 dr["项次"] = j.ToString();
                 dr["数量"] = dr1["数量"].ToString();
                 dr["报出价"] = dr1["报出价"].ToString();
                 dr["报价编号"] = dr1["报价编号"].ToString();
                 dt.Rows.Add(dr);
                 j = j + 1;
          
             }
             if (dt.Rows.Count > 0 && dt.Rows.Count < 6)
             {
                 int n = 6 - dt.Rows.Count;
                 for (int i = 0; i < n; i++)
                 {
                     DataRow dr = dt.NewRow();
                     int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                     dr["项次"] = Convert.ToString(b1 + 1);
                     dt.Rows.Add(dr);
                 }
             }
         }
         else
         {
             dt = total1();
         }
         dataGridView1.DataSource = dt;
         dgvStateControl();

        }
        #endregion

        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cNO_PAPER_OFFER.GetTableInfo();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView1.AllowUserToAddRows = false;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            //dataGridView1.Columns["联系人"].DefaultCellStyle.BackColor = Color.Yellow;
            //dataGridView1.Columns["公司地址"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["数量"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["报出价"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["报价编号"].ReadOnly = true;
            dataGridView1.Columns["项次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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
            IDO = cNO_PAPER_OFFER.GETID();
            IFExecution_SUCCESS = false;
            bind();
            ADD_OR_UPDATE = "ADD";
            pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/61.png"));
            label13.Text = "未审核";
        }
        private void save()
        {

            btnSave.Focus();
            //dgvfoucs();
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss");
            string SAMPLE_CODE = bc.getOnlyString("SELECT SAMPLE_CODE FROM EMPLOYEEINFO WHERE EMID='"+LOGIN .EMID +"'");
            string SAMPLE_CODE_FIRST = SAMPLE_CODE.Substring(0, 1);
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "数量 IS NOT NULL");
            string v1 = bc.getOnlyString("SELECT AUDIT_STATUS FROM NO_PAPER_OFFER_MST WHERE NPID='" + IDO + "'");
            cNO_PAPER_OFFER.PIID = bc.getOnlyString("SELECT PIID FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'");//报价编号的取得要先初始化项目号
            cNO_PAPER_OFFER.MAKERID = LOGIN.EMID;
            string vproject_Id = bc.getOnlyString("SELECT PROJECT_ID FROM NO_PAPER_OFFER_MST WHERE NPID='" + IDO  + "'");
            string OFFER_TYPE_CODE = bc.getOnlyString("SELECT SUBSTRING(OFFER_ID,5,1) FROM NO_PAPER_OFFER_DET WHERE NPID='" + IDO + "'");
            if (!bc.exists(cNO_PAPER_OFFER .sql  + " WHERE A.NPID='" + IDO + "'"))
            {
                if (label13.Text == "未审核")
                {
                    cNO_PAPER_OFFER.AUDIT_STATUS = "N";

                }
                else
                {
                    cNO_PAPER_OFFER.AUDIT_STATUS = "Y";
                }
            
            }
       
            else if (v1 != "Y" && vproject_Id ==comboBox1 .Text && OFFER_TYPE_CODE ==comboBox4.Text )//项目号不变且报价类别不变，才修改原单据
            {
     
                cNO_PAPER_OFFER.AUDIT_STATUS = "N";
            }
            else
            {
                IDO = cNO_PAPER_OFFER.GETID();
           
                cNO_PAPER_OFFER.AUDIT_STATUS = "N";
              
                //AUDIT();
            }
            cNO_PAPER_OFFER.NPID = IDO;
            cNO_PAPER_OFFER.OFFER_DATE = dateTimePicker1.Text;
            cNO_PAPER_OFFER.PIID = bc.getOnlyString("SELECT PIID FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'"); 
            cNO_PAPER_OFFER.OFFER_TYPE_CODE = comboBox4.Text;
            cNO_PAPER_OFFER.PROJECT_ID = comboBox1.Text;
            cNO_PAPER_OFFER.REMARK = textBox10.Text;
            cNO_PAPER_OFFER.MAKERID = LOGIN.EMID;
            cNO_PAPER_OFFER.save(dtx);
        
            IFExecution_SUCCESS = cNO_PAPER_OFFER.IFExecution_SUCCESS;
            hint.Text = cNO_PAPER_OFFER.ErrowInfo;
            if (IFExecution_SUCCESS)
            {
                if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
                {

                    hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
                }
                else
                {
                    hint.Text = "";
                }
                bind();
                F1.load();

            }
        }
        private bool juage()
        {
      
           bool b = false;
           if (IDO ==null )
           {
               hint.Text = "编号不能为空";
               b = true;
           }
           /*if (bc.exists(cNO_PAPER_OFFER.sql + " WHERE A.PIID='" + IDO + "'") && EDIT != "有权限")
           {
               hint.Text = "本账号无修改权限！";
               b = true;
           }*/
           else if (comboBox1 .Text == "")
           {
               hint.Text = "项目号不能为空";
               b = true;
           }
           else if (!bc.exists ("SELECT * FROM PROJECT_INFO WHERE PROJECT_ID='"+comboBox1 .Text +"'"))
           {
               hint.Text = "项目号不存在系统中";
               b = true;
           }
   
           else if (comboBox4.Text == "")
           {
               hint.Text = "报价类别不能为空";
               b = true;
           }
           else if (juage2())
           {

               b = true;
           }
            return b;
        }

        #region juage2()

        private bool juage2()
        {
            bool b = false;
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "数量 IS NOT NULL ");
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    if (dr["数量"].ToString() == "")
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的数量不能为空", dr["项次"].ToString());
                        break;
                    }
                    else if (bc.yesno(dr["数量"].ToString()) == 0)
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的数量只能输入数字", dr["项次"].ToString());
                        break;
                    }
                    else if (dr["报出价"].ToString() == "")
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的报出价不能为空", dr["项次"].ToString());
                        break;
                    }
                    else if (bc.yesno(dr["报出价"].ToString()) == 0)
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的报出价只能输入数字", dr["项次"].ToString());
                        break;
                    }

                }
            }
            else
            {

                b = true;
                hint.Text = "至少有一项数量才能保存";

            }
            return b;
        }
        #endregion


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

     

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
         
            if (comboBox1.Text != "" && bc.exists("SELECT * FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'"))
            {
                DataTable dtx1 = basec.getdts(cproject_info.sql + " WHERE A.PROJECT_ID='" + comboBox1.Text + "'");
                if (dtx1.Rows.Count > 0)
                {
                    textBox1.Text = dtx1.Rows[0]["项目名称"].ToString();
                    textBox9.Text  = dtx1.Rows[0]["客户名称"].ToString();
                    textBox2.Text = dtx1.Rows[0]["品牌"].ToString();
                    textBox3.Text = dtx1.Rows[0]["AE01工号"].ToString();
                    textBox4.Text = dtx1.Rows[0]["结构01工号"].ToString();
                    textBox5.Text = dtx1.Rows[0]["平面01工号"].ToString();
                    label40.Text = dtx1.Rows[0]["AE01"].ToString();
                    label41.Text = dtx1.Rows[0]["结构01"].ToString();
                    label42.Text = dtx1.Rows[0]["平面01"].ToString();

                }
            }
            else
            {
                ClearText();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            audit();
        }
        #region audit
        private void audit()
        {
            
            try
            {
                SAMPLE_CODE = bc.getOnlyString("SELECT SAMPLE_CODE FROM EMPLOYEEINFO WHERE EMID='"+LOGIN.EMID+"'");
                SAMPLE_CODE_FIRST = SAMPLE_CODE.Substring(0, 1);
                DataTable dtx = bc.getdt(cNO_PAPER_OFFER.sql + " WHERE A.NPID='" + IDO + "'");
                int j = 1;
                if (dtx.Rows.Count > 0)
                {
                    if (juage())
                    {

                    }
                    else
                    {

                        if (label13.Text == "未审核")
                        {
                            basec.getcoms(@"UPDATE NO_PAPER_OFFER_MST SET AUDIT_STATUS='Y'  WHERE NPID='" + IDO + "'");
                            pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/audit.png"));
                            label13.Text = "已审核";
                 
                            foreach (DataRow dr in dtx.Rows)
                            {
                                sqb = new StringBuilder();
                                sqb.AppendFormat("UPDATE NO_PAPER_OFFER_DET SET ");
                                sqb.AppendFormat(" OFFER_ID='{0}' ", dr["报价编号"].ToString() + "-" + SAMPLE_CODE_FIRST);
                                sqb.AppendFormat(" WHERE NPID='{0}'AND SN='{1}'", IDO, j.ToString());
                                basec.getcoms(sqb.ToString());
                                j = j + 1;
                            }

                        }
                        else
                        {
                            basec.getcoms("UPDATE NO_PAPER_OFFER_MST SET AUDIT_STATUS='N' WHERE NPID='" + IDO + "'");
                            pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/61.png"));
                            label13.Text = "未审核";
                            j = 1;
                            foreach (DataRow dr in dtx.Rows)
                            {
                                sqb = new StringBuilder();
                                sqb.AppendFormat("UPDATE NO_PAPER_OFFER_DET SET ");
                                sqb.AppendFormat(" OFFER_ID='{0}' ",
                                    dr["报价编号"].ToString().Substring(0, dr["报价编号"].ToString().Length - 2));
                                sqb.AppendFormat(" WHERE NPID='{0}'AND SN='{1}'", IDO, j.ToString());
                                basec.getcoms(sqb.ToString());
                                j = j + 1;
                            }


                        }
                        bind();
                        F1.bind();
                    }
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
         

        }
        #endregion

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            CSPSS.OFFER_MANAGE.PROJECT_INFO FRM = new OFFER_MANAGE.PROJECT_INFO();
            FRM.WindowState = FormWindowState.Normal;
            FRM.NO_PAPER_OFFER();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = GET_PROJECT_ID;

            }
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
            int a = dataGridView1.CurrentCell.ColumnIndex;
            int b = dataGridView1.CurrentCell.RowIndex;
            int c = dataGridView1.Columns.Count - 1;
            int d = dataGridView1.Rows.Count - 1;


            if (a == c && b == d)
            {
                if (dt.Rows.Count >= 6)
                {

                    DataRow dr = dt.NewRow();
                    int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                    dr["项次"] = Convert.ToString(b1 + 1);
                    dt.Rows.Add(dr);
                }

            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            try
            {
                if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    if (bc.exists("SELECT * FROM PN_PRODUCTION_INSTRUCTIONS WHERE PFID IN (SELECT NPKEY FROM NO_PAPER_OFFER_DET WHERE NPID='" + IDO + "')"))
                    {
                        //hint.Text = "";
                        MessageBox.Show("此订单编号已经存在生产指示书作业中，不允许删除！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {

                      DataTable dtx = bc.getdt("SELECT * FROM NO_PAPER_OFFER_MST");
                      if (dtx.Rows.Count ==1)
                      {
                          basec.getcoms("DELETE NO_PAPER_OFFER_ID_NO");
                          basec.getcoms("DELETE NO_PAPER_OFFER_NO");
                      }
                      basec.getcoms("DELETE NO_PAPER_OFFER_MST WHERE NPID='" + IDO + "'");
                      basec.getcoms("DELETE NO_PAPER_OFFER_DET WHERE NPID='" + IDO + "'");
                      add();
                      F1.load();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
    }
}
