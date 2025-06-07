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
    public partial class OTHER_COST : Form
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
        basec bc = new basec();
        COTHER_COST cOTHER_COST = new COTHER_COST();
        CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
        StringBuilder sqb = new StringBuilder();
        protected int M_int_judge, i;
        protected int select;
   
        public OTHER_COST()
        {
            InitializeComponent();
        }
        private void DEPAET_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            if (Screen.AllScreens[0].Bounds.Width == 1366 && Screen.AllScreens[0].Bounds.Height == 768 ||
                 Screen.AllScreens[0].Bounds.Width == 1280 && Screen.AllScreens[0].Bounds.Height == 800)
            {
              

            }
            else if (Screen.AllScreens[0].Bounds.Width == 1920 && Screen.AllScreens[0].Bounds.Height == 1080)
            {

            }
            else
            {
                this.AutoScroll = true;
                this.AutoScrollMinSize = new Size(1920, 1080);
            }
       
            try
            {
                bind();//此作业数据太多，为避免加载过慢，开启作业不做加载动作 160419
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"提示",MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            label14.Text = "本作业录入与修改查询共用输入窗口";
            label14.ForeColor  = CCOLOR.shop2;
         
         
        }
        private void bind()
        {
            sqb = new StringBuilder(cOTHER_COST.sql);
            //sqb.AppendFormat(" WHERE DateDiff(day,A.DATE,getdate()) >-1 and DateDiff(day,A.DATE,getdate()) <+1");
            sqb.AppendFormat(" ORDER BY C.CNAME,A.BRAND ASC");
            dt = basec.getdts(sqb.ToString ());
            dt = cOTHER_COST.RETURN_HAVE_ID_DT(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.AllowUserToAddRows = false;
            textBox2.Focus();
            textBox2.BackColor = Color.Yellow;
            textBox3.BackColor = Color.Yellow;
            textBox3.TextAlign = HorizontalAlignment.Right;
            dgvStateControl();
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
            DataTable dtx = bc.getdt("SELECT * FROM CUSTOMERINFO_MST");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {

                    comboBox1.Items.Add(dr["CNAME"].ToString());
                }
            }
            dtx = bc.getdt(ccustomer_info.sql);
            if (dtx.Rows.Count > 0)
            {
                comboBox2.Items.Clear();
                comboBox2.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {

                    comboBox2.Items.Add(dr["品牌"].ToString());
                }
            }
             
      
        }
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
               
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].ReadOnly = true;

            }
            dataGridView1.Columns["制单人"].Width = 70;
            dataGridView1.Columns["客户比例"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["编号"].Visible = false;
         
        }
        #endregion
    
        #region save
        private void save()
        {
            cOTHER_COST.OCID = IDO;
            cOTHER_COST.PROJECT_NAME  = textBox2.Text;
            cOTHER_COST.CUID = bc.getOnlyString("SELECT CUID FROM CUSTOMERINFO_MST WHERE CNAME='"+comboBox1 .Text +"'");
            cOTHER_COST.BRAND = comboBox2.Text;
            cOTHER_COST.CUSTOMER_PERCENT = textBox3.Text;
            cOTHER_COST.REMARK  = textBox4.Text;
            cOTHER_COST.MAKERID = LOGIN.EMID;
            cOTHER_COST.save();
        }
        #endregion
        #region juage()
        private bool juage()
        {
            bool b = false;
            if (IDO == "")
            {
                b = true;
                hint.Text = "编号不能为空！";
            }
            else if (textBox2.Text == "")
            {
                b = true;
                hint.Text = "项目不能为空！";
            }
            else if (comboBox1.Text == "")
            {
                b = true;
                hint.Text = "客户不能为空";
            }
            else if (comboBox2.Text == "")
            {
                b = true;
                hint.Text = "品牌不能为空";
            }
            else if (textBox3.Text == "")
            {
                b = true;
                hint.Text = "客户比例不能为空";
            }
            else if (textBox3.Text != "" && bc.yesno (textBox3 .Text )==0)
            {
                b = true;
                hint.Text = "客户比例只能输入数字";
            }
         
     
            return b;

        }
        #endregion
        public void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
           
        }
        private void add()
        {
            ClearText();
            IDO= cOTHER_COST.GETID();
            textBox2.Focus();
            bind();

        }
        private void btnSave_Click(object sender, EventArgs e)
        {
   
            try
            {
                hint.Text = "";
                if (juage())
                {

                }
                else
                {

                    save();
                    IFExecution_SUCCESS = cOTHER_COST.IFExecution_SUCCESS;
                    hint.Text = cOTHER_COST.ErrowInfo;

                    if (IFExecution_SUCCESS)
                    {
                        add();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                sqb = new StringBuilder(cOTHER_COST.sql);
                sqb.AppendFormat("WHERE  A.PROJECT_NAME LIKE '%{0}%'", textBox2.Text);
                sqb.AppendFormat(" AND C.CNAME LIKE '%{0}%'", comboBox1 .Text );
                sqb.AppendFormat(" AND A.BRAND LIKE '%{0}%'", comboBox2.Text );
                sqb.AppendFormat(" AND A.CUSTOMER_PERCENT LIKE '%{0}%'", textBox3.Text);
                sqb.AppendFormat(" AND A.REMARK LIKE '%{0}%'", textBox4.Text);
                sqb.AppendFormat(" ORDER BY C.CNAME,A.BRAND ASC");
                dt = bc.getdt(sqb.ToString ());
                dt = cOTHER_COST.RETURN_HAVE_ID_DT(dt);
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dgvStateControl();
                }
                else
                {
                    hint.Text = "没有找到相关信息！";
                    dataGridView1.DataSource = null;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            string id = dt.Rows[dataGridView1.CurrentCell.RowIndex]["项目"].ToString();
            try
            {
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM OTHER_COST WHERE OCID='" +IDO+ "'";
                basec.getcoms(strSql);
                IDO = cOTHER_COST.GETID();
                bind();
                ClearText();
            }
            catch (Exception)
            {


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

                    dataGridView1.Focus();

                    return true;
                }
            
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            hint.Text = "";
            try
            {
                string v1 = dt.Rows[dataGridView1.CurrentCell.RowIndex]["项目"].ToString();
                int i = dataGridView1.CurrentCell.RowIndex;
                IDO = dt.Rows[i]["编号"].ToString();
                textBox2.Text = dt.Rows[i]["项目"].ToString();
                comboBox1.Text = dt.Rows[i]["客户"].ToString();
                comboBox2.Text = dt.Rows[i]["品牌"].ToString();
                textBox3.Text = bc.RETURN_UNTIL_CHAR(dt.Rows[i]["客户比例"].ToString(), '%');
                textBox4.Text = dt.Rows[i]["说明"].ToString();
             
            }
            catch (Exception)
            {
               // MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1,this .Text );

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) //判断是不是右键
            {
                Control control = new Control();
                Point ClickPoint = new Point(e.X, e.Y);
                control.GetChildAtPoint(ClickPoint);
                if (dataGridView1.HitTest(e.X, e.Y).RowIndex >= 0 && dataGridView1.HitTest(e.X, e.Y).ColumnIndex >= 0)//判断你点的是不是一个信息行里
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.HitTest(e.X, e.Y).RowIndex].Cells[dataGridView1.HitTest(e.X, e.Y).ColumnIndex];
                    ContextMenu con = new ContextMenu();
                    MenuItem menuDeleteknowledge = new MenuItem("复制");
                    menuDeleteknowledge.Click += new EventHandler(btndgvInfoCopy_Click);
                    con.MenuItems.Add(menuDeleteknowledge);
                    this.dataGridView1.ContextMenu = con;
                    con.Show(dataGridView1, new Point(e.X + 10, e.Y));
                }
            }
        }
        private void btndgvInfoCopy_Click(object sender, EventArgs e)
        {
            dgvCopy(ref dataGridView1);
        }
        private void dgvCopy(ref DataGridView dgv)
        {
            if (dgv.GetCellCount(DataGridViewElementStates.Selected) > 0)
            {
                try
                {
                    Clipboard.SetDataObject(dgv.GetClipboardContent());
                }
                catch (Exception MyEx)
                {
                    MessageBox.Show(MyEx.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
     
        }

        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            try
            {
                DataTable dtx = bc.getdt(ccustomer_info.sql + " WHERE B.CNAME='" + comboBox1.Text + "'");
                if (dtx.Rows.Count > 0)
                {
                    comboBox2.Items.Clear();
                    comboBox2.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                    
                        comboBox2.Items.Add(dr["品牌"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
    }
}
