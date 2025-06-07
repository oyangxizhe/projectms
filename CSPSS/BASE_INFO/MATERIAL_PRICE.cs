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
    public partial class MATERIAL_PRICE : Form
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
        CMATERIAL_PRICE cMATERIAL_PRICE = new CMATERIAL_PRICE();
        protected string sql = @"
SELECT 
A.MATERIAL_TYPE AS 类型,
A.STARTING_PRICE 起步价,
A.STARTING_PRICE_UNIT AS 起步价单位,
A.UNIT_PRICE AS 单位计价,
A.UNIT_PRICE_UNIT AS 单位计价单位,
A.MAX_PRICE AS 封顶金额,
A.MAX_PRICE_UNIT AS 封顶金额单位,
(SELECT ENAME FROM EMPLOYEEINFO 
WHERE EMID=A.MAKERID ) AS 制单人,
A.DATE AS 制单日期
FROM
MATERIAL_PRICE A ";
   
        protected int M_int_judge, i;
        protected int select;
        public MATERIAL_PRICE()
        {
            InitializeComponent();
        }
     

        private void DEPAET_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            textBox1.Text = IDO;
            bind();
            //this.WindowState = FormWindowState.Maximized;

        }

        private void bind()
        {
           
            dt = basec.getdts(sql);
            dataGridView1.DataSource = dt;
            dataGridView1.AllowUserToAddRows = false;
            textBox2.Focus();
            textBox2.BackColor = Color.Yellow;
            dgvStateControl();
            hint.Location = new Point(256, 136);
            hint.ForeColor = Color.Red;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            DataTable dtx = bc.getdt("SELECT * FROM UNIT");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("");
                comboBox2.Items.Clear();
                comboBox2.Items.Add("");
                comboBox3.Items.Clear();
                comboBox3.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {
                    comboBox1.Items.Add(dr["UNIT"].ToString());
                    comboBox2.Items.Add(dr["UNIT"].ToString());
                    comboBox3.Items.Add(dr["UNIT"].ToString());
                }
            }
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
          
            label14.Text = "类型：";
            groupBox1.Text = "类型信息";
            label1.Text = "编号";
            label2.Text = "类型";
            this.Text = "类型信息";
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
               
                dataGridView1.Columns["类型"].Width = 120;
                dataGridView1.Columns["制单人"].Width = 80;
                dataGridView1.Columns["制单日期"].Width = 120;
            
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
            dataGridView1.Columns["起步价"].HeaderText = "起步价(小POP)";
            dataGridView1.Columns["单位计价"].HeaderText = "单位计价(陈列架)";
            dataGridView1.Columns["封顶金额"].HeaderText = "封顶金额(堆头)";

            dataGridView1.Columns["起步价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["单位计价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["封顶金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

        }
        #endregion
    
        #region save
        private void save()
        {
            cMATERIAL_PRICE.MRID = textBox1.Text;
            cMATERIAL_PRICE.MATERIAL_TYPE = textBox2.Text;
            cMATERIAL_PRICE.STARTING_PRICE = textBox3.Text;
            cMATERIAL_PRICE.STARTING_PRICE_UNIT = comboBox1.Text;
            cMATERIAL_PRICE.UNIT_PRICE = textBox4.Text;
            cMATERIAL_PRICE.UNIT_PRICE_UNIT = comboBox2.Text;
            cMATERIAL_PRICE.MAX_PRICE = textBox6.Text;
            cMATERIAL_PRICE.MAX_PRICE_UNIT = comboBox3.Text;
            cMATERIAL_PRICE.MAKERID = LOGIN.EMID;
            cMATERIAL_PRICE.save();
       
           
        }
        #endregion
        #region juage()
        private bool juage()
        {


            bool b = false;
            if (textBox1.Text == "")
            {
                b = true;
                hint.Text = "编号不能为空！";
            }
            else if (textBox2.Text == "")
            {
                b = true;
                hint.Text = "类型不能为空！";
            }
            else if (textBox3.Text == "" && textBox4.Text == "" && textBox6.Text == "")
            {
                b = true;
                hint.Text = "起步价 单位计价 封顶金额至少要输入一项";
            }
            else if (textBox3.Text != "" && bc.yesno (textBox3 .Text )==0 ||  textBox4.Text != "" && bc.yesno (textBox4 .Text )==0 || 
                textBox6.Text != "" && bc.yesno (textBox6 .Text )==0)
            {
                b = true;
                hint.Text = "起步价 单位计价 封顶金额只能输入数字";
            }
            else if (textBox3.Text != "" && comboBox1 .Text =="" || textBox4.Text != "" && comboBox2 .Text =="" ||
                textBox6.Text != "" && comboBox3 .Text =="")
            {
                b = true;
                hint.Text = "起步价 单位计价 封顶金额有值时需填写单位";
            }
            return b;

        }
        #endregion
        public void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox6.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
        }
 
        private void add()
        {
            ClearText();
            textBox1.Text = cMATERIAL_PRICE.GETID();
            textBox2.Focus();
            bind();

        }
      

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (juage())
                {

                }
                else
                {

                    save();
                    IFExecution_SUCCESS = cMATERIAL_PRICE.IFExecution_SUCCESS;
                    hint.Text = cMATERIAL_PRICE.ErrowInfo;
                    if (IFExecution_SUCCESS)
                    {
                        add();
                    }
                }
            }
            catch (Exception)
            {

            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {


                dt = bc.getdt(sql+" WHERE  A.MATERIAL_PRICE LIKE '%"+textBox5 .Text +"%'");
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
            string id = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
            try
            {
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM MATERIAL_PRICE WHERE MATERIAL_TYPE='" + id + "'";
                basec.getcoms(strSql);
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
            string v1 = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
            int i=dataGridView1 .CurrentCell .RowIndex ;
            if (v1 != "")
            {
                textBox1.Text = bc.getOnlyString(string.Format("SELECT MRID FROM MATERIAL_PRICE WHERE MATERIAL_TYPE='{0}'",v1 ));
                textBox2.Text = dt.Rows[i]["类型"].ToString();
                textBox3.Text = dt.Rows[i]["起步价"].ToString();
                textBox4.Text = dt.Rows[i]["单位计价"].ToString();
                textBox6.Text = dt.Rows[i]["封顶金额"].ToString();
                comboBox1.Text = dt.Rows[i]["起步价单位"].ToString();
                comboBox2.Text = dt.Rows[i]["单位计价单位"].ToString();
                comboBox3.Text = dt.Rows[i]["封顶金额单位"].ToString();
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
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

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, this.Text );

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
