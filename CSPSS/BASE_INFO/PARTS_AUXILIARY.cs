﻿using System;
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
    public partial class PARTS_AUXILIARY : Form
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
        CPARTS_AUXILIARY cPARTS_AUXILIARY = new CPARTS_AUXILIARY();

   
        protected int M_int_judge, i;
        protected int select;
        public PARTS_AUXILIARY()
        {
            InitializeComponent();
        }
    
        private void DEPAET_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            textBox1.Text = IDO;
            bind();

        }
        private void bind()
        {
            textBox4.Text = "17.0";
            dt = basec.getdts(cPARTS_AUXILIARY .sql );
            dt = cPARTS_AUXILIARY.RETURN_HAVE_ID_DT(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.AllowUserToAddRows = false;
            textBox2.Focus();
            textBox2.BackColor = Color.Yellow;
            textBox3.BackColor = Color.Yellow;
            comboBox1.BackColor = Color.Yellow;
            textBox4.BackColor = Color.Yellow;
            textBox3.TextAlign = HorizontalAlignment.Right;
            textBox4.TextAlign = HorizontalAlignment.Right;
            textBox4.TextAlign = HorizontalAlignment.Right;
            dgvStateControl();
            hint.Location = new Point(256, 136);
            hint.ForeColor = Color.Red;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            DataTable dtx = bc.getdt("SELECT * FROM UNIT");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("");
              
                foreach (DataRow dr in dtx.Rows)
                {
                    comboBox1.Items.Add(dr["UNIT"].ToString());
                }
            }
         
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
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
               
                dataGridView1.Columns["配件名"].Width = 120;
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
            dataGridView1.Columns["含税单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["未税单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

          
        }
        #endregion
    
        #region save
        private void save()
        {
            cPARTS_AUXILIARY.PAID = textBox1.Text;
            cPARTS_AUXILIARY.PARTS_AUXILIARY = textBox2.Text;
            cPARTS_AUXILIARY.TAX_UNIT_PRICE  = textBox3.Text;
            cPARTS_AUXILIARY.UNIT  = comboBox1.Text;
            cPARTS_AUXILIARY.TAX_MACHINE_COST  = textBox4.Text;
            cPARTS_AUXILIARY.TAX_RATE = textBox4.Text;
            cPARTS_AUXILIARY.REMARK  = textBox6.Text;
            cPARTS_AUXILIARY.MAKERID = LOGIN.EMID;
            cPARTS_AUXILIARY.save();
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
                hint.Text = "配件名不能为空！";
            }
            else if (textBox3.Text == "")
            {
                b = true;
                hint.Text = "含税单价不能为空";
            }
            else if (textBox4.Text == "")
            {
                b = true;
                hint.Text = "税率不能为空";
            }
            else if (textBox3.Text != "" && bc.yesno (textBox3 .Text )==0 ||  textBox4.Text != "" && bc.yesno (textBox4 .Text )==0 || 
                textBox4.Text != "" && bc.yesno (textBox4 .Text )==0)
            {
                b = true;
                hint.Text = "单价 税率只能输入数字";
            }
            else if (comboBox1 .Text =="" )
            {
                b = true;
                hint.Text = "单位不能为空";
            }
            return b;

        }
        #endregion
        public void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox4.Text = "17.0";
            comboBox1.Text = "";
            textBox6.Text = "";
        
        }
 
        private void add()
        {
            ClearText();
            textBox1.Text = cPARTS_AUXILIARY.GETID();
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
                    IFExecution_SUCCESS = cPARTS_AUXILIARY.IFExecution_SUCCESS;
                    hint.Text = cPARTS_AUXILIARY.ErrowInfo;
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


                dt = bc.getdt(cPARTS_AUXILIARY .sql +" WHERE  A.PARTS_AUXILIARY LIKE '%"+textBox5 .Text +"%'");
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
            string id = dt.Rows[dataGridView1.CurrentCell.RowIndex]["配件名"].ToString();
            try
            {
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM PARTS_AUXILIARY WHERE PARTS_AUXILIARY='" + id + "'";
                basec.getcoms(strSql);
                textBox1.Text = cPARTS_AUXILIARY.GETID();
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
            string v1 = dt.Rows[dataGridView1.CurrentCell.RowIndex]["配件名"].ToString();
            int i=dataGridView1 .CurrentCell .RowIndex ;
            if (v1 != "")
            {
                textBox1.Text = bc.getOnlyString(string.Format("SELECT PAID FROM PARTS_AUXILIARY WHERE PARTS_AUXILIARY='{0}'",v1 ));
                textBox2.Text = dt.Rows[i]["配件名"].ToString();
                textBox3.Text = dt.Rows[i]["含税单价"].ToString();
                comboBox1.Text = dt.Rows[i]["单位"].ToString();
                textBox4.Text = bc.RETURN_UNTIL_CHAR (dt.Rows[i]["税率"].ToString(), '%');
                textBox6.Text = dt.Rows[i]["说明"].ToString();
              
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

                bc.dgvtoExcel(dataGridView1, "配件辅材");

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
    }
}
