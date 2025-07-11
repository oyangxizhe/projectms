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
    public partial class NOTICE_LIST : Form
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
        CNOTICE_LIST cnotice_list = new CNOTICE_LIST();
   
        protected int M_int_judge, i;
        protected int select;
        public NOTICE_LIST()
        {
            InitializeComponent();
        }
        #region double_click
        private void dgvEmployeeInfo_DoubleClick(object sender, EventArgs e)
        {
            
        }
        #endregion

        private void DEPAET_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            bind();

        }

        private void bind()
        {
           
            textBox1.Text = IDO;
            dt = basec.getdts(cnotice_list .sql );
            dataGridView1.DataSource = dt;
  
            dgvStateControl();
            hint.Location = new Point(256, 136);
            hint.ForeColor = Color.Red;
            dataGridView1.AllowUserToAddRows = false;
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

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            
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
        }
        #endregion
    
        #region save
        private void save()
        {

            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss");
            string EMID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='"+comboBox1.Text  +"'");
            string varMakerID = LOGIN.EMID;
            if (!bc.exists("SELECT NLID FROM NOTICE_LIST WHERE NLID='" + textBox1 .Text + "'"))
            {

                if (bc.exists("SELECT * FROM NOTICE_LIST WHERE EMID='"+EMID   +"'"))
                {
                    hint.Text = "员工工号已经存在";
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    basec.getcoms(@"INSERT INTO NOTICE_LIST(NLID,EMID,MAKERID,DATE,YEAR,
                                   MONTH) VALUES ('" + textBox1.Text + "','" +EMID  +
                     "','" + varMakerID + "','" + varDate +
                     "','" + year + "','" + month + "')");
                    IFExecution_SUCCESS = true;
                    bind();
                }

            }
            else
            {
                if (bc.exists("SELECT * FROM NOTICE_LIST WHERE EMID='" + EMID  + "'"))
                {
                    hint.Text = "员工工号已经存在";
                    IFExecution_SUCCESS = false;
                }
                else
                {
                    basec.getcoms(@"UPDATE NOTICE_LIST SET EMID='" + EMID   + "',MAKERID='" + varMakerID +
                          "',DATE='" + varDate + "' WHERE NLID='" + textBox1.Text + "'");
                    IFExecution_SUCCESS = true;
                    bind();
                }
            }
           
        }
        #endregion
        #region juage()
        private bool juage()
        {

            bool b = false;
            if (comboBox1 .Text == "")
            {
                b = true;
                hint.Text = "员工工号不能为空！";
            }
            else if (!bc.exists(string.Format ("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",comboBox1 .Text )))
            {
                b = true;
                hint.Text = "员工工号不存在系统！";
            }
            return b;

        }
        #endregion
        public void ClearText()
        {
            comboBox1.Text = "";
            textBox2.Text = "";
        
        }

        private void dgvEmployeeInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string v1 = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
            if (v1 != "")
            {
                textBox1.Text = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                comboBox1 .Text = Convert.ToString(dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                textBox2.Text  = Convert.ToString(dataGridView1[2, dataGridView1.CurrentCell.RowIndex].Value).Trim();
             
            }

        }
        private void add()
        {
            ClearText();
            textBox1.Text = cnotice_list.GETID();
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


                dt = bc.getdt(cnotice_list .sql +" WHERE A.NLID LIKE '%"+textBox4.Text +"%' AND B.EMPLOYEE_ID LIKE '%"+textBox5 .Text +"%'");
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
                string strSql = "DELETE FROM NOTICE_LIST WHERE NLID='" + id + "'";
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
                if (dataGridView1.CurrentCell.ColumnIndex == 9)
                {
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else
                {
                    SendKeys.SendWait("{Tab}");
                }
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

        private void NOTICE_LIST_Load(object sender, EventArgs e)
        {

        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            EMPLOYEE_INFO FRM = new EMPLOYEE_INFO();
            FRM.NOTICE_LIST_USE();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = EMPLOYEE_ID;
                textBox2.Text = ENAME;
            }
        }
  
    }
}
