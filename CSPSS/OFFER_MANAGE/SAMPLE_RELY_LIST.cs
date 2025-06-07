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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSPSS.OFFER_MANAGE
{
    public partial class SAMPLE_RELY_LIST : Form
    {
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        StringBuilder sqb = new StringBuilder();
        basec bc=new basec ();
        #region nature
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private static string _EMID;
        public static string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private static string _ENAME;
        public static string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        private int _GET_DATA_INT;
        public int GET_DATA_INT
        {
            set { _GET_DATA_INT = value; }
            get { return _GET_DATA_INT; }

        }
        private bool _LOAD_OR_SEARCH;
        public bool LOAD_OR_SEARCH
        {
            set { _LOAD_OR_SEARCH = value; }
            get { return _LOAD_OR_SEARCH; }

        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private string _OFFFER_ID;
        public string OFFFER_ID
        {
            set { _OFFFER_ID = value; }
            get { return _OFFFER_ID; }
        }
        private string _SRID;
        public string SRID
        {
            set { _SRID = value; }
            get { return _SRID; }
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
        #endregion
        protected int M_int_judge, i;
        protected int select;
        CSAMPLE_RELY_LIST cSAMPLE_RELY_LIST = new CSAMPLE_RELY_LIST();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        public SAMPLE_RELY_LIST()
        {
            InitializeComponent();
        }
        private void SAMPLE_RELY_LIST_Load(object sender, EventArgs e)
        {
        
           try
           {
               right();
             this.Icon = Resource1.xz_200X200;
               hint.Location = new Point(400, 100);
               hint.ForeColor = Color.Red;
               dateTimePicker1.CustomFormat = "yyyy/MM/dd";
               dateTimePicker2.CustomFormat = "yyyy/MM/dd";
               dateTimePicker1.Format = DateTimePickerFormat.Custom;
               dateTimePicker2.Format = DateTimePickerFormat.Custom;
               hint.Text = "";
               LOAD_OR_SEARCH = false;
               hint.Text = "";
           }
           catch (Exception)
           {
               //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
           }
         
        }
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("打样单新增", LOGIN.USID);
            btnAdd.Visible = false;
            label17.Visible = false;
            if (dtx.Rows.Count > 0)
            {
                if (dtx.Rows[0]["新增权限"].ToString() == "有权限")
                {
                    btnAdd.Visible = true;
                    label17.Visible = true;
                }
            }

        }
        #endregion
        private void btnSearch_Click(object sender, EventArgs e)
        {
           
            try
            {
                LOAD_OR_SEARCH = false;
                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
            }
        }
        #region bind
        public  void bind()
        {
            hint.Text = "";
            if (bc.getOnlyString("SELECT UNAME FROM USERINFO WHERE USID='" + LOGIN.USID + "'") == "admin")
            {
                //btnToExcel.Visible = true;
            }
            else
            {
                //btnToExcel.Visible = true;
            }
            StringBuilder stb = new StringBuilder();
            stb.Append(cSAMPLE_RELY_LIST.sql);
            stb.Append("  WHERE  C.PROJECT_NAME LIKE '%" + textBox1.Text + "%'");
            stb.Append(" AND A.SAMPLE_ID LIKE '%" + comboBox2.Text + "%'");
            stb.Append(" AND C.PROJECT_ID LIKE '%" + comboBox1.Text + "%'");
            string v1 = dateTimePicker1.Text + " 0:00:00";
            string v2 = dateTimePicker2.Text + " 23:59:59";
            if (checkBox1.Checked)
            {
                stb.Append(" AND A.DATE  BETWEEN  '" + v1 + "' AND '" + v2 + "'");
                //MessageBox.Show(" AND B.DATE  '" + v1 + "' AND '" + v2 + "'");
            }

            dataGridView1.AllowUserToAddRows = false;
            //dataGridView1.ContextMenuStrip = contextMenuStrip1;

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
            search_o(stb.ToString());
            try
            {
    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
           
        }
        #endregion
 
        #region search_o()
        public void search_o(string sql)
        {
            string sqlo;
            if (LOAD_OR_SEARCH)
            {
                sqlo = " ORDER BY A.SAMPLE_ID ASC";
            }
            else
            {
                 sqlo = " ORDER BY A.SAMPLE_ID ASC";
            }
            string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
            //string v7 = "Y";
            if (comboBox1.Text == "" && textBox1.Text == "" && comboBox2 .Text  == "" && checkBox1.Checked == false)
            {
                //hint.Text = "未选择查询内容或是查询日期期间";
                dataGridView1.DataSource = null;
                return;
            }
            else  if (v7 == "Y")
            {
               
                dt = bc.getdt(sql + sqlo);
             
            }
            else if (v7 == "GROUP")
            {
                dt = bc.getdt(sql + @" AND A.MAKERID IN (SELECT EMID FROM USERINFO A WHERE UGID IN 
 (SELECT UGID FROM USERINFO WHERE USID='" + LOGIN.USID + "'))" + sqlo);
            }
            else
            {
                dt = bc.getdt(sql + " AND A.MAKERID='" + LOGIN.EMID + "'" + sqlo);

            }
           dt = cSAMPLE_RELY_LIST.RETURN_DT_SEARCH(dt);
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
                dgvStateControl();
            }
            else
            {
                hint.Text = "找不到所要搜索项！";
                dataGridView1.DataSource = null;

            }
        }
        #endregion
      
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (Screen.AllScreens[0].Bounds.Width == 1920)
                {
                    OFFER_MANAGE.SAMPLE_RELY_LISTT FRM = new CSPSS.OFFER_MANAGE.SAMPLE_RELY_LISTT();
                    FRM.ADD_OR_UPDATE = "ADD";
                    FRM.IDO = cSAMPLE_RELY_LIST.GETID();
                    FRM.Show();
                }
                else
                {
                    OFFER_MANAGE.SAMPLE_RELY_LISTT FRM = new CSPSS.OFFER_MANAGE.SAMPLE_RELY_LISTT();
                    FRM.ADD_OR_UPDATE = "ADD";
                    FRM.IDO = cSAMPLE_RELY_LIST.GETID();
                    FRM.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
            }
          
        }
        public void load()
        {
            bind();
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btndgvInfoCopy_Click(object sender, EventArgs e)
        {

            dgvCopy(ref dataGridView1);
        }
        private void dgvCopy(ref dgvInfo dgv)
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
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            for (i = 0; i < numCols1; i++)
            {
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].ReadOnly = true;
            }
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView1.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
        }
        #endregion
        public void WORKORDER_USE()
        {
            select = 1;
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            bind();
            /*if (checkBox2.Checked)
            {
                pictureBox1.Visible = false;
                label7.Visible = false;
            }
            else
            {
                pictureBox1.Visible = true;
                label7.Visible = true;
            }*/
        }
        #region dataGridView1_DataSourceChanged
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {

                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    if (dataGridView1.Columns[i].DataPropertyName == "部品总数")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "表面加工小计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "裱工小计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "正面防晒合计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "CTP单张价")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "面纸内耗")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00000";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "部品数")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.000";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "表面处理用量")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.000";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "裱工用量")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.000";
                    }
                 
                    else if (dataGridView1.Columns[i].DataPropertyName == "面纸小计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "芯纸用量")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "芯纸单个用量")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.000";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "芯纸小计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "面纸单个用量")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.000";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "底纸内耗")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "底纸下单")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "底纸单个用量")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.000";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "底纸小计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "部品总价")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "正面印工合计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else if (dataGridView1.Columns[i].DataPropertyName == "正反印工合计")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    }
                    else
                    {

                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00";
                    }
                  
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
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

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

            if (select != 0)
            {
                int intCurrentRowNumber = this.dataGridView1.CurrentCell.RowIndex;
                string s1 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[0].Value.ToString().Trim();
                string s2 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[1].Value.ToString().Trim();
                string s3 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[2].Value.ToString().Trim();
                string s4 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[3].Value.ToString().Trim();
                this.Close();
            }
            else
            {
                if (Screen.AllScreens[0].Bounds.Width == 1920)
                {
                    SAMPLE_RELY_LISTT FRM = new SAMPLE_RELY_LISTT(this);
                    string v1 = dt.Rows[dataGridView1.CurrentCell.RowIndex]["打样单号"].ToString();
                    FRM.IDO = bc.getOnlyString(string.Format("SELECT SRID FROM SAMPLE_RELY_LIST WHERE SAMPLE_ID='{0}'", v1));
                    FRM.ADD_OR_UPDATE = "UPDATE";
                    FRM.Show();
                }
                else
                {
                    OFFER_MANAGE.SAMPLE_RELY_LISTT FRM = new CSPSS.OFFER_MANAGE.SAMPLE_RELY_LISTT(this);
                    string v1 = dt.Rows[dataGridView1.CurrentCell.RowIndex]["打样单号"].ToString();
                    FRM.IDO = bc.getOnlyString(string.Format("SELECT SRID FROM SAMPLE_RELY_LIST WHERE SAMPLE_ID='{0}'", v1));
                    FRM.ADD_OR_UPDATE = "UPDATE";
                    FRM.Show();
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {

                    if (bc.RETURN_NOHAVE_REPEAT_DT(dt, "打样单号").Rows.Count > 1)
                    {
                        hint.Text = "每次只能导出一个项目";
                    }
                    else
                    {

                        cSAMPLE_RELY_LIST.ExcelPrint(dt, "xxx样板依赖单", System.IO.Path.GetFullPath("xxx样板依赖单.xls"));
                    }
                }
                else
                {
                    hint.Text = "没有内容可导出";
                }

            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, "机加工信息");

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
        
            bind();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            bind();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            bind();
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {

            try
            {
                sqb = new StringBuilder();
                sqb.AppendFormat(cSAMPLE_RELY_LIST.sql);
                sqb.AppendFormat(" WHERE DateDiff(day,A.DATE,getdate()) >-1 and DateDiff(day,A.DATE,getdate()) <+20");
                string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
                if (v7 == "Y")
                {
                    dtx = bc.getdt(sqb.ToString());
                }
                else if (v7 == "GROUP")
                {
                    sqb.AppendFormat(@" AND A.MAKERID IN (SELECT EMID FROM USERINFO A WHERE UGID IN 
 (SELECT UGID FROM USERINFO WHERE USID='" + LOGIN.USID + "'))");
                    dtx = bc.getdt(sqb.ToString());
                }
                else
                {
                    sqb.AppendFormat(" AND A.MAKERID='" + LOGIN.EMID + "'");
                    dtx = bc.getdt(sqb.ToString());
                }
                dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "项目号");
                if (dtx.Rows.Count > 0)
                {
                    comboBox1.Items.Clear();
                    //comboBox1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        comboBox1.Items.Add(dr["VALUE"].ToString());
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            sqb = new StringBuilder();
            sqb.AppendFormat(cSAMPLE_RELY_LIST.sql);
            sqb.AppendFormat(" WHERE DateDiff(day,A.DATE,getdate()) >-1 and DateDiff(day,A.DATE,getdate()) <+20");
            string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
            if (v7 == "Y")
            {
                dtx = bc.getdt(sqb.ToString());
            }
            else if (v7 == "GROUP")
            {
                sqb.AppendFormat(@" AND B.MAKERID IN (SELECT EMID FROM USERINFO A WHERE UGID IN 
 (SELECT UGID FROM USERINFO WHERE USID='" + LOGIN.USID + "'))");
                dtx = bc.getdt(sqb.ToString());
            }
            else
            {
                sqb.AppendFormat(" AND B.MAKERID='" + LOGIN.EMID + "'");
                dtx = bc.getdt(sqb.ToString());
            }
            dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "打样单号");
            if (dtx.Rows.Count > 0)
            {
                comboBox2.Items.Clear();
                foreach (DataRow dr in dtx.Rows)
                {
                    comboBox2.Items.Add(dr["VALUE"].ToString());
                }

            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
           
            try
            {
                dtx = bc.getdt(cprinting_offer.sqlfi + string.Format(" WHERE B.PROJECT_ID='{0}'", comboBox1.Text));
                if (dtx.Rows.Count > 0)
                {
                    bind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
         
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dtx = bc.getdt(cprinting_offer.sqlfi + string.Format(" WHERE B.PROJECT_NAME='{0}'",textBox1 .Text ));
                if (dtx.Rows.Count > 0)
                {
                    bind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dtx = bc.getdt(cprinting_offer.sqlfi + string.Format(" WHERE A.SAMPLE_ID='{0}'",comboBox2 .Text ));
                if (dtx.Rows.Count > 0)
                {
                    bind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                bc.dgvtoExcel(dataGridView1, this.Text);
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

  
      
    }
}
