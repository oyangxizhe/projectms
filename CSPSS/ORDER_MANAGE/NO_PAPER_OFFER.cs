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

namespace CSPSS.ORDER_MANAGE
{
    public partial class NO_PAPER_OFFER : Form
    {
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        StringBuilder sqb = new StringBuilder();
        basec bc=new basec ();
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
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private bool _IF_CHECKBOX;
        public bool IF_CHECKBOX
        {
            set { _IF_CHECKBOX = value; }
            get { return _IF_CHECKBOX; }

        }
        private static bool _IF_DOUBLE_CLICK;
        public static bool IF_DOUBLE_CLICK
        {
            set { _IF_DOUBLE_CLICK = value; }
            get { return _IF_DOUBLE_CLICK; }

        }
        protected int M_int_judge, i;
        protected int select;
        CNO_PAPER_OFFER cNO_PAPER_OFFER = new CNO_PAPER_OFFER();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        public NO_PAPER_OFFER()
        {
            InitializeComponent();
        }
        private void NO_PAPER_OFFER_Load(object sender, EventArgs e)
        {  
            try
            {
            
              this.Icon = Resource1.xz_200X200;
                //textBox3.Text = "15BJH-Z001-01-M";
                hint.Location = new Point(400, 100);
                hint.ForeColor = Color.Red;
                dateTimePicker1.CustomFormat = "yyyy/MM/dd";
                dateTimePicker2.CustomFormat = "yyyy/MM/dd";
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker2.Format = DateTimePickerFormat.Custom;
                hint.Text = "";
                comboBox1.Focus();
                LOAD_OR_SEARCH = false;
                hint.Text = "";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

      
        }
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("纸品报价新增", LOGIN.USID);
            btnAdd.Visible = false;
            label17.Visible = false;
            checkBox1.Visible = false;
            label21.Visible = false;
            label6.Visible = false;
            dateTimePicker1.Visible = false;
            dateTimePicker2.Visible = false;
            if (dtx.Rows.Count > 0)
            {
                if (dtx.Rows[0]["新增权限"].ToString() == "有权限")
                {
                    btnAdd.Visible = true;
                    label17.Visible = true;
                }
                if (dtx.Rows[0]["报价日期查询"].ToString() == "有权限")
                {
                    checkBox1.Visible = true;
                    label21.Visible = true;
                    label6.Visible = true;
                    dateTimePicker1.Visible = true;
                    dateTimePicker2.Visible = true;
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
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #region bind
        public  void bind()
        {
            try
            {
                hint.Text = "";
                StringBuilder stb = new StringBuilder();
                stb.Append("  WHERE  B.PROJECT_NAME LIKE '%" + textBox1.Text + "%'");
                stb.Append(" AND D.OFFER_ID LIKE '%" + textBox2.Text + "%'");
                stb.Append(" AND A.PROJECT_ID LIKE '%" + comboBox1.Text + "%'");
                string v1 = dateTimePicker1.Text + " 0:00:00";
                string v2 = dateTimePicker2.Text + " 23:59:59";
                if (checkBox1.Checked)
                {
                    stb.Append(" AND A.DATE  BETWEEN  '" + v1 + "' AND '" + v2 + "'");
                    //MessageBox.Show(" AND B.DATE  '" + v1 + "' AND '" + v2 + "'");
                }
                dataGridView1.AllowUserToAddRows = false;
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
            hint.Text = "";
            string sqlo;
            string sqlx = "";
            if (LOAD_OR_SEARCH)
            {
                sqlo = " ORDER BY D.OFFER_ID ASC";
            }
            else
            {

                 sqlo = " ORDER BY D.OFFER_ID ASC";
            }
            //string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
            string v7 = "Y";/*本作业不按组权限，单个用户权限，所有用户权限限制 170303*/
         
            if (comboBox1.Text == "" && textBox1.Text == "" && textBox2.Text == "" && checkBox1 .Checked ==false )
            {
                //hint.Text = "未选择查询内容或是查询日期期间";
                dataGridView1.DataSource = null;
                return;

            }
            else  if (v7 == "Y")
            {
                sqlx = sql + sqlo;
                dt = bc.getdt(cNO_PAPER_OFFER .sql + sqlx);

        
            }
            else if (v7 == "GROUP")
            {


                sqlx = sql + @" AND A.MAKERID IN (SELECT EMID FROM USERINFO A WHERE UGID IN 
 (SELECT UGID FROM USERINFO WHERE USID='" + LOGIN.USID + "'))" + sqlo;
                dt = bc.getdt(cNO_PAPER_OFFER.sql +sqlx);
            }
            else
            {
             
                sqlx=sql + " AND A.MAKERID='" + LOGIN.EMID + "'" + sqlo;
                dt = bc.getdt(cNO_PAPER_OFFER.sql +sqlx);
            }
            if (dt.Rows.Count > 0)
            {
               
             dt = cNO_PAPER_OFFER.RETURN_HAVE_ID_DT (dt,LOGIN .EMID,LOGIN.POSITION  );
      
            }
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
            IDO = cNO_PAPER_OFFER.GETID();
            if (Screen.AllScreens[0].Bounds.Width == 1920)
            {
              
                NO_PAPER_OFFERT FRM = new NO_PAPER_OFFERT(this);
                FRM.IDO = cNO_PAPER_OFFER.GETID();
                FRM.ADD_OR_UPDATE = "ADD";
                FRM.Show();
            }
            else
            {
                NO_PAPER_OFFERT FRM = new NO_PAPER_OFFERT(this);
                FRM.IDO = cNO_PAPER_OFFER.GETID();
                FRM.ADD_OR_UPDATE = "ADD";
                FRM.Show();
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
           
            dataGridView1.ClearSelection();//取消默认选中行
            int i;
            //this.dataGridView1.MergeColumnNames.Add("序号");
            /*this.dataGridView1.MergeColumnNames.Add("项目名称");
            this.dataGridView1.MergeColumnNames.Add("数量");
            this.dataGridView1.MergeColumnNames.Add("项目号");
            this.dataGridView1.MergeColumnNames.Add("报价编号");
            this.dataGridView1.MergeColumnNames.Add("报价");
            this.dataGridView1.MergeColumnNames.Add("日期");*/
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView1.Columns["序号"].Width = 40;
            int numCols1 = dataGridView1.Columns.Count;
           
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

                if (
                    dataGridView1.Columns[i].DataPropertyName == "数量" || dataGridView1.Columns[i].DataPropertyName == "项目号" ||
                    dataGridView1.Columns[i].DataPropertyName == "单价" || dataGridView1.Columns[i].DataPropertyName == "序号")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else if (dataGridView1.Columns[i].DataPropertyName == "项目名称" || dataGridView1.Columns[i].DataPropertyName == "报价编号")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }
                dataGridView1.Columns[i].ReadOnly = true;
               


            }
         

                if (OFFFER_ID != null)
                {
                    dataGridView1.Columns["项目名称"].Visible = false;
                    dataGridView1.Columns["数量"].Visible = false;
                    dataGridView1.Columns["项目号"].Visible = false;
                    dataGridView1.Columns["报价编号"].Visible = false;
                    //dataGridView1.Columns["报价"].Visible = false;
                    //dataGridView1.Columns["日期"].Visible = false;
                  
                }
                /*dataGridView1.Columns["客户"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns["品牌"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns["AE"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns["审核批注"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns["项目号"].HeaderText = "项目编号";
                dataGridView1.Columns["AE"].Width = 50;
                dataGridView1.Columns["打样金额"].Width = 50;
                dataGridView1.Columns["报价数量"].Width = 40;
                dataGridView1.Columns["项目名称"].Width = 200;
                dataGridView1.Columns["报出价"].Width = 40;
                dataGridView1.Columns["审核批注"].Width = 200;
                dataGridView1.Columns["报价编号"].Width = 120;*/
            
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView1.Rows.Count-1; i++)
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



        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            try
            {
                sqb = new StringBuilder();
                sqb.AppendFormat(cNO_PAPER_OFFER.sql);
                sqb.AppendFormat(" WHERE DateDiff(day,A.DATE,getdate()) >-1 and DateDiff(day,A.DATE,getdate()) <+20");
                //string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
                string v7 = "Y";
                if (v7 == "Y")
                {
                    dtx = bc.getdt(sqb.ToString ());
                }
                else if (v7 == "GROUP")
                {
                    sqb.AppendFormat (@" AND A.MAKERID IN (SELECT EMID FROM USERINFO A WHERE UGID IN 
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
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                IF_CHECKBOX = true;
            }
            else
            {
                IF_CHECKBOX = false;
            }
        }
        private void btnToExcel_Click(object sender, EventArgs e)
        {
           
            if (dataGridView1.Rows.Count> 0)
            {
                bc.dgvtoExcel(dataGridView1,this.Text );
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
        
            try
            {
                dtx = bc.getdt(cNO_PAPER_OFFER.sql + string.Format(" WHERE A.PROJECT_ID='{0}'", comboBox1.Text));
                if (dtx.Rows.Count > 0)
                {

                    bind();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dtx = bc.getdt(cNO_PAPER_OFFER.sql + string.Format(" WHERE B.PROJECT_NAME='{0}'", textBox1.Text));
                if (dtx.Rows.Count > 0)
                {
                    bind();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
         
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dtx = bc.getdt(cNO_PAPER_OFFER.sql + string.Format(" WHERE A.OFFER_ID='{0}'", textBox2.Text));
                if (dtx.Rows.Count > 0)
                {
                    bind();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                if (select == 1)
                {
                }
                this.Close();


            }
            else
            {
                string v1 = dt.Rows[dataGridView1.CurrentCell.RowIndex]["报价编号"].ToString();
                if (Screen.AllScreens[0].Bounds.Width == 1920)
                {
                    NO_PAPER_OFFERT FRM = new NO_PAPER_OFFERT(this);
                    sqb = new StringBuilder();
                    sqb.AppendFormat("");
                    FRM.IDO = bc.getOnlyString(string.Format("SELECT NPID FROM NO_PAPER_OFFER_DET WHERE OFFER_ID='{0}'", v1));
                    FRM.ADD_OR_UPDATE = "UPDATE";
                    // MessageBox.Show(bc.getOnlyString(string.Format("SELECT PFID FROM NO_PAPER_OFFER_MST WHERE OFFER_ID='{0}'", v1)));
                    FRM.Show();

                }
                else
                {
                    NO_PAPER_OFFERT FRM = new NO_PAPER_OFFERT(this);

                    FRM.IDO = bc.getOnlyString(string.Format("SELECT NPID FROM NO_PAPER_OFFER_DET WHERE OFFER_ID='{0}'", v1));
                    FRM.ADD_OR_UPDATE = "UPDATE";
                    // MessageBox.Show(bc.getOnlyString(string.Format("SELECT PFID FROM NO_PAPER_OFFER_MST WHERE OFFER_ID='{0}'", v1)));
                    FRM.Show();
                }
            }
            try
            {
      
            }
            catch (Exception)
            {

            }
        }
    }
}
