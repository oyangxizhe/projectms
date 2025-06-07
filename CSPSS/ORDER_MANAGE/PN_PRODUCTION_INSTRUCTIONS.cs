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
    public partial class PN_PRODUCTION_INSTRUCTIONS : Form
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
        private string _OFFER_ID;
        public string OFFER_ID
        {
            set { _OFFER_ID = value; }
            get { return _OFFER_ID; }
        }
        private  string _ORDER_ID;
        public  string ORDER_ID
        {
            set { _ORDER_ID = value; }
            get { return _ORDER_ID; }
        }
        private string _CNAME;
        public string CNAME
        {
            set { _CNAME = value; }
            get { return _CNAME; }
        }
        private string _WNAME;
        public string WNAME
        {
            set { _WNAME = value; }
            get { return _WNAME; }
        }
        private string _PROCESS_COUNT;
        public string PROCESS_COUNT
        {
            set { _PROCESS_COUNT = value; }
            get { return _PROCESS_COUNT; }
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
        private bool _IF_COMPLETED;
        public bool IF_COMPLETED
        {
            set { _IF_COMPLETED = value; }
            get { return _IF_COMPLETED; }
        }
        CPN_PRODUCTION_INSTRUCTIONS cPN_PRODUCTION_INSTRUCTIONS = new CPN_PRODUCTION_INSTRUCTIONS();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        CNO_PAPER_OFFER cno_paper_offer = new CNO_PAPER_OFFER();
        DataGridViewTextBoxColumn d1 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d2 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d4 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d5 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d6 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d7 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d8 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d9 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d10 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d11 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d12 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d13 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d14 = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn d15 = new DataGridViewTextBoxColumn();
        public PN_PRODUCTION_INSTRUCTIONS()
        {
            InitializeComponent();
        }
        private void PN_PRODUCTION_INSTRUCTIONS_Load(object sender, EventArgs e)
        {
           
            IF_COMPLETED = true;
            Control.CheckForIllegalCrossThreadCalls = false;//避免出现线程间操作无效: 从不是创建控件“progressBar1”的线程访问它 160120
            d1.Name = "序号";
            d1.HeaderText = "序号";
            dataGridView1.Columns.Add(d1);
        
            d2.Name = "订单编号";
            d2.HeaderText = "订单编号";
            dataGridView1.Columns.Add(d2);
      
            d4.Name = "品号";
            d4.HeaderText = "品号";
            dataGridView1.Columns.Add(d4);
         
            d5.Name = "客户名称";
            d5.HeaderText = "客户名称";
            dataGridView1.Columns.Add(d5);
            
            d6.Name = "AE";
            d6.HeaderText = "AE";
            dataGridView1.Columns.Add(d6);
           
            d7.Name = "结构设计";
            d7.HeaderText = "结构设计";
            dataGridView1.Columns.Add(d7);
          
            d8.Name = "平面设计";
            d8.HeaderText = "平面设计";
            dataGridView1.Columns.Add(d8);
          
            d9.Name = "生产数量";
            d9.HeaderText = "生产数量";
            dataGridView1.Columns.Add(d9);
       
            d10.Name = "下单日期";
            d10.HeaderText = "下单日期";
            dataGridView1.Columns.Add(d10);
     
            d11.Name = "交货日期";
            d11.HeaderText = "交货日期";
            dataGridView1.Columns.Add(d11);
       
            d12.Name = "报价";
            d12.HeaderText = "报价";
            dataGridView1.Columns.Add(d12);
        
            d13.Name = "会签";
            d13.HeaderText = "会签";
            dataGridView1.Columns.Add(d13);
       
            d14.Name = "制单人";
            d14.HeaderText = "制单人";
            dataGridView1.Columns.Add(d14);

            d15.Name = "编号";
            d15.HeaderText = "编号";
            dataGridView1.Columns.Add(d15);
            d1.Visible = false;
            d2.Visible = false;
            d4.Visible = false;
            d5.Visible = false;
            d6.Visible = false;
            d7.Visible = false;
            d8.Visible = false;
            d9.Visible = false;
            d10.Visible = false;
            d11.Visible = false;
            d12.Visible = false;
            d13.Visible = false;
            d14.Visible = false;
            d15.Visible = false;
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
           try
           {
           
           }
           catch (Exception)
           {
               //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
           }
         
        }
     
     
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("生产指示书", LOGIN.USID);
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
            LOAD_OR_SEARCH = false;
            IFExecution_SUCCESS = false;
            hint.Text = "";
            if (backgroundWorker1.IsBusy)
            {
                hint.Text = "同一时刻只能执行一个任务";
            }
            else
            {
                progressBar1.Value = 0;//初始化进度条 16/01/20
                backgroundWorker1.RunWorkerAsync();//线程开始开始运行 16/01/20
                backgroundWorker1.WorkerReportsProgress = true;//允许使用线程进度  16/01/20
                backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);//线程开始后激发该事件,在此事件里处理进度条显示效果 16/01/20
                bind();
            }
        
            try
            {
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
         
                progressBar1.Maximum = 600;
                for (int i = 0; i <= 600; i++)
                {
                    if (IFExecution_SUCCESS)
                    {
                        progressBar1.Value = progressBar1.Maximum;
                        break;
                    }
                    else
                    {
                        progressBar1.Value = i;
                        System.Threading.Thread.Sleep(100);//线程开始后激发该事件,在此事件里处理进度条显示效果 16/01/20
                    }

                }
        }
        #region bind
        public  void bind()
        {
            dataGridView1.Rows.Clear();
            hint.Text = "";
            StringBuilder stb = new StringBuilder();
            stb.Append(cPN_PRODUCTION_INSTRUCTIONS.sql);
            stb.Append(" WHERE A.ORDER_ID LIKE '%" + comboBox1.Text + "%'");
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
                sqlo = " ORDER BY A.ORDER_ID ASC";
            }
            else
            {
                 sqlo = " ORDER BY A.ORDER_ID ASC";
            }
            //string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
            string v7 = "Y";//本做业不受组权限限制 170306
            if ( comboBox1 .Text  == "" && checkBox1.Checked == false)
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
            if (dt.Rows.Count > 0)
            {
                int j = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    d1.Visible = true;
                    d2.Visible = true;
                     CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
                    //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
                     if (LOGIN.POSITION == "总经理" || LOGIN.POSITION == "财务" || LOGIN .UNAME =="admin")
                     {
                         d4.Visible = true;
                         d5.Visible = true;
                         d6.Visible = true;
                         d7.Visible = true;
                         d8.Visible = true;
                         d9.Visible = true;
                         d10.Visible = true;
                         d11.Visible = true;
                         d12.Visible = true;
                         d13.Visible = true;
                         d14.Visible = true;
                     }
                     if (!bc.exists(ccustomer_info.sqlsi + " WHERE B.CNAME='" + dt.Rows[i]["客户名称"].ToString() +
                           "' AND A.USER_MAKERID='" +LOGIN .EMID + "' ") && LOGIN.POSITION =="AE")
                    {
                      
                    }
                    else
                    {
                        //MessageBox.Show(dt.Rows[i]["客户名称"].ToString() + "," + LOGIN.EMID + "," + LOGIN.POSITION);
                        DataGridViewRow dar = new DataGridViewRow();
                        dataGridView1.Rows.Add(dar);
                        dataGridView1["序号", j].Value = (j + 1).ToString();
                        dataGridView1["订单编号", j].Value = dt.Rows[i]["订单编号"].ToString();
                        dataGridView1["品号", j].Value = dt.Rows[i]["品号"].ToString();
                        dataGridView1["客户名称", j].Value = dt.Rows[i]["客户名称"].ToString();
                        dataGridView1["AE", j].Value = dt.Rows[i]["AE01"].ToString();
                        dataGridView1["结构设计", j].Value = dt.Rows[i]["结构01"].ToString();
                        dataGridView1["平面设计", j].Value = dt.Rows[i]["平面01"].ToString();
                        dataGridView1["生产数量", j].Value = dt.Rows[i]["生产数量"].ToString();
                        dataGridView1["下单日期", j].Value = dt.Rows[i]["下单日期"].ToString();
                        dataGridView1["交货日期", j].Value = dt.Rows[i]["交货日期"].ToString();
                        dataGridView1["报价", j].Value = dt.Rows[i]["报价"].ToString();
                        if (JUAGE_IF_AUDIT_END(dt.Rows[i]["编号"].ToString()))
                        {
                            dataGridView1["会签", j].Value = "已会签";
                        }
                        else
                        {
                            dataGridView1["会签", j].Value = "待会签";
                        }
                        dataGridView1["制单人", j].Value = dt.Rows[i]["制单人"].ToString();
                        j = j + 1;
                    }
              
                }
                dgvStateControl();
                IFExecution_SUCCESS = true;
            }
            else
            {
                hint.Text = "找不到所要搜索项！";
                IFExecution_SUCCESS = true;
                dataGridView1.DataSource = null;

            }
        }
             #region JUAGE_IF_AUDIT_END 
        public bool JUAGE_IF_AUDIT_END(string PNID)
        {
            bool b = true;
            DataTable dt = new DataTable();
            dt = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS .sql  + string.Format(" WHERE A.PNID='{0}'", PNID));
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["是否需纸品生产签核"].ToString() == "是" && dt.Rows[0]["纸品生产签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需木铁生产签核"].ToString() == "是" && dt.Rows[0]["木铁生产签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需亚克力生产签核"].ToString() == "是" && dt.Rows[0]["亚克力生产签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需纸品计划签核"].ToString() == "是" && dt.Rows[0]["纸品计划签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需木铁计划签核"].ToString() == "是" && dt.Rows[0]["木铁计划签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需平面设计签核"].ToString() == "是" && dt.Rows[0]["平面设计签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需结构设计签核"].ToString() == "是" && dt.Rows[0]["结构设计签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需纸品采购签核"].ToString() == "是" && dt.Rows[0]["纸品采购签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
                else if (dt.Rows[0]["是否需木铁采购签核"].ToString() == "是" && dt.Rows[0]["木铁采购签核状态"].ToString() == "未签核")
                {
                    b = false;
                }
            }
            else
            {
                b = false;
            }
            return b;

        }
        #endregion
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.Columns["序号"].Width = 40;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
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
        private void btnAdd_Click(object sender, EventArgs e)
        {

            if (Screen.AllScreens[0].Bounds.Width == 1920)
            {
                ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST FRM = new CSPSS.ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST();
                FRM.ADD_OR_UPDATE = "ADD";
                FRM.IDO = cPN_PRODUCTION_INSTRUCTIONS.GETID();
                FRM.Show();
            }
            else
            {
                ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST FRM = new CSPSS.ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST();
                FRM.ADD_OR_UPDATE = "ADD";
                FRM.IDO = cPN_PRODUCTION_INSTRUCTIONS.GETID();
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
  
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            bind();
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

            }
            else
            {
                
                if (Screen.AllScreens[0].Bounds.Width == 1920)
                {
                    PN_PRODUCTION_INSTRUCTIONST FRM = new PN_PRODUCTION_INSTRUCTIONST(this);
                    string v1 = dataGridView1 ["订单编号",dataGridView1.CurrentCell.RowIndex].Value .ToString();

                    FRM.IDO = bc.getOnlyString(string.Format("SELECT PNID FROM PN_PRODUCTION_INSTRUCTIONS WHERE ORDER_ID='{0}'", v1));
                    FRM.ADD_OR_UPDATE = "UPDATE";
                    FRM.Show();
                }
                else
                {
                    ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST FRM = new CSPSS.ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST(this);
                    string v1 = dataGridView1["订单编号", dataGridView1.CurrentCell.RowIndex].Value .ToString();
                    FRM.IDO = bc.getOnlyString(string.Format("SELECT PNID FROM PN_PRODUCTION_INSTRUCTIONS WHERE ORDER_ID='{0}'", v1));
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

                        cPN_PRODUCTION_INSTRUCTIONS.ExcelPrint(dt, "xxx样板依赖单", System.IO.Path.GetFullPath("xxx样板依赖单.xls"));
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
                sqb.AppendFormat(cPN_PRODUCTION_INSTRUCTIONS.sql);
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
                dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "订单编号");
                comboBox1.Items.Clear();
                if (dtx.Rows.Count > 0)
                {
                    foreach (DataRow dr in dtx.Rows)
                    {
                        comboBox1.Items.Add(dr["VALUE"].ToString()); ;
                    }
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
        public void INVENTORY_USE()
        {
            select = 1;
        }
        public void INVOICE_USE()
        {
            select = 2;
        }
        public void RECEIVABLE_USE()
        {
            select = 3;
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
         

            try
            {
                sqb = new StringBuilder();
                sqb.AppendFormat(cPN_PRODUCTION_INSTRUCTIONS.sql);
                sqb.AppendFormat(" WHERE A.ORDER_ID='{0}'", comboBox1.Text);
                dtx = bc.getdt(sqb.ToString());
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

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            if (select != 0)
            {
                int intCurrentRowNumber = this.dataGridView1.CurrentCell.RowIndex;
                string s1 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[0].Value.ToString().Trim();
                string s2 = this.dataGridView1.Rows[intCurrentRowNumber].Cells["订单编号"].Value.ToString().Trim();
                string s3 = this.dataGridView1.Rows[intCurrentRowNumber].Cells["客户名称"].Value.ToString().Trim();
                string s4 = this.dataGridView1.Rows[intCurrentRowNumber].Cells["生产数量"].Value.ToString().Trim();
                string s5 = this.dataGridView1.Rows[intCurrentRowNumber].Cells["品号"].Value.ToString().Trim();
                if (select == 1)
                {
                    ORDER_ID = s2;
                    CNAME = s3;
                    PROCESS_COUNT = s4;
                    WNAME = s5;
                    INVENTORYT.IF_DOUBLE_CLICK = true;
                }
                if (select == 2)
                {
                    ORDER_ID = s2;
                    CNAME = s3;
                    PROCESS_COUNT = s4;
                    WNAME = s5;
                    INVOICET.IF_DOUBLE_CLICK = true;
                }
                if (select == 3)
                {
                    ORDER_ID = s2;
                    CNAME = s3;
                    PROCESS_COUNT = s4;
                    WNAME = s5;
                    RECEIVABLET.IF_DOUBLE_CLICK = true;
                }
                this.Close();
            }
        }


      
    }
}
