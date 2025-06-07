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
    public partial class PAPER_CORET : Form
    {
        DataTable dt = new DataTable();
        basec bc=new basec ();
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
        private static string _WAREID;
        public static string WAREID
        {
            set { _WAREID = value; }
            get { return _WAREID; }

        }
        private static string _CO_WAREID;
        public static string CO_WAREID
        {
            set { _CO_WAREID = value; }
            get { return _CO_WAREID; }

        }
        private static string _WNAME;
        public static string WNAME
        {
            set { _WNAME = value; }
            get { return _WNAME; }

        }
        private static string _STID;
        public static string STID
        {
            set { _STID = value; }
            get { return _STID; }

        }
        private static string _STEP_ID;
        public static string STEP_ID
        {
            set { _STEP_ID = value; }
            get { return _STEP_ID; }

        }
        private static string _STEP;
        public static string STEP
        {
            set { _STEP = value; }
            get { return _STEP; }

        }
        private  delegate bool dele(string a1,string a2);
        private delegate void delex();
        PAPER_CORE F1 = new PAPER_CORE();
        protected int M_int_judge, i;
        protected int select;
        CPAPER_CORE cPAPER_CORE = new CPAPER_CORE();
       
        public PAPER_CORET()
        {
            InitializeComponent();
        }
        public PAPER_CORET(PAPER_CORE FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        private void PAPER_CORET_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            textBox1.Text = IDO;
            bind();
        }
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cPAPER_CORE.GetTableInfo();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        private void dgvClientInfo_DoubleClick(object sender, EventArgs e)
        {
   
        }
      
        public void a1()
        {
            dataGridView1.ReadOnly = true;
            select = 0;
        }
        public void a2()
        {
            dataGridView1.ReadOnly = true;
            select = 1;
        }

 
        private void dgvClientInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
         
     
          
        }

        public void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.Text = "";
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            bind();
            try
            {
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #region bind
        private void bind()
        {

          
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ContextMenuStrip = contextMenuStrip1;
            textBox2.Focus();
            hint.Location = new Point(256, 136);
            hint.ForeColor = Color.Red;
            textBox2.BackColor = Color.Yellow;
          
         
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {

                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
            DataTable dtx = basec.getdts(cPAPER_CORE.sql + " where A.PCID='" +textBox1 .Text  + "' ORDER BY  B.PCID ASC ");
            if (dtx.Rows.Count > 0)
            {
               
                dt = cPAPER_CORE.GetTableInfo();
                textBox2.Text = dtx.Rows[0]["芯纸"].ToString();
                textBox3.Text = dtx.Rows[0]["税率"].ToString();
                comboBox1.Text = dtx.Rows[0]["客户类别"].ToString();
                foreach (DataRow dr1 in dtx.Rows)
                {
           
                    DataRow dr = dt.NewRow();
                    dr["项次"] = dr1["项次"].ToString();
                    dr["规格"] = dr1["规格"].ToString();
                    dr["含税单价"] = dr1["含税单价"].ToString();
                    dr["单位"] = dr1["单位"].ToString();
                    dr["芯纸门幅"] = dr1["芯纸门幅"].ToString();
                    dt.Rows.Add(dr);
                 
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
            this.Text = "复合用芯纸信息编辑";
        

        }
        #endregion
        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

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
                    if (IFExecution_SUCCESS == true && ADD_OR_UPDATE == "ADD")
                    {
                        add();
                    }

                    F1.load();
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
            textBox1 .Text = cPAPER_CORE.GETID();
        
            bind();
         
            ADD_OR_UPDATE = "ADD";
           

        }
        private void save()
        {

            btnSave.Focus();
            //dgvfoucs();
            if (dt.Rows.Count > 0)
            {
                DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "规格 IS NOT NULL");
                if (dtx.Rows.Count > 0)
                {

                    cPAPER_CORE.EMID = LOGIN.EMID;
                    cPAPER_CORE.PCID = textBox1.Text;
                    cPAPER_CORE.PAPER_CORE = textBox2.Text;
                    cPAPER_CORE.TAX_RATE = decimal.Parse(textBox3.Text);
                    cPAPER_CORE.CUSTOMER_TYPE = comboBox1.Text;
                    cPAPER_CORE.save(dtx);
                    IFExecution_SUCCESS = cPAPER_CORE.IFExecution_SUCCESS;
                    hint.Text = cPAPER_CORE.ErrowInfo;
                    if (IFExecution_SUCCESS)
                    {
                      
                        bind();
                    }
                    /*F1.Bind();
                    F1.search();*/

                }
                else
                {
                
                    hint.Text = "至少有一项规格才能保存！";

                }
            }
           
            try
            {
       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        private bool juage()
        {
            
            bool b = false;
            if (IDO == "")
            {
                hint.Text = "面纸规格编号不能为空！";
                b = true;
            }
            else if (textBox2.Text == "")
            {
                hint.Text = "芯纸不能为空！";
                b = true;
            }
            else if (textBox3.Text == "" || bc.yesno(textBox3.Text) == 0)
            {
                hint.Text = "税率不能为空且只能为数字！";
                b = true;
            }
            else if (comboBox1.Text == "")
            {
                hint.Text = "客户类别不能为空！";
                b = true;
            }
           else if(juage2())
           {
            
               b = true;
            }
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE PCID='{0}'",bc.RETURN_PCID(textBox2 .Text ))))
            {
                hint.Text = string.Format("面纸规格 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
            return b;
        }
        #region juage2()
  
        private bool juage2()
        {
            bool b = false;
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "规格 IS NOT NULL");
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    if (dr["含税单价"].ToString() != "" && bc.yesno(dr["含税单价"].ToString()) == 0)
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 的含税单价只能输入数字", dr["项次"].ToString());
                        break;
                    }
                    else if (dr["芯纸门幅"].ToString() != "" && bc.yesno(dr["芯纸门幅"].ToString()) == 0)
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 的含税单价只能输入数字", dr["项次"].ToString());
                        break;
                    }
                    {


                    }
                }
            }
            return b;
        }
        #endregion
     
        private void btnDel_Click(object sender, EventArgs e)
        {
            string id = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
           /* if (bc.exists(string.Format("SELECT * FROM WORKORDER_MST WHERE PCID='{0}'", bc.RETURN_PCID(textBox2.Text))))
            {
                hint.Text = string.Format("面纸规格 {0} 已经在工单中使用不允许删除", textBox2.Text);
             
            }
            else
            {
               
            }*/
            if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                basec.getcoms("DELETE PAPER_CORE_MST WHERE PCID='" + textBox1 .Text + "'");
                basec.getcoms("DELETE PAPER_CORE_DET WHERE PCID='" + textBox1 .Text  + "'");
               
                bind();
                ClearText();
                add();
                textBox1.Text = "";
                F1.load();

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
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
       
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            dataGridView1.Columns["项次"].Width = 40;
       
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


       
            dataGridView1.Columns["规格"].DefaultCellStyle.BackColor = Color.Yellow;
 
            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["含税单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["芯纸门幅"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["项次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        #endregion


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

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
            //dgvfoucs();

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
       
        }

        private void 删除此项ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            /*BASE_INFO.WAREINFO FRM = new CSPSS.BASE_INFO.WAREINFO();
            FRM.PAPER_CORE_USE();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = WAREID;
                textBox4.Text = CO_WAREID;
                textBox5.Text = WNAME;
            }
            textBox6.Focus();*/
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            int intC= this.dataGridView1.CurrentCell.RowIndex;


   if (select == 0)
   {
       /*BASE_INFO.STEP step = new STEP();
       step.PAPER_CORE_USE();
       step.ShowDialog();
       if (IF_DOUBLE_CLICK)
       {
         
           dt.Rows[intC]["规格代码"] = STEP_ID;
           dt.Rows[intC]["规格名称"] = STEP;
       }*/

   }
   if (select == 1)
   {

             

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
