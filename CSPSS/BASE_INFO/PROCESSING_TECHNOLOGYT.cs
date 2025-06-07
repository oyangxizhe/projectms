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
    public partial class PROCESSING_TECHNOLOGYT : Form
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
        private static string _PROCESSING_TECHNOLOGY;
        public static string PROCESSING_TECHNOLOGY
        {
            set { _PROCESSING_TECHNOLOGY = value; }
            get { return _PROCESSING_TECHNOLOGY; }

        }
        private static string _LAMINATING_PROCESS;
        public static string LAMINATING_PROCESS
        {
            set { _LAMINATING_PROCESS = value; }
            get { return _LAMINATING_PROCESS; }

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
        PROCESSING_TECHNOLOGY F1 = new PROCESSING_TECHNOLOGY();
        protected int M_int_judge, i;
        protected int select;
        CPROCESSING_TECHNOLOGY cPROCESSING_TECHNOLOGY = new CPROCESSING_TECHNOLOGY();
       
        public PROCESSING_TECHNOLOGYT()
        {
            InitializeComponent();
        }
        public PROCESSING_TECHNOLOGYT(PROCESSING_TECHNOLOGY FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        private void PROCESSING_TECHNOLOGYT_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            textBox1.Text = IDO;
            bind();
        }
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cPROCESSING_TECHNOLOGY.GetTableInfo();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
      
      
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

        public void ClearText()
        {
            textBox2.Text = "";
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
            textBox2.BackColor = Color.Yellow;
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
          
            DataTable dtx = basec.getdts(cPROCESSING_TECHNOLOGY.sql +" WHERE A.PTID='"+textBox1 .Text +"'");
            if (dtx.Rows.Count > 0)
            {
                textBox2.Text = dtx.Rows[0]["加工内容"].ToString();
                dt = cPROCESSING_TECHNOLOGY.GetTableInfo();
                foreach (DataRow dr1 in dtx.Rows)
                {
                    DataRow dr = dt.NewRow();
                    dr["项次"] = dr1["项次"].ToString();
                    dr["工艺"] = dr1["工艺"].ToString();
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
            this.Text = "刀模费用信息编辑";
        

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
            textBox1.Text = cPROCESSING_TECHNOLOGY.GETID();
            bind();
            ADD_OR_UPDATE = "ADD";
        }
        private void save()
        {
            btnSave.Focus();
            //dgvfoucs();
            if (dt.Rows.Count > 0)
            {
                DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "工艺 IS NOT NULL");
                if (dtx.Rows.Count > 0)
                {
                    cPROCESSING_TECHNOLOGY.EMID = LOGIN.EMID;
                    cPROCESSING_TECHNOLOGY.PTID = textBox1.Text;
                    cPROCESSING_TECHNOLOGY.MATERIAL_TYPE  = textBox2.Text;
                    cPROCESSING_TECHNOLOGY.save(dtx);
                    IFExecution_SUCCESS = cPROCESSING_TECHNOLOGY.IFExecution_SUCCESS;
                    hint.Text = cPROCESSING_TECHNOLOGY.ErrowInfo;
                    if (IFExecution_SUCCESS)
                    {
                      
                        bind();
                    }
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
           if (textBox1 .Text  == "")
           {
               hint.Text = "编号不能为空";

           }
           else if (textBox2.Text == "")
           {
               hint.Text = "加工内容不能为空";
           }
           else  if(juage2())
           {
            
               b = true;
            }
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE PTID='{0}'",bc.RETURN_PTID(textBox2 .Text ))))
            {
                hint.Text = string.Format("加工材质 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
            return b;
        }
        #region juage2()
  
        private bool juage2()
        {
            bool b = false;
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "工艺 IS NOT NULL");
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    if (dr["工艺"].ToString() == "")
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的按工艺不能为空", dr["项次"].ToString());
                        break;
                    }
                }
            }
            else
            {
             
                b = true;
                hint.Text = "至少有一项工艺才能保存";

            }
            return b;
        }
        #endregion
     
        private void btnDel_Click(object sender, EventArgs e)
        {
            string id = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
           /* if (bc.exists(string.Format("SELECT * FROM WORKORDER_MST WHERE PTID='{0}'", bc.RETURN_PTID(textBox2.Text))))
            {
                hint.Text = string.Format("加工材质 {0} 已经在工单中使用不允许删除", textBox2.Text);
             
            }
            else
            {
               
            }*/
                 if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                  
                    basec.getcoms("DELETE PROCESSING_TECHNOLOGY_DET WHERE PTID='"+textBox1 .Text +"'");
                    basec.getcoms("DELETE PROCESSING_TECHNOLOGY_MST WHERE PTID='" +textBox1 .Text + "'");
                    bind();
                    ClearText();
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
                if (dataGridView1.Columns[i].DataPropertyName == "项次")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                   
                }
             
            }
   
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            dataGridView1.Columns["工艺"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["项次"].ReadOnly = true;
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
          
             
             
            
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            /*BASE_INFO.WAREINFO FRM = new CSPSS.BASE_INFO.WAREINFO();
            FRM.PROCESSING_TECHNOLOGY_USE();
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
            IF_DOUBLE_CLICK = false;
            BASE_INFO.SUN_SCREEN FRM = new CSPSS.BASE_INFO.SUN_SCREEN();
            FRM.PARAMETERS_SELECT = 4;
            FRM.ShowDialog();
            if (IF_DOUBLE_CLICK)
            {
                dt.Rows[intC]["工艺"] = PROCESSING_TECHNOLOGY;
                dataGridView1.CurrentCell = dataGridView1["按工艺起机", dataGridView1.CurrentCell.RowIndex];

            }
        
  
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
       
         
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    if (dataGridView1.Columns[i].DataPropertyName == "按工艺起机")
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00";
                    }
                    else
                    {
                        dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00";
                    }
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
