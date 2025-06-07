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
    public partial class DOOR_PARAMETERST : Form
    {
        DataTable dt = new DataTable();
        basec bc=new basec ();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private static string _RETURN_DATA;
        public static string RETURN_DATA
        {
            set { _RETURN_DATA = value; }
            get { return _RETURN_DATA; }
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
        DOOR_PARAMETERS F1 = new DOOR_PARAMETERS();
        protected int M_int_judge, i;
        protected int select;
        CDOOR_PARAMETERS cDOOR_PARAMETERS = new CDOOR_PARAMETERS();
       
        public DOOR_PARAMETERST()
        {
            InitializeComponent();
        }
        public DOOR_PARAMETERST(DOOR_PARAMETERS FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        private void DOOR_PARAMETERST_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            try
            {
                textBox1.Text = IDO;
                bind();

            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } 
         
        }
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cDOOR_PARAMETERS.GetTableInfo();
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

        public void ClearText()
        {
            textBox2.Text = "";
           
       
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
        
            try
            {
                bind();
            }
            catch (Exception )
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            DataTable dtx = basec.getdts(cDOOR_PARAMETERS.sql + " where A.DPID='" +textBox1 .Text  + "' ORDER BY  B.DPID ASC ");
            if (dtx.Rows.Count > 0)
            {
               
                dt = cDOOR_PARAMETERS.GetTableInfo();
                textBox2.Text = dtx.Rows[0]["印刷用纸或芯纸"].ToString();
                foreach (DataRow dr1 in dtx.Rows)
                {
           
                    DataRow dr = dt.NewRow();
                    dr["项次"] = dr1["项次"].ToString();
                    dr["值"] = dr1["值"].ToString();
                    dr["客户类别"] = dr1["客户类别"].ToString();
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
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void add()
        {
            ClearText();
            textBox1 .Text = cDOOR_PARAMETERS.GETID();
        
            bind();
         
            ADD_OR_UPDATE = "ADD";
           

        }
        private void save()
        {
            try
            {
                btnSave.Focus();
                //dgvfoucs();
                if (dt.Rows.Count > 0)
                {
                    DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "值 IS NOT NULL");
                    if (dtx.Rows.Count > 0)
                    {
                        cDOOR_PARAMETERS.EMID = LOGIN.EMID;
                        cDOOR_PARAMETERS.DPID = textBox1.Text;
                        cDOOR_PARAMETERS.DOOR_PARAMETERS = textBox2.Text;
                        cDOOR_PARAMETERS.save(dtx);
                        IFExecution_SUCCESS = cDOOR_PARAMETERS.IFExecution_SUCCESS;
                        hint.Text = cDOOR_PARAMETERS.ErrowInfo;
                        if (IFExecution_SUCCESS)
                        {
                            bind();
                        }
                        /*F1.Bind();
                        F1.search();*/
                    }
                }
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
                hint.Text = "印刷用纸或芯纸编号不能为空！";
                b = true;
            }
            else if (textBox2.Text == "")
            {
                hint.Text = "印刷用纸或芯纸不能为空！";
                b = true;
            }
           else if(juage2())
           {
            
               b = true;
            }
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE DPID='{0}'",bc.RETURN_DPID(textBox2 .Text ))))
            {
                hint.Text = string.Format("印刷用纸或印刷用纸或芯纸 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
            return b;
        }
        #region juage2()

        private bool juage2()
        {
            bool b = false;
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "值 IS NOT NULL");
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {
                    if (dr["值"].ToString() == "")
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的值不能为空", dr["项次"].ToString());
                        break;
                    }
                    else if (bc.yesno(dr["值"].ToString())==0)
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的值只能输入数字", dr["项次"].ToString());
                        break;
                    }
                    else if (dr["客户类别"].ToString()=="")
                    {

                        b = true;
                        hint.Text = string.Format("项次 {0} 的客户类别不能为空", dr["项次"].ToString());
                        break;
                    }

                }
            }
            else
            {

                b = true;
                hint.Text = "至少有一项值才能保存";

            }
            return b;
        }
        #endregion
     
        private void btnDel_Click(object sender, EventArgs e)
        {
    
            try
            {
                string id = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                /* if (bc.exists(string.Format("SELECT * FROM WORKORDER_MST WHERE DPID='{0}'", bc.RETURN_DPID(textBox2.Text))))
                 {
                     hint.Text = string.Format("印刷用纸或印刷用纸或芯纸 {0} 已经在工单中使用不允许删除", textBox2.Text);
             
                 }
                 else
                 {
               
                 }*/
                if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    basec.getcoms("DELETE DOOR_PARAMETERS_MST WHERE DPID='" + textBox1.Text + "'");
                    basec.getcoms("DELETE DOOR_PARAMETERS_DET WHERE DPID='" + textBox1.Text + "'");
                    bind();
                    ClearText();
                    add();
                    textBox1.Text = "";
                    F1.load();

                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

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


       
            dataGridView1.Columns["值"].DefaultCellStyle.BackColor = Color.Yellow;

            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["客户类别"].ReadOnly = true;
           dataGridView1 .Columns ["值"].DefaultCellStyle .Alignment =DataGridViewContentAlignment .BottomRight;
            dataGridView1.Columns["项次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        #endregion
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

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
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
       
            //dgvfoucs();

        }
 

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            int rows = dataGridView1.CurrentCell.RowIndex;
            int columns = dataGridView1.CurrentCell.ColumnIndex;
            if (dataGridView1.Columns[columns].DataPropertyName.ToString() == "客户类别")
            {
                TEMP FRM = new TEMP();
                FRM.PARAMETERS_SELECT = 1;
                FRM.ShowDialog();
                if (IF_DOUBLE_CLICK)
                {
                    dt.Rows[rows]["客户类别"] = RETURN_DATA;
                    dataGridView1.CurrentCell = dataGridView1["值", rows+1];
                }
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
