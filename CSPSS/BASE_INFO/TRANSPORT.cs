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
    public partial class TRANSPORT : Form
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
        CTRANSPORT cTRANSPORT = new CTRANSPORT();

   
        protected int M_int_judge, i;
        protected int select;
        public TRANSPORT()
        {
            InitializeComponent();
        }
     

        private void DEPAET_Load(object sender, EventArgs e)
        {
            textBox1.Text = IDO;
            try
            {
              this.Icon = Resource1.xz_200X200;
                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            this.WindowState = FormWindowState.Maximized;
        }

        private void bind()
        {
            textBox6.Text = "11";
            dt = basec.getdts(cTRANSPORT .sql );
            dt = cTRANSPORT.RETURN_HAVE_ID_DT(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.AllowUserToAddRows = false;
            textBox2.Focus();
            textBox2.BackColor = Color.Yellow;
            textBox3.BackColor = Color.Yellow;
       
            textBox6.BackColor = Color.Yellow;
    
            textBox4.BackColor = Color.Yellow;
            textBox5.BackColor = Color.Yellow;
            textBox3.TextAlign = HorizontalAlignment.Right;
            textBox5.TextAlign = HorizontalAlignment.Right;
            textBox6.TextAlign = HorizontalAlignment.Right;
            textBox4.TextAlign = HorizontalAlignment.Right;
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
               
                dataGridView1.Columns["物流运输"].Width = 120;
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
            dataGridView1.Columns["含税单价一"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["未税单价一"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["含税单价二"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["未税单价二"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["未税起运费"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["含税起运费"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
        }
        #endregion
    
        #region save
        private void save()
        {
            cTRANSPORT.TRID  = textBox1.Text;
            cTRANSPORT.TRANSPORT= textBox2.Text;
            cTRANSPORT.TAX_UNIT_PRICE_ONE  = textBox3.Text;
            cTRANSPORT.TAX_UNIT_PRICE_TWO = textBox4.Text;
            cTRANSPORT.TAX_TRANSPORT_COST  = textBox5.Text;
            cTRANSPORT.TAX_RATE = textBox6.Text;
            cTRANSPORT.REMARK  = textBox7.Text;
            cTRANSPORT.MAKERID = LOGIN.EMID;
            cTRANSPORT.CUSTOMER_TYPE = comboBox1.Text;
            cTRANSPORT.save();
       
           
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
                hint.Text = "物流运输不能为空！";
            }
            else if (textBox3.Text == "")
            {
                b = true;
                hint.Text = "含税单价不能为空";
            }
            else if (textBox4.Text == "")
            {
                b = true;
                hint.Text = "含税单价不能为空";
            }
            else if (textBox5.Text == "")
            {
                b = true;
                hint.Text = "起运费不能为空";
            }
            else if (textBox6.Text == "")
            {
                b = true;
                hint.Text = "税率不能为空";
            }
        
            else if (textBox3.Text != "" && bc.yesno (textBox3 .Text )==0 ||  textBox4.Text != "" && bc.yesno (textBox4 .Text )==0 || 
                textBox5.Text != "" && bc.yesno (textBox5 .Text )==0 || bc.yesno (textBox6.Text )==0 || bc.yesno (textBox6 .Text )==0)
            {
                b = true;
                hint.Text = "单价 费用 税率 只能输入数字";
            }
            else if (comboBox1.Text == "")
            {
                b = true;
                hint.Text = "客户类别不能为空";
            }
          
            return b;

        }
        #endregion
        public void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox5.Text = "";
            textBox6.Text = "11";
    
            textBox7.Text = "";
            textBox7.Text = "";
            textBox4.Text = "";
            comboBox1.Text = "";
        }
 
        private void add()
        {
            ClearText();
            textBox1.Text = cTRANSPORT.GETID();
            textBox2.Focus();
            bind();

        }
      

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (juage())
            {

            }
            else
            {

                save();
                IFExecution_SUCCESS = cTRANSPORT.IFExecution_SUCCESS;
                hint.Text = cTRANSPORT.ErrowInfo;
                if (IFExecution_SUCCESS)
                {
                    add();
                }
            }
            try
            {
               
            }
            catch (Exception)
            {

            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {


                dt = bc.getdt(cTRANSPORT .sql +" WHERE  A.TRANSPORT LIKE '%"+textBox50 .Text +"%'");
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
            int row=dataGridView1 .CurrentCell .RowIndex ;
            StringBuilder sqb = new StringBuilder();
            sqb.AppendFormat("DELETE FROM TRANSPORT WHERE TRID='{0}'", textBox1 .Text );
            try
            {
                IFExecution_SUCCESS = false;
                basec.getcoms(sqb.ToString ());
                bind();
                textBox1.Text = cTRANSPORT.GETID();
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
            string v1 = Convert.ToString(dataGridView1["物流运输", dataGridView1.CurrentCell.RowIndex].Value).Trim();
            int i=dataGridView1 .CurrentCell .RowIndex ;

                textBox1.Text = bc.getOnlyString(string.Format(@"SELECT TRID FROM TRANSPORT WHERE TRANSPORT='{0}' AND
CUSTOMER_TYPE='{1}'", dt.Rows[i]["物流运输"].ToString(), dt.Rows[i]["客户类别"].ToString()));
                textBox2.Text = dt.Rows[i]["物流运输"].ToString();
                textBox3.Text = dt.Rows[i]["含税单价一"].ToString();
                textBox4.Text = dt.Rows[i]["含税单价二"].ToString();
                textBox5.Text = dt.Rows[i]["含税起运费"].ToString();
                textBox6.Text = bc.RETURN_UNTIL_CHAR(dt.Rows[i]["税率"].ToString(), '%');
                textBox7.Text = dt.Rows[i]["说明"].ToString();
                comboBox1.Text = dt.Rows[i]["客户类别"].ToString();
             
            
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                bc.dgvtoExcel(dataGridView1,this.Text );
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

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
