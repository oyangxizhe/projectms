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

namespace CSPSS.USER_MANAGE
{
    public partial class LOGIN_INFO : Form
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
        CUSER cuser = new CUSER();
        protected int M_int_judge, i;
        protected int select;
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        public LOGIN_INFO()
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
           
            dt = basec.getdts(cuser .sqlf );
            dataGridView1.DataSource = dt;
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
            string sql1 = @"
select 
A.PIID AS 项目编号,
B.PROJECT_ID AS 项目号,
A.OFFER_ID_SENVEN,
A.YEAR AS 年,
A.MONTH AS 月,
A.DATE AS 日期,
(SELECT ENAME FROM EmployeeInfo WHERE EMID=A.MAKERID) AS 产生编号制单人
from PRINTING_OFFER_ID_NO A 
LEFT JOIN PROJECT_INFO B ON A.PIID=B.PIID 
where A.DATE is null and SUBSTRING (offer_id_senven,3,2)>5 
ORDER BY A.PIID  DESC";
            dt2 = bc.getdt(sql1);
            dataGridView2.DataSource = dt2;
            
            string sql3= @"
select 
DISTINCT(A.PIID) AS 项目编号,
B.PROJECT_ID AS 项目号,
B.PROJECT_NAME AS 项目名称 
from PRINTING_OFFER_MST  A
LEFT JOIN PROJECT_INFO B ON A.PIID=B.PIID
WHERE A.PIID NOT IN (select DISTINCT(PIID)  from PRINTING_OFFER_ID_NO )
ORDER BY A.PIID DESC
";
            dt3 = bc.getdt(sql3);
            //dataGridView3.DataSource = dt3;
            dgvStateControl(dataGridView1);
            dgvStateControl(dataGridView2);
       
        }
        #region dgvStateControl
        private void dgvStateControl(DataGridView dgv)
        {
            int i;
            dgv.AllowUserToAddRows = false;
            dgv.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dgv.Columns.Count;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            for (i = 0; i < numCols1; i++)
            {

               dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
               dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
               
            
                dgv.EnableHeadersVisualStyles = false;
                dgv.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dgv.Columns.Count; i++)
            {
                dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dgv.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            for (i = 0; i < dgv.Columns.Count; i++)
            {
                dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dgv.Columns[i].ReadOnly = true;
            }
            
        }
        #endregion
    
        #region save
        private void save()
        {
                string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("/", "-");
            basec.getcoms(@"
UPDATE AUTHORIZATION_USER 
SET 
STATUS ='N',
COMPUTER_UPDATE_DATE='"+varDate +"',IF_COMPUTER_UPDATE='Y' WHERE STATUS='Y'");
            IFExecution_SUCCESS = true;
        }
        #endregion
        #region juage()
        private bool juage()
        {

            bool b = false;
            return b;
        }
        #endregion
        public void ClearText()
        {
      
        
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
        private void dgvEmployeeInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        

        }
        private void add()
        {
        }
      

        private void btnSave_Click(object sender, EventArgs e)
        {
            
            if (juage())
            {

            }
            else
            {
                save();
                bind();
         
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {


                //dt = bc.getdt(cuser .sqlf+" WHERE A.DEID LIKE '%"+textBox4.Text +"%' AND A.LOGIN_INFO LIKE '%"+textBox5 .Text +"%'");
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dgvStateControl(dataGridView1);

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
                string strSql = "DELETE FROM LOGIN_INFO WHERE DEID='" + id + "'";
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && ((!(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)))
            {

                if (dataGridView1.CurrentCell.ColumnIndex == 7 &&
                    dataGridView1["借方原币金额", dataGridView1.CurrentCell.RowIndex].Value.ToString() != null)
                {

                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else if (dataGridView1.CurrentCell.ColumnIndex == 9)
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
        private void dataGridView2_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) //判断是不是右键
            {
                Control control = new Control();
                Point ClickPoint = new Point(e.X, e.Y);
                control.GetChildAtPoint(ClickPoint);
                if (dataGridView2.HitTest(e.X, e.Y).RowIndex >= 0 && dataGridView2.HitTest(e.X, e.Y).ColumnIndex >= 0)//判断你点的是不是一个信息行里
                {
                    dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.HitTest(e.X, e.Y).RowIndex].Cells[dataGridView2.HitTest(e.X, e.Y).ColumnIndex];
                    ContextMenu con = new ContextMenu();
                    MenuItem menuDeleteknowledge = new MenuItem("复制");
                    menuDeleteknowledge.Click += new EventHandler(btndgvInfoCopy2_Click);
                    con.MenuItems.Add(menuDeleteknowledge);
                    this.dataGridView2.ContextMenu = con;
                    con.Show(dataGridView2, new Point(e.X + 10, e.Y));
                }
            }
        }
    
        private void btndgvInfoCopy_Click(object sender, EventArgs e)
        {

            dgvCopy(ref dataGridView1);
        }
        private void btndgvInfoCopy2_Click(object sender, EventArgs e)
        {

            dgvCopy(ref dataGridView2);
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

        private void LOGIN_INFO_Load(object sender, EventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }
  
    }
}
