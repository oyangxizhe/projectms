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
    public partial class TEMP : Form
    {

        List<string> list = new List<string>();
        basec bc = new basec();
        CUSER cuser = new CUSER();
        DataTable dt = new DataTable();
        CSAMPLE_RELY_LIST csample_rely_list = new CSAMPLE_RELY_LIST();
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        private int _PARAMETERS_SELECT;
        public int PARAMETERS_SELECT
        {
            set { _PARAMETERS_SELECT = value; }
            get { return _PARAMETERS_SELECT; }

        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private string _PROJECT_ID;
        public string PROJECT_ID
        {
            set { _PROJECT_ID = value; }
            get { return _PROJECT_ID; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        CSPSS.OFFER_MANAGE.PROJECT_INFOT F1 = new CSPSS.OFFER_MANAGE.PROJECT_INFOT();
        public TEMP()
        {
            InitializeComponent();
        }
        public TEMP(OFFER_MANAGE .PROJECT_INFOT FRM)
        {
            F1 = FRM;
            InitializeComponent();
        }
        private void TEMP_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            bind();
        }
        #region bind()
        private void bind()
        {
            if (PARAMETERS_SELECT == 0)
            {
                dt = RETURN_O();
            }
            else if (PARAMETERS_SELECT == 1)
            {
                dt = RETURN_TWO();
            }
            else if (PARAMETERS_SELECT == 2)
            {
                dt = RETURN_THREE();
            }
            dataGridView1.DataSource = dt;
            dgvStateControl();
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            dataGridView1.AllowUserToAddRows = false;
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            dataGridView1.Columns["值"].Width = 40;
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
            dataGridView1.Columns["值"].ReadOnly = true;
            dataGridView1.Columns["值"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        #endregion
    
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
             
            }
            catch (Exception)
            {


            }
        }
        private DataTable RETURN_O()
        {
            dt.Columns.Add("值", typeof(string));
            list.Add("A 数据库");
            list.Add("B 数据库");
            list.Add("C 数据库");
            for (int i = 0; i < list.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr["值"] = list[i];
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private DataTable RETURN_TWO()
        {
            dt.Columns.Add("值", typeof(string));
            list.Add("A 材料门幅参数");
            list.Add("B 材料门幅参数");
            list.Add("C 材料门幅参数");
            for (int i = 0; i < list.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr["值"] = list[i];
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private DataTable RETURN_THREE()
        {
            dt.Columns.Add("值", typeof(string));
            list.Add("按平方");
            list.Add("按米计");
            for (int i = 0; i < list.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr["值"] = list[i];
                dt.Rows.Add(dr);
            }
            return dt;
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            F1.Close();
        }

        #region ProcessCmdKey
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&
             (
             (
              !(ActiveControl is System.Windows.Forms.TextBox) ||
              !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)
             )
             )
            {
                SendKeys.SendWait("{Tab}");
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
         
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            if (PARAMETERS_SELECT == 0)
            {
                CUSTOMER_INFOT.IF_DOUBLE_CLICK = true;
                CUSTOMER_INFOT.RETURN_DATA = dt.Rows[dataGridView1.CurrentCell.RowIndex]["值"].ToString();
                //MessageBox.Show(dt.Rows[dataGridView1.CurrentCell.RowIndex]["值"].ToString());
                this.Close();
            }
            else if (PARAMETERS_SELECT == 1)
            {
               DOOR_PARAMETERST.IF_DOUBLE_CLICK = true;
               DOOR_PARAMETERST.RETURN_DATA = dt.Rows[dataGridView1.CurrentCell.RowIndex]["值"].ToString();
                //MessageBox.Show(dt.Rows[dataGridView1.CurrentCell.RowIndex]["值"].ToString());
                this.Close();
            }
         
        }
    }
}
