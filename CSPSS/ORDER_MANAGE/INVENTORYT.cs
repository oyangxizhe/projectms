using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using XizheC;
using System.Net;
using System.Web;
using System.Xml;
using System.Collections;
using System.Data.OleDb;
using System.Web.UI;
using System.Web.UI.Adapters;
using System.Web.UI.HtmlControls;
using System.Web.Util;



namespace CSPSS.ORDER_MANAGE
{
    public partial class INVENTORYT : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        StringBuilder sqb = new StringBuilder();

        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _GECOUNT;
        public string GECOUNT
        {
            set { _GECOUNT = value; }
            get { return _GECOUNT; }
        }
        private string _MRCOUNT;
        public string MRCOUNT
        {
            set { _MRCOUNT = value; }
            get { return _MRCOUNT; }

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
        private string _BILL_DATE;
        public string BILL_DATE
        {
            set { _BILL_DATE = value; }
            get { return _BILL_DATE; }
        }
        private string _EDIT;
        public string EDIT
        {
            set { _EDIT = value; }
            get { return _EDIT; }
        }
        protected int i, j;
        protected int M_int_judge, t;
        basec bc = new basec();
        CINVENTORY cinventory = new CINVENTORY();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        INVENTORY F1 = new INVENTORY();
        CNO_PAPER_OFFER cno_paper_offer = new CNO_PAPER_OFFER();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        CPN_PRODUCTION_INSTRUCTIONS cpn_production_instructions = new CPN_PRODUCTION_INSTRUCTIONS();
        string varDate = DateTime.Now.ToString("yyy/MM/dd").Replace("-", "/");
        public INVENTORYT()
        {
            InitializeComponent();
        }
        public INVENTORYT(INVENTORY Frm)
        {
            InitializeComponent();
            F1 = Frm;
        }
        private void INVENTORYT_Load(object sender, EventArgs e)
        {
        
            try
            {
                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        #region bind
        private void bind()
        {
            right();
          this.Icon = Resource1.xz_200X200;
            hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
            comboBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }

            DataTable dtx = basec.getdts(cinventory.sql +" where A.INID='" + IDO + "' ORDER BY  A.INKEY ASC ");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Text = dtx.Rows[0]["订单编号"].ToString();
                dt = cinventory.GET_CALCULATE(dtx);
                /*if (!string.IsNullOrEmpty(dt.Compute("SUM(入库数量)", "").ToString()))
                {
                    textBox4.Text = dt.Compute("SUM(入库数量)", "").ToString();
                }*/
                if (dt.Rows.Count > 0 && dt.Rows.Count < 6)
                {
                    int n = 6 - dt.Rows.Count;
                    for (int i = 0; i < n; i++)
                    {
                        DataRow dr = dt.NewRow();
                        int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                        dr["项次"] = Convert.ToString(b1 + 1);
                        dr["日期"] = varDate;
                        dt.Rows.Add(dr);
                    }
                }
            }
            else
            {
                dt = total1();
            }
            sqb = new StringBuilder();
            sqb.Append(cinventory.sqlo);
            sqb.Append(" WHERE C.ORDER_ID =  '"+comboBox1.Text + "'");
            sqb.Append(" GROUP BY C.ORDER_ID,E.CName,C.WAREID, C.PRODUCTION_COUNT  ORDER BY C.ORDER_ID ASC");
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                textBox4.Text = dtx.Rows[0]["已入库数量"].ToString();
                textBox5.Text = dtx.Rows[0]["已出货数量"].ToString();
            }
           dataGridView1.DataSource = dt;
           dgvStateControl();
        }
        #endregion
        #region right
        private void right()
        {
           DataTable  dtx = cedit_right.RETURN_RIGHT_LIST("库存维护", LOGIN.USID);
            btnAdd.Visible = false;
            btnSave.Visible = false;
            btnDel.Visible = false;
            label100.Visible = false;
            label101.Visible = false;
            label102.Visible = false;

            if (dtx.Rows.Count > 0)
            {

                if (dtx.Rows[0]["新增权限"].ToString() == "有权限")
                {
                    btnAdd.Visible = true;
                    btnSave.Visible = true;
                    label100.Visible = true;
                    label101.Visible = true;
                }
                if (dtx.Rows[0]["删除权限"].ToString() == "有权限")
                {
                    btnDel.Visible = true;
                    label102.Visible = true;
                }
                if (dtx.Rows[0]["修改权限"].ToString() == "有权限")
                {
                    btnSave.Visible = true;
                    label101.Visible = true;
                    EDIT = "有权限";
                }

            }
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            dataGridView1.Columns["项次"].Width =40;
            dataGridView1.Columns["日期"].Width = 120;
            dataGridView1.Columns["备注"].Width = 200;
            for (i = 0; i < numCols1; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
            }
            dataGridView1.Columns["入库数量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["出库数量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["库存结余"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            dataGridView1.Columns["入库数量"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["出库数量"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["日期"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["库存结余"].ReadOnly = true;
        }
        #endregion
     
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cinventory.GetTableInfo();

            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dr["日期"] = varDate;
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&(( !(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn) ))
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
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
      
        #region juage()
        private bool juage()
        {
            bool b = false;
            if (comboBox1.Text == "")
            {
                hint.Text = string.Format("订单编号不能为空");
                b = true;
            }
            else if (!bc.exists("SELECT * FROM PN_PRODUCTION_INSTRUCTIONS WHERE ORDER_ID='"+comboBox1 .Text +"'"))
            {
                hint.Text = string.Format("订单编号 {0} 不存在系统中", comboBox1 .Text );
                b = true;
            }
            else
            {
                
                DataTable  dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "入库数量 IS NOT NULL OR 出库数量 IS NOT NULL");
                if (dtx.Rows.Count > 0)
                {
                    for (i = 0; i < dtx.Rows.Count; i++)
                    {
                      
                        GECOUNT = dtx.Rows[i]["入库数量"].ToString();
                        MRCOUNT = dtx.Rows[i]["出库数量"].ToString();
                        BILL_DATE = dtx.Rows[i]["日期"].ToString();
                        DateTime temp = DateTime.MinValue;
                        if (string.IsNullOrEmpty(dtx.Rows[i]["入库数量"].ToString()) && string.IsNullOrEmpty(dtx.Rows[i]["出库数量"].ToString()))
                        {
                            hint.Text = string.Format("第 {0} 行出现空行", i + 1);
                            b = true;
                            break;
                        }
                        else if ((!string.IsNullOrEmpty(GECOUNT) || !string.IsNullOrEmpty(MRCOUNT)) && BILL_DATE == "")
                        {

                            hint.Text = string.Format("第 {0} 行日期不能为空", i + 1);
                            b = true;
                            break;
                        }
                        else if (!DateTime.TryParse(BILL_DATE, out temp))
                        {
                            hint.Text = string.Format("第 {0} 行日期格式不正确 需为格式yyyy/MM/dd", i + 1);
                            b = true;
                            break;
                        }
                    }
                }
            
               
            }
            return b;
        }
        #endregion
        #region dgvDataSourceChanged
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
    
        }
        #endregion
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
       
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        private void ClearText()
        {
            comboBox1.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
       
        }
        #region save
        private void btnSave_Click(object sender, EventArgs e)
        {

            save();
        }
        #endregion
        private void save()
        {
     
            try
            {
                btnSave.Focus();
                string INITIAL_MAKERID = "";
                DataTable   dtt = bc.getdt(cinventory.sql + " WHERE A.INID='" + IDO + "'");
                if (dtt.Rows.Count > 0)
                {
                    INITIAL_MAKERID = bc.RETURN_APPOINT_UNTIL_CHAR(dtt.Rows[0]["制单人编号"].ToString(),1,' ');

                    if (EDIT != "有权限" && LOGIN.EMID != INITIAL_MAKERID)
                    {
                        //MessageBox.Show(INITIAL_MAKERID+","+INITIAL_MAKERID .Length .ToString ()+"+"+LOGIN.EMID +","+LOGIN .EMID .Length .ToString ());
                        hint.Text = "本账号无修改权限！";
                        return;
                    }
                }
                if (juage())
                {

                }
                else
                {
                    DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "入库数量 IS NOT NULL OR 出库数量 IS NOT NULL");
                    if (dtx.Rows.Count > 0)
                    {
                        cinventory.INID = IDO;
                        cinventory.PNID = bc.getOnlyString("SELECT PNID FROM PN_PRODUCTION_INSTRUCTIONS WHERE ORDER_ID='" + comboBox1.Text + "'");
                        cinventory.EMID = LOGIN.EMID;
                        cinventory.WAREID = "";
                        cinventory.WNAME = "";
                        cinventory.IF_IMPORT = false;
                        cinventory.save("INVENTORY_MST", "inventory_DET", "INID", IDO, dtx);
                        IFExecution_SUCCESS = true;
                        bind();
                        F1.bind();
                        //F1.search();
                    }
                    else
                    {
                        hint.Text = "至少有一项入库数量或出库数量不为空的行才能保存！";

                    }
                }
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
      
        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
               if (MessageBox.Show("确定要删除该条凭证吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    basec.getcoms("DELETE INVENTORY_MST WHERE INID='" + IDO+ "'");
                    basec.getcoms("DELETE INVENTORY_DET WHERE INID='" + IDO + "'");
                    bind();
                    ClearText();
                    textBox2.Text = "";
                    F1.bind();
              
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region dgvDoubleClick
        #endregion
        #region dgvCellEnter
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int a = dataGridView1.CurrentCell.ColumnIndex;
                int b = dataGridView1.CurrentCell.RowIndex;
                int c = dataGridView1.Columns.Count - 1;
                int d = dataGridView1.Rows.Count - 1;
                string varDate = DateTime.Now.ToString("yyy/MM/dd").Replace("-", "/");
                if (a == c && b == d)
                {
                    if (dt.Rows.Count >= 6)
                    {

                        DataRow dr = dt.NewRow();
                        int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                        dr["项次"] = Convert.ToString(b1 + 1);
                        dr["日期"] = varDate;
                        dt.Rows.Add(dr);
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
        #region dgvCellValidating
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {

            try
            {
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                int columnsindex = dataGridView1.CurrentCell.ColumnIndex;
                if (dataGridView1.Columns[columnsindex].Name == "入库数量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView1.Columns[columnsindex].Name == "出库数量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
           
            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        #endregion

        private void dataGridView1_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            try
            {
                decimal d4 = 0, d5 = 0;
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                string v1 = dataGridView1["入库数量", rowsindex].Value.ToString();
                string v2 = dataGridView1["出库数量", rowsindex].Value.ToString();
                string v3=dataGridView1["日期", rowsindex].Value.ToString();
        
                if ((v1 != "" || v2!= "") && v3=="")
                {
                    e.Cancel = true;
                    hint.Text = string.Format("第 {0} 行日期不能空行", rowsindex+1);
                }
           
                else if (rowsindex >= 1)
                {
                    if (dataGridView1["入库数量", rowsindex - 1].FormattedValue.ToString() != "")
                    {
                        d4 = decimal.Parse(dt.Rows[rowsindex - 1]["入库数量"].ToString());
                    }
                    if (dataGridView1["出库数量", rowsindex - 1].FormattedValue.ToString() != "")
                    {
                        d5 = decimal.Parse(dt.Rows[rowsindex - 1]["出库数量"].ToString());
                    }
                    if (d4 == 0 && d5 == 0)
                    {
                        //e.Cancel = true;
                        hint.Text = string.Format("第 {0} 行出现空行", rowsindex);
                    }
                }
            }
            catch (Exception)
            {

            }    
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            ClearText();
            IFExecution_SUCCESS = false;
            IDO = cinventory.GETID();
            bind();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            try
            {
                bind();
                F1.bind();    
            }
            catch (Exception)
            {
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                int columnsindex = dataGridView1.CurrentCell.ColumnIndex;
                decimal d1 = 0, d2 = 0, d3 = 0;
                string v1 = dataGridView1["入库数量", rowsindex].Value.ToString();
                string v2 = dataGridView1["出库数量", rowsindex].Value.ToString();
                string v3 = dataGridView1["日期", rowsindex].Value.ToString();
                DateTime temp = DateTime.MinValue;
                if (dataGridView1["入库数量", rowsindex].FormattedValue.ToString() != "")
                {
                    d1 = decimal.Parse(dt.Rows[rowsindex]["入库数量"].ToString());
                }
                if (dataGridView1["出库数量", rowsindex].FormattedValue.ToString() != "")
                {
                    d2 = decimal.Parse(dt.Rows[rowsindex]["出库数量"].ToString());
                }
                if (rowsindex == 0)
                {
                    dataGridView1["库存结余", rowsindex].Value = (d1 - d2).ToString();
                }
                else if (!DateTime.TryParse(v3, out temp))
                {
                    hint.Text = string.Format("第 {0} 行日期格式不正确 需为格式yyyy/MM/dd", rowsindex);
                }
                else if(rowsindex >=1)
                {
                    if (dataGridView1["库存结余", rowsindex-1].FormattedValue.ToString() != "")
                    {
                        d3 = decimal.Parse(dt.Rows[rowsindex-1]["库存结余"].ToString());
                    }
                    dataGridView1["库存结余", rowsindex].Value = (d3 + d1 - d2).ToString(); 
                }
            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            PN_PRODUCTION_INSTRUCTIONS FRM = new PN_PRODUCTION_INSTRUCTIONS();
            FRM.WindowState = FormWindowState.Normal;
            FRM.INVENTORY_USE();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = FRM.ORDER_ID;
                textBox1.Text = FRM.CNAME;
                textBox3.Text = FRM.PROCESS_COUNT;
                textBox2.Text = FRM.WNAME;
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
               DataTable  dtx = bc.getdt(cpn_production_instructions .sql + " WHERE A.ORDER_ID='" + comboBox1.Text + "'");
               if (dtx.Rows.Count > 0)
               {
                   comboBox1.Text = dtx.Rows[0]["订单编号"].ToString();
                   textBox1.Text = dtx.Rows[0]["客户名称"].ToString();
                   textBox3.Text = dtx.Rows[0]["生产数量"].ToString();

                   bind();
               }
               else
               {
                   ClearText();
               }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
    }
}
