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
    public partial class RECEIVABLET : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private decimal _ORDER_COUNT;
        public decimal ORDER_COUNT
        {
            set { _ORDER_COUNT = value; }
            get { return _ORDER_COUNT; }

        }
        private decimal _HAVE_TAX_UNIT_PRICE;
        public decimal HAVE_TAX_UNIT_PRICE
        {
            set { _HAVE_TAX_UNIT_PRICE = value; }
            get { return _HAVE_TAX_UNIT_PRICE; }

        }
        private decimal _ORDER_AMOUNT;
        public decimal ORDER_AMOUNT
        {
            set { _ORDER_AMOUNT = value; }
            get { return _ORDER_AMOUNT; }

        }

        private decimal _WAIT_NOTICE_AMOUNT;
        public decimal WAIT_NOTICE_AMOUNT
        {
            set { _WAIT_NOTICE_AMOUNT = value; }
            get { return _WAIT_NOTICE_AMOUNT; }

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
        protected int i, j;
        protected int M_int_judge, t;
        basec bc = new basec();
        CRECEIVABLE cRECEIVABLE = new CRECEIVABLE();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        RECEIVABLE F1 = new RECEIVABLE();
        CNO_PAPER_OFFER cno_paper_offer = new CNO_PAPER_OFFER();
        CPN_PRODUCTION_INSTRUCTIONS cpn_production_instructions = new CPN_PRODUCTION_INSTRUCTIONS();
        StringBuilder sqb = new StringBuilder();
        string varDate = DateTime.Now.ToString("yyy/MM/dd").Replace("-", "/");
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        private string _EDIT;
        public string EDIT
        {
            set { _EDIT = value; }
            get { return _EDIT; }
        }
        public RECEIVABLET()
        {
            InitializeComponent();
        }
        public RECEIVABLET(RECEIVABLE Frm)
        {
            InitializeComponent();
            F1 = Frm;
        }
        private void RECEIVABLET_Load(object sender, EventArgs e)
        {
            right();
            WAIT_NOTICE_AMOUNT = 0;
            bind();
            try
            {
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #region right
        private void right()
        {
            DataTable dtx = cedit_right.RETURN_RIGHT_LIST("收款维护", LOGIN.USID);
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
        #region bind
        private void bind()
        {
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
            DataTable dtx = basec.getdts(cRECEIVABLE.sql +" where A.RCID='" + IDO + "' ORDER BY  B.RCKEY ASC ");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Text = dtx.Rows[0]["订单编号"].ToString();
                dt = cRECEIVABLE.GET_CALCULATE(dtx);
                textBox4.Text = dtx.Rows[0]["含税单价"].ToString();
                if (dt.Rows.Count > 0 && dt.Rows.Count < 6)
                {
                    int n = 6 - dt.Rows.Count;
                    for (int i = 0; i < n; i++)
                    {
                        DataRow dr = dt.NewRow();
                        int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                        dr["项次"] = Convert.ToString(b1 + 1);
                        dr["收款日期"] = varDate;
                        dt.Rows.Add(dr);
                    }
                }
            }
            else
            {
                dt = total1();
            }
            sqb = new StringBuilder();
            sqb.Append(cRECEIVABLE.sqlo);
            sqb.Append(" WHERE C.ORDER_ID='" + comboBox1.Text + "'");
            sqb.Append(" GROUP BY C.ORDER_ID,E.CName,C.WAREID,C.HAVE_TAX_UNIT_PRICE, C.PRODUCTION_COUNT,B.TAX_RATE  ORDER BY C.ORDER_ID ASC");
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                textBox6.Text = dtx.Rows[0]["实收金额"].ToString();
                textBox7.Text = dtx.Rows[0]["待收金额"].ToString();
            }
            dataGridView1.DataSource = dt;
           dgvStateControl();
        }
        #endregion

        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.ClearSelection();//加载不选中第一列
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            dataGridView1.Columns["项次"].Width =40;
            dataGridView1.Columns["收款日期"].Width = 120;
            for (i = 0; i < numCols1; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
            }
            dataGridView1.Columns["未税单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["数量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["税率"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["未税金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["税额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["含税金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["待收金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }

            dataGridView1.Columns["收款日期"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["未税单价"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["数量"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["税率"].DefaultCellStyle.BackColor = CCOLOR.CUSTOMER_YELLOW;
            dataGridView1.Columns["税率"].HeaderText = "税率(%)";
        }
        #endregion
     
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cRECEIVABLE.GetTableInfo();
        
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dr["收款日期"] = varDate;
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
            else if (ADD_OR_UPDATE!="UPDATE" && bc.exists(cRECEIVABLE.sql + " WHERE C.ORDER_ID='"+comboBox1.Text +"'"))
            {
                hint.Text = string.Format("订单编号 {0} 已经有收款记录，请查询该订单编号后双击进入修改", comboBox1.Text);
                b = true;
            }
            else if (textBox5.Text  == "")
            {
                hint.Text = string.Format("订单金额不能为空");
                b = true;
            }
            else
            {
                DataTable  dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "未税单价 IS NOT NULL");
                if (dtx.Rows.Count > 0)
                {
                    for (i = 0; i < dtx.Rows.Count; i++)
                    {
                        DateTime temp = DateTime.MinValue;
                
                        if (!DateTime.TryParse(dtx.Rows[i]["收款日期"].ToString(), out temp))
                        {
                            hint.Text = string.Format("第 {0} 行收款日期格式不正确 需为格式yyyy/MM/dd", i + 1);
                            b = true;
                            break;
                        }
                
                       else if (string.IsNullOrEmpty(dtx.Rows[i]["未税单价"].ToString()))
                        {
                            hint.Text = string.Format("第 {0} 行未税单价不能为空", i + 1);
                            b = true;
                            break;
                        }
                        else if (bc.yesno(dtx.Rows[i]["未税单价"].ToString()) == 0)
                        {
                            hint.Text = string.Format("第 {0} 行未税单价只能输入数字", i + 1);
                            b = true;
                            break;
                        }
                        else if (string.IsNullOrEmpty(dtx.Rows[i]["数量"].ToString()))
                        {
                            hint.Text = string.Format("第 {0} 行数量不能为空", i + 1);
                            b = true;
                            break;
                        }
                        else if (bc.yesno(dtx.Rows[i]["数量"].ToString()) == 0)
                        {
                            hint.Text = string.Format("第 {0} 行数量只能输入数字", i + 1);
                            b = true;
                            break;
                        }
                        else if (string.IsNullOrEmpty(dtx.Rows[i]["税率"].ToString()))
                        {
                            hint.Text = string.Format("第 {0} 行税率不能为空", i + 1);
                            b = true;
                            break;
                        }
                        else if (bc.yesno(dtx.Rows[i]["税率"].ToString()) == 0)
                        {
                            hint.Text = string.Format("第 {0} 行税率只能输入数字", i + 1);
                            b = true;
                            break;
                        }
              
                        else if (!string.IsNullOrEmpty(dtx.Rows[i]["待收金额"].ToString()))
                        {
                            WAIT_NOTICE_AMOUNT = decimal.Parse(dtx.Rows[i]["待收金额"].ToString());
                            if (WAIT_NOTICE_AMOUNT < 0)
                            {
                                hint.Text = string.Format("第 {0} 行待收金额不能出现负数", i + 1);
                                b = true;
                                break;
                            }
                        }
                    }
                }
            }
            return b;
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
            textBox6.Text = "";
            textBox7.Text = "";
        }
        #region save
        private void btnSave_Click(object sender, EventArgs e)
        {

            save();
        }
        #endregion
        private void save()
        {
            btnSave.Focus();
            string INITIAL_MAKERID = "";
            DataTable dtt = bc.getdt(cRECEIVABLE.sql + " WHERE A.RCID='" + IDO + "'");
            if (dtt.Rows.Count > 0)
            {
                INITIAL_MAKERID = bc.RETURN_APPOINT_UNTIL_CHAR(dtt.Rows[0]["制单人编号"].ToString(), 1, ' ');

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
                DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "未税单价 IS NOT NULL ");
                if (dtx.Rows.Count > 0)
                {
                    cRECEIVABLE.RCID = IDO;
                    cRECEIVABLE.PNID = bc.getOnlyString("SELECT PNID FROM PN_PRODUCTION_INSTRUCTIONS WHERE ORDER_ID='" + comboBox1.Text + "'");
                    cRECEIVABLE.EMID = LOGIN.EMID;
                    cRECEIVABLE.IF_IMPORT = false;
                    cRECEIVABLE.save("RECEIVABLE_MST", "RECEIVABLE_DET", "RCID", IDO, dtx);
                    IFExecution_SUCCESS = true;
                    bind();
                    F1.bind();
                    //F1.search();
                }
                else
                {
                    hint.Text = "至少有一项才能保存且该项未税单价栏位不能为空！";

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
                    basec.getcoms("DELETE RECEIVABLE_MST WHERE RCID='" + IDO+ "'");
                    basec.getcoms("DELETE RECEIVABLE_DET WHERE RCID='" + IDO + "'");
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
                string varDate = DateTime.Now.ToString("yyy/MM/dd").Replace("-", "/");
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
                        dr["收款日期"] = varDate;
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
                if (dataGridView1.Columns[columnsindex].Name == "未税单价" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView1.Columns[columnsindex].Name == "数量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView1.Columns[columnsindex].Name == "税率" && bc.yesno(e.FormattedValue.ToString()) == 0)
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
           
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                i = dataGridView1.CurrentCell.RowIndex;
                string v1 = dataGridView1["未税单价", rowsindex].FormattedValue.ToString();
                string v2 = dataGridView1["税率", rowsindex].FormattedValue .ToString();
                string v3=dataGridView1["数量", rowsindex].FormattedValue .ToString();
                string v4 = dataGridView1["收款日期", rowsindex].FormattedValue.ToString();
                if (v1 == "")
                {
                }
                else  if (bc.exists (cRECEIVABLE .sqlo+" WHERE C.ORDER_ID='"+comboBox1 .Text +"'") && textBox5.Text == "")
                {
                    hint.Text = "订单金额不能为空";
                }
          
                else if (bc.yesno(v1) == 0)
                {
                    hint.Text = string.Format("第 {0} 行未税单价只能输入数字", i + 1);
                }
                else if (v2 == "")
                {
                    hint.Text = string.Format("第 {0} 行税率不能为空", i + 1);
                }
                else if (bc.yesno(v2) == 0)
                {
                    hint.Text = string.Format("第 {0} 行税率只能输入数字", i + 1);
                }
                else if (v3 == "")
                {
                    hint.Text = string.Format("第 {0} 行数量不能为空", i + 1);
                }
                else if (bc.yesno(v3) == 0)
                {
                    hint.Text = string.Format("第 {0} 行数量只能输入数字", i + 1);
                }
                else  if (rowsindex >= 1)
                {
                    string v5 = dataGridView1["收款日期", rowsindex-1].FormattedValue.ToString();
                    if (v5 =="")
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
            ADD_OR_UPDATE = "ADD";
            IFExecution_SUCCESS = false;
            IDO = cRECEIVABLE.GETID();
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
                decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0;
                string v1 = dataGridView1["未税单价", rowsindex].Value.ToString();
                string v2 = dataGridView1["税率", rowsindex].Value.ToString();
                string v3 = dataGridView1["数量", rowsindex].Value.ToString();
                DateTime temp = DateTime.MinValue;
                if (dataGridView1["未税单价", rowsindex].FormattedValue.ToString() != "")
                {
                    d1 = decimal.Parse(dt.Rows[rowsindex]["未税单价"].ToString());
                }
                if (dataGridView1["数量", rowsindex].FormattedValue.ToString() != "")
                {
                    d2 = decimal.Parse(dt.Rows[rowsindex]["数量"].ToString());
                }
                if (dataGridView1["税率", rowsindex].FormattedValue.ToString() != "")
                {
                    d3 = decimal.Parse(dt.Rows[rowsindex]["税率"].ToString());
                }
                if (rowsindex == 0)
                {
                    if (d1 * d2*d3 / 100 > 0)
                    {

                        dataGridView1["未税金额", rowsindex].Value = (d1 * d2).ToString("0.00");
                        dataGridView1["税额", rowsindex].Value = (d1 * d2 * d3 / 100).ToString("0.00");
                        dataGridView1["含税金额", rowsindex].Value = (d1 * d2 * (1 + d3 / 100)).ToString("0.00");
                        if (ORDER_AMOUNT > 0)
                        {
                            dataGridView1["待收金额", rowsindex].Value = (ORDER_AMOUNT - (d1 * d2*(1 + d3 / 100))).ToString("0.00");

                        }
                    }
                }
                else
                {
                    if (d1 * d2*d3 / 100 > 0)
                    {
                        if (dataGridView1["待收金额", rowsindex - 1].FormattedValue.ToString() != "")
                        {
                            d4 = decimal.Parse(dt.Rows[rowsindex - 1]["待收金额"].ToString());
                        }
                        dataGridView1["未税金额", rowsindex].Value = (d1 * d2).ToString("0.00");
                        dataGridView1["税额", rowsindex].Value = (d1 * d2 *d3/ 100).ToString("0.00");
                        dataGridView1["含税金额", rowsindex].Value = (d1 *d2* (1 + d3 / 100)).ToString("0.00");

                        if (ORDER_AMOUNT > 0)
                        {
                            dataGridView1["待收金额", rowsindex].Value = (d4 - (d1 * d2*(1 + d3 / 100))).ToString("0.00");

                        }
                    }

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
            FRM.RECEIVABLE_USE();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = FRM.ORDER_ID;
                textBox1.Text = FRM.CNAME;
                textBox2.Text = FRM.WNAME;
                textBox3.Text = FRM.PROCESS_COUNT;
            }
         
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
           
            ORDER_COUNT = 0;
            HAVE_TAX_UNIT_PRICE = 0;
            ORDER_AMOUNT = 0;
            DataTable dtx = bc.getdt(cpn_production_instructions.sql + " WHERE A.ORDER_ID='"+comboBox1 .Text +"'");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Text = dtx.Rows[0]["订单编号"].ToString();
                textBox1.Text = dtx.Rows[0]["客户名称"].ToString();
                textBox3.Text = dtx.Rows[0]["生产数量"].ToString();
                textBox4.Text = dtx.Rows[0]["含税单价"].ToString();
                if (!string.IsNullOrEmpty(dtx.Rows[0]["含税单价"].ToString()))
                {
                    textBox5.Text = (decimal.Parse(dtx.Rows[0]["生产数量"].ToString()) * decimal.Parse(dtx.Rows[0]["含税单价"].ToString())).ToString();
                }
                if (!string.IsNullOrEmpty(dtx.Rows[0]["生产数量"].ToString()))
                {
                    ORDER_COUNT = decimal.Parse(dtx.Rows[0]["生产数量"].ToString());
                }
                if (!string.IsNullOrEmpty(dtx.Rows[0]["含税单价"].ToString()))
                {
                    HAVE_TAX_UNIT_PRICE = decimal.Parse(dtx.Rows[0]["含税单价"].ToString());
                }
                if (ORDER_COUNT * HAVE_TAX_UNIT_PRICE > 0)
                {
                    ORDER_AMOUNT = ORDER_COUNT * HAVE_TAX_UNIT_PRICE;
                }
                bind();
            }
            else
            {
                ClearText();
            }
            try
            {

            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }
    }
}
