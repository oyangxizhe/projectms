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

namespace CSPSS.ORDER_MANAGE
{
    public partial class INVENTORY : Form
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dtx = new DataTable();
        basec bc = new basec();
        protected int M_int_judge, i, look;
        protected int getdata;
        CINVENTORY cinventory = new CINVENTORY();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        StringBuilder sqb = new StringBuilder();
        public INVENTORY()
        {
            InitializeComponent();
        }
        private bool _IF_IMPORT_SUCCESS;
        public bool IF_IMPORT_SUCCESS
        {
            set { _IF_IMPORT_SUCCESS = value; }
            get { return _IF_IMPORT_SUCCESS; }
        }
        #region init
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(INVENTORY));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.btnToCSharp = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.hint = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.PictureBox();
            this.btnExit = new System.Windows.Forms.PictureBox();
            this.btnSearch = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Location = new System.Drawing.Point(0, 216);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(943, 400);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataSourceChanged += new System.EventHandler(this.dataGridView1_DataSourceChanged);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            this.dataGridView1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseUp);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.checkBox2);
            this.groupBox1.Controls.Add(this.btnToCSharp);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label21);
            this.groupBox1.Controls.Add(this.dateTimePicker2);
            this.groupBox1.Controls.Add(this.dateTimePicker1);
            this.groupBox1.Controls.Add(this.btnToExcel);
            this.groupBox1.Location = new System.Drawing.Point(3, 127);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(936, 84);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "查询条件";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(604, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 137;
            this.label1.Text = "显示编号";
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(583, 54);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(15, 14);
            this.checkBox2.TabIndex = 136;
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // btnToCSharp
            // 
            this.btnToCSharp.FlatAppearance.BorderSize = 0;
            this.btnToCSharp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToCSharp.Font = new System.Drawing.Font("宋体", 9F);
            this.btnToCSharp.Image = ((System.Drawing.Image)(resources.GetObject("btnToCSharp.Image")));
            this.btnToCSharp.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnToCSharp.Location = new System.Drawing.Point(761, 13);
            this.btnToCSharp.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnToCSharp.Name = "btnToCSharp";
            this.btnToCSharp.Size = new System.Drawing.Size(50, 64);
            this.btnToCSharp.TabIndex = 46;
            this.btnToCSharp.Text = "导入";
            this.btnToCSharp.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnToCSharp.UseVisualStyleBackColor = false;
            this.btnToCSharp.Click += new System.EventHandler(this.btnToCSharp_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(115, 20);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(140, 20);
            this.comboBox1.TabIndex = 135;
            this.comboBox1.DropDown += new System.EventHandler(this.comboBox1_DropDown);
            this.comboBox1.TextChanged += new System.EventHandler(this.comboBox1_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(56, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 134;
            this.label3.Text = "订单编号";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(34, 56);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(15, 14);
            this.checkBox1.TabIndex = 133;
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(299, 59);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(11, 12);
            this.label6.TabIndex = 132;
            this.label6.Text = "~";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(55, 56);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(53, 12);
            this.label21.TabIndex = 131;
            this.label21.Text = "日期期间";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(355, 50);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(141, 21);
            this.dateTimePicker2.TabIndex = 130;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Cursor = System.Windows.Forms.Cursors.Default;
            this.dateTimePicker1.Location = new System.Drawing.Point(114, 50);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(141, 21);
            this.dateTimePicker1.TabIndex = 129;
            // 
            // btnToExcel
            // 
            this.btnToExcel.FlatAppearance.BorderSize = 0;
            this.btnToExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToExcel.Font = new System.Drawing.Font("宋体", 9F);
            this.btnToExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnToExcel.Image")));
            this.btnToExcel.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnToExcel.Location = new System.Drawing.Point(847, 12);
            this.btnToExcel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(50, 64);
            this.btnToExcel.TabIndex = 11;
            this.btnToExcel.Text = "导出";
            this.btnToExcel.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnToExcel.UseVisualStyleBackColor = false;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(857, 95);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(29, 12);
            this.label11.TabIndex = 29;
            this.label11.Text = "退出";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(771, 95);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(29, 12);
            this.label12.TabIndex = 28;
            this.label12.Text = "搜索";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.hint);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.btnAdd);
            this.groupBox2.Controls.Add(this.btnExit);
            this.groupBox2.Controls.Add(this.btnSearch);
            this.groupBox2.Location = new System.Drawing.Point(3, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(936, 121);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "菜单栏";
            // 
            // hint
            // 
            this.hint.AutoSize = true;
            this.hint.Location = new System.Drawing.Point(400, 100);
            this.hint.Name = "hint";
            this.hint.Size = new System.Drawing.Size(29, 12);
            this.hint.TabIndex = 104;
            this.hint.Text = "hint";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(28, 95);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(29, 12);
            this.label17.TabIndex = 24;
            this.label17.Text = "新增";
            this.label17.Click += new System.EventHandler(this.label17_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.InitialImage = null;
            this.btnAdd.Location = new System.Drawing.Point(12, 20);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(60, 60);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.TabStop = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnExit
            // 
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.InitialImage = null;
            this.btnExit.Location = new System.Drawing.Point(843, 20);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(60, 60);
            this.btnExit.TabIndex = 19;
            this.btnExit.TabStop = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Image = ((System.Drawing.Image)(resources.GetObject("btnSearch.Image")));
            this.btnSearch.InitialImage = null;
            this.btnSearch.Location = new System.Drawing.Point(757, 20);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(60, 60);
            this.btnSearch.TabIndex = 18;
            this.btnSearch.TabStop = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // INVENTORY
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(942, 616);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "INVENTORY";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "库存信息汇总表";
            this.Load += new System.EventHandler(this.INVENTORY_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion
  
        #region bind
        public  void bind()
        {
            try
            {
                hint.ForeColor = Color.Red;
                if (checkBox2.Checked == true)
                {
                    dataGridView1.DataSource = null;
                    sqb = new StringBuilder();
                    sqb.Append(cinventory.sqlfi);
                    sqb.Append(" WHERE C.ORDER_ID LIKE '%" + comboBox1.Text + "%'");
                    if (LOGIN.POSITION == "AE")
                    {
                        sqb.Append(" AND E.CUID IN (SELECT CUID FROM CUSTOMERINFO_DET WHERE USER_MAKERID LIKE '%" + LOGIN.EMID + "%')");//只能查询登录者在客户信息里维护过的信息
                    }
                    string v1 = dateTimePicker1.Text + " 0:00:00";
                    string v2 = dateTimePicker2.Text + " 23:59:59";
                    if (checkBox1.Checked)
                    {
                        sqb.Append(" AND A.DATE  BETWEEN  '" + v1 + "' AND '" + v2 + "'");
                    }
                    if (IF_IMPORT_SUCCESS)
                    {
                        sqb.Append(" AND SUBSTRING (A.DATE,1,10)=CONVERT(varchar(12) , getdate(), 111 )");
                    }
                    sqb.Append(" GROUP BY A.INID,C.ORDER_ID,E.CName,C.WAREID, C.PRODUCTION_COUNT  ORDER BY C.ORDER_ID ASC");
                    dt = cinventory.RETURN_HAVE_ID_DT(bc.getdt(sqb.ToString()),true);
                }
                else
                {
                    sqb = new StringBuilder();
                    sqb.Append(cinventory.sqlo);
                    sqb.Append(" WHERE C.ORDER_ID LIKE '%" + comboBox1.Text + "%'");
                    if (LOGIN.POSITION == "AE")
                    {
                        sqb.Append(" AND E.CUID IN (SELECT CUID FROM CUSTOMERINFO_DET WHERE USER_MAKERID LIKE '%" + LOGIN.EMID + "%')");//只能查询登录者在客户信息里维护过的信息
                    }
                    string v1 = dateTimePicker1.Text + " 0:00:00";
                    string v2 = dateTimePicker2.Text + " 23:59:59";
                    if (checkBox1.Checked)
                    {
                        sqb.Append(" AND A.DATE  BETWEEN  '" + v1 + "' AND '" + v2 + "'");
                    }
                    if (IF_IMPORT_SUCCESS)
                    {
                        sqb.Append(" AND SUBSTRING (A.DATE,1,10)=CONVERT(varchar(12) , getdate(), 111 )");
                    }
                    sqb.Append(" GROUP BY C.ORDER_ID,E.CName,C.WAREID, C.PRODUCTION_COUNT  ORDER BY C.ORDER_ID ASC");
                    dt = cinventory.RETURN_HAVE_ID_DT(bc.getdt(sqb.ToString()),false);
                }
            
                if (dt.Rows.Count > 0)
                {
                    if (_IF_IMPORT_SUCCESS)
                    {
                    }
                    else
                    {
                        hint.Text = "";
                    }
                    dataGridView1.DataSource = dt;
                    dgvStateControl();
                }
                else
                {
                    hint.Text = "找不到所要搜索项！";
                    dataGridView1.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
  
        #endregion
        private void INVENTORY_Load(object sender, EventArgs e)
        {
            right();
            hint.Text = "";
            //this.WindowState = FormWindowState.Maximized;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";
            dateTimePicker2.CustomFormat = "yyyy/MM/dd";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
          this.Icon = Resource1.xz_200X200;
            dgvStateControl();
            //bind();
        }
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("库存维护", LOGIN.USID);
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
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
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

        #region add

        #endregion
        #region look

        #endregion
        #region override enter
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
        #region doubleclick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            int intCurrentRowNumber = this.dataGridView1.CurrentCell.RowIndex;
            string s1 = this.dataGridView1.Rows[intCurrentRowNumber].Cells["订单编号"].Value.ToString().Trim();
            if (getdata != 0)
            {
                if (getdata == 1)
                {
                
                 
                    this.Close();
                }

            }
            else
            {
                if (checkBox2.Checked==false)
                {
                    hint.Text = "在选中显示编号下才可进入编辑";
                 
                }
                else
                {
                    INVENTORYT FRM = new INVENTORYT(this);
                    FRM.ADD_OR_UPDATE = "UPDATE";
                    sqb = new StringBuilder();
                    sqb.Append(@"SELECT 
A.INID AS 编号,
C.ORDER_ID AS 订单编号
FROM INVENTORY_MST A 
LEFT JOIN PN_PRODUCTION_INSTRUCTIONS C ON A.PNID=C.PNID
LEFT JOIN PROJECT_INFO D ON C.PIID=D.PIID
LEFT JOIN CustomerInfo_MST E ON D.CUID=E.CUID WHERE A.INID='" + dt.Rows[intCurrentRowNumber]["编号"].ToString() + "'");
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        FRM.IDO = dtx.Rows[0]["编号"].ToString();
                        FRM.ShowDialog();
                    }
                }
            }
        }
        #endregion
        public void a2()
        {
            getdata = 1;
        }
        public void a3()
        {

            getdata = 2;

        }
   
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, "库存信息汇总表");
                
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            ORDER_MANAGE.INVENTORYT FRM = new INVENTORYT();
            FRM.ADD_OR_UPDATE = "ADD";
            FRM.IDO = cinventory.GETID();
            FRM.Show();  
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                IF_IMPORT_SUCCESS = false;
                bind();
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

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            try
            {
                sqb = new StringBuilder();
                sqb.AppendFormat(cinventory .sqlo);
                sqb.AppendFormat(" WHERE DateDiff(day,B.BILL_DATE,getdate()) >-1 and DateDiff(day,B.BILL_DATE,getdate()) <+20");
                sqb.Append(" GROUP BY C.ORDER_ID,E.CName,c.WNAME, C.PRODUCTION_COUNT  ORDER BY C.ORDER_ID ASC");
                dtx = bc.getdt(sqb.ToString());
                dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "订单编号");
                if (dtx.Rows.Count > 0)
                {
                    comboBox1.Items.Clear();
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

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                sqb = new StringBuilder();
                sqb.AppendFormat(cinventory.sqlt);
                sqb.AppendFormat(" WHERE C.ORDER_ID='{0}'",comboBox1.Text );
                sqb.Append(" GROUP BY A.INID,C.ORDER_ID,E.CName,c.WNAME, C.PRODUCTION_COUNT  ORDER BY C.ORDER_ID ASC"); ;
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

        private void btnToCSharp_Click(object sender, EventArgs e)
        {
   
            try
            {
                IF_IMPORT_SUCCESS = false;
                OpenFileDialog opfv = new OpenFileDialog();
                if (opfv.ShowDialog() == DialogResult.OK)
                {
                    string path = opfv.FileName;
                    cinventory.EMID = "";
                    cinventory.WAREID = "";
                    cinventory.WNAME = "";
                    cinventory.IF_IMPORT = true;
                    cinventory.showdata(path);
                    if (cinventory.IFExecution_SUCCESS)
                    {
                        IF_IMPORT_SUCCESS = true;
                        hint.Text = "导入成功";

                    }
                    bind();
                }
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            IF_IMPORT_SUCCESS = false;
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
