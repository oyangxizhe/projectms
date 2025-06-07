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
    public partial class PRINT_OPTION : Form
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
        private string _PMID;
        public string PMID
        {
            set { _PMID = value; }
            get { return _PMID; }
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
        private string _DPID;
        public string DPID
        {
            set { _DPID = value; }
            get { return _DPID; }
        }
        basec bc = new basec();
        CPRINT_OPTION cPRINT_OPTION = new CPRINT_OPTION();
        CPRINTING_MACHINE_SIZE cprinting_machine_size = new CPRINTING_MACHINE_SIZE();
        CPAPER_CORE_OPTION cpaper_core_option = new CPAPER_CORE_OPTION();
        protected int M_int_judge, i;
        protected int select;
        public PRINT_OPTION()
        {
            InitializeComponent();
        }
        private void DEPAET_Load(object sender, EventArgs e)
        {
            try
            {
                this.Icon = Resource1.xz_200X200;
                textBox3.Text = PMID;
                this.WindowState = FormWindowState.Maximized;
                if (Screen.AllScreens[0].Bounds.Width == 1366 && Screen.AllScreens[0].Bounds.Height == 768 ||
                Screen.AllScreens[0].Bounds.Width == 1280 && Screen.AllScreens[0].Bounds.Height == 800)
                {


                }
                else if (Screen.AllScreens[0].Bounds.Width == 1920 && Screen.AllScreens[0].Bounds.Height == 1080)
                {

                }
                else
                {
                    this.AutoScroll = true;
                    this.AutoScrollMinSize = new Size(1150, 768);
                }
                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void bind()
        {
           
            dt = basec.getdts(cPRINT_OPTION .sql );
            dt = cPRINT_OPTION.RETURN_HAVE_ID_DT(dt);
            if (dt.Rows.Count > 0)
            {

            }
            else
            {
                dt = total1();
            }
            dataGridView1.DataSource = dt;
            dataGridView1.ClearSelection();//加载不选中第一列
            dt2 = basec.getdts(cprinting_machine_size.sql);
            dataGridView2.DataSource = dt2;
            dt3 = basec.getdts(cpaper_core_option .sql );
            dt3 = cpaper_core_option.RETURN_HAVE_ID_DT(dt3);
            if (dt3.Rows.Count > 0)
            {

            }
            else
            {
                dt3 = total3();
            }
            dataGridView3.DataSource = dt3;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView3.AllowUserToAddRows = false;
         
            textBox4.BackColor = Color.Yellow;
       
            dgvStateControl();
            dgvStateControl_dgv2();
            dgvStateControl_dgv3();
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
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = cPRINT_OPTION.emptydatatable_T();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region total1
        private DataTable total3()
        {
            DataTable dtt2 = cpaper_core_option.emptydatatable_T();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            //this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            this.dataGridView1.ColumnHeadersHeight = 80; //标题列高度
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {
                //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
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
            
            }
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Height = 18;
            }
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView1.ClearSelection();//加载不选中第一列
            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["制单人"].ReadOnly = true;
            dataGridView1.Columns["制单日期"].ReadOnly = true;
            dataGridView1.Columns["项次"].Width = 40;
            dataGridView1.Columns["印刷选项"].Width = 60;
            dataGridView1.Columns["修边"].Width = 40;
            dataGridView1.Columns["面纸内耗1到300"].Width = 60;
            dataGridView1.Columns["面纸内耗大于300"].Width = 60;
            dataGridView1.Columns["底纸内耗1到300"].Width = 70;
            dataGridView1.Columns["底纸内耗大于300"].Width = 60;
            dataGridView1.Columns["无印刷用纸表面处理损耗_固定值"].Width = 60;
            dataGridView1.Columns["无印刷用纸表面处理损耗_百分比"].Width = 60;
            dataGridView1.Columns["正面印刷纸张损耗_A"].Width = 60;
            dataGridView1.Columns["正面印刷纸张损耗_B"].Width = 60;
            dataGridView1.Columns["正面印刷纸张损耗_C"].Width = 60;
            dataGridView1.Columns["正面印刷纸张损耗_D"].Width = 60;
            dataGridView1.Columns["正面印刷纸张损耗_E"].Width = 80;
            dataGridView1.Columns["正面印刷纸张损耗_F"].Width = 70;
            dataGridView1.Columns["正面印刷纸张损耗_G"].Width = 70;
            dataGridView1.Columns["正面印刷纸张损耗_H"].Width = 70;
            dataGridView1.Columns["正面印刷纸张损耗_I"].Width = 70;
            dataGridView1.Columns["正面印刷纸张损耗_J"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_A"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_B"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_C"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_D"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_E"].Width = 80;
            dataGridView1.Columns["反面印刷纸张损耗_F"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_G"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_H"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_I"].Width = 70;
            dataGridView1.Columns["反面印刷纸张损耗_J"].Width = 70;
         
            dataGridView1.Columns["面纸内耗1到300"].HeaderText = "面纸内耗1-300";
            dataGridView1.Columns["面纸内耗大于300"].HeaderText = "面纸内耗大于300百分比%";
            dataGridView1.Columns["底纸内耗1到300"].HeaderText = "底纸内耗1-300";
            dataGridView1.Columns["底纸内耗大于300"].HeaderText = "底纸内耗大于300百分比%";

            dataGridView1.Columns["无印刷用纸表面处理损耗_固定值"].HeaderText = "无印刷表面处理损耗固定值";
            dataGridView1.Columns["无印刷用纸表面处理损耗_百分比"].HeaderText = "无印刷表面处理损耗百分比";
            dataGridView1.Columns["正面印刷纸张损耗_A"].HeaderText = "正印纸耗<=4<=3000保底数";
            dataGridView1.Columns["正面印刷纸张损耗_B"].HeaderText = "正印纸耗<=4<=3000单色数";
            dataGridView1.Columns["正面印刷纸张损耗_C"].HeaderText = "正印纸耗<=4>3000保底数";
            dataGridView1.Columns["正面印刷纸张损耗_D"].HeaderText = "正印纸耗<=4>3000单色数";
            dataGridView1.Columns["正面印刷纸张损耗_E"].HeaderText = "正印纸耗<=4>3000超出百分比";

            dataGridView1.Columns["正面印刷纸张损耗_F"].HeaderText = "正印纸耗>4<=3000保底数";
            dataGridView1.Columns["正面印刷纸张损耗_G"].HeaderText = "正印纸耗>4<=3000单色数";
            dataGridView1.Columns["正面印刷纸张损耗_H"].HeaderText = "正印纸耗>4>3000保底数";
            dataGridView1.Columns["正面印刷纸张损耗_I"].HeaderText = "正印纸耗>4>3000单色数";
            dataGridView1.Columns["正面印刷纸张损耗_J"].HeaderText = "正印纸耗>4>3000超出百分比";

            dataGridView1.Columns["反面印刷纸张损耗_A"].HeaderText = "反印纸耗<=4<=3000保底数";
            dataGridView1.Columns["反面印刷纸张损耗_B"].HeaderText = "反印纸耗<=4<=3000单色数";
            dataGridView1.Columns["反面印刷纸张损耗_C"].HeaderText = "反印纸耗<=4>3000保底数";
            dataGridView1.Columns["反面印刷纸张损耗_D"].HeaderText = "反印纸耗<=4>3000单色数";
            dataGridView1.Columns["反面印刷纸张损耗_E"].HeaderText = "反印纸耗<=4>3000超出百分比";

            dataGridView1.Columns["反面印刷纸张损耗_F"].HeaderText = "反印纸耗>4<=3000保底数";
            dataGridView1.Columns["反面印刷纸张损耗_G"].HeaderText = "反印纸耗>4<=3000单色数";
            dataGridView1.Columns["反面印刷纸张损耗_H"].HeaderText = "反印纸耗>4>3000保底数";
            dataGridView1.Columns["反面印刷纸张损耗_I"].HeaderText = "反印纸耗>4>3000单色数";
            dataGridView1.Columns["反面印刷纸张损耗_J"].HeaderText = "反印纸耗>4>3000超出百分比";
            dataGridView1.Columns["制单人"].Width = 70;
            dataGridView1.Columns["制单日期"].Width = 120;
        }
        #endregion
        #region dgvStateControl_dgv2
        private void dgvStateControl_dgv2()
        {
            int i;
            dataGridView2.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView2.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
                dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView2.EnableHeadersVisualStyles = false;
                dataGridView2.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView2.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            for (i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView2.Columns[i].ReadOnly = true;

            }
        }
        #endregion
        #region dgvStateControl_dgv3
        private void dgvStateControl_dgv3()
        {
            int i;
            dataGridView3.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView3.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {
                //dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView3.EnableHeadersVisualStyles = false;
                dataGridView3.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView3.Columns.Count; i++)
            {
                dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView3.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            for (i = 0; i < dataGridView3.Columns.Count; i++)
            {
                dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            

            }
            dataGridView3.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView3.ClearSelection();//加载不选中第一列
            dataGridView3.Columns["项次"].ReadOnly = true;
            dataGridView3.Columns["制单人"].ReadOnly = true;
            dataGridView3.Columns["制单日期"].ReadOnly = true;
            for (i = 0; i < dataGridView3.Rows.Count; i++)
            {
                dataGridView3.Rows[i].Height = 18;
            }
            dataGridView3.Columns["项次"].Width = 40;
            dataGridView3.Columns["芯纸选项"].Width = 60;
            dataGridView3.Columns["芯纸内耗1到300"].Width = 60;
            dataGridView3.Columns["芯纸内耗大于300"].Width = 60;
            dataGridView3.Columns["芯纸内耗1到300"].HeaderText = "芯纸内耗1-300";
            dataGridView3.Columns["芯纸内耗大于300"].HeaderText = "芯纸内耗大于300百分比%";
            dataGridView3.Columns["制单人"].Width = 70;
            dataGridView3.Columns["制单日期"].Width = 120;
          
        }
        #endregion
      
        #region save
        private void save()
        {
            if (dt.Rows.Count > 0)
            {
                cPRINT_OPTION.MAKERID = LOGIN.EMID;
                if (LOGIN.EMID != null)
                {
                    dt = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "印刷选项 IS NOT NULL");
                    cPRINT_OPTION.save(dt);
                }
             
            }
            else
            {

            }
       
        }
        #endregion
        #region juage()
        private bool juage()
        {
            bool b = false;
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "", "印刷选项 IS NOT NULL");
            for(i=0;i<dtx.Rows .Count ;i++)
            {
                if (b==true)
                    break;
                for (int j = 2; j < dtx.Columns.Count-2; j++)
                {
                    if (j!=7 && (dtx.Rows[i][j].ToString() != "" && bc.yesno(dtx.Rows[i][j].ToString()) == 0))
                    {
                        b = true;
                        hint.Text = string.Format("项次" + "{0}" + " 存在不为数值的栏位！" + "位置 第{1}行" + "," + "第{2}列", dtx.Rows[i]["项次"].ToString(), i + 1, j + 1);
                        break;
                    }
                }

            }
            return b;
        }
        #endregion
        #region juage_dgv2()
        private bool juage_dgv2()
        {
            bool b = false;
            if (textBox3.Text == "")
            {
                b = true;
                hint.Text = "编号不能为空！";
            }
            else if (textBox4.Text == "")
            {
                b = true;
                hint.Text = "印刷机型不能为空！";
            }
            else if (textBox5.Text =="" || bc.yesno(textBox5.Text) == 0)
            {
                b = true;
                hint.Text = "最大宽不能为空且只能输入数字！";
            }
            else if (textBox6.Text == "" || bc.yesno(textBox6.Text) == 0)
            {
                b = true;
                hint.Text = "最大长不能为空且只能输入数字！";
            }
            else if (textBox7.Text == "" || bc.yesno(textBox7.Text) == 0)
            {
                b = true;
                hint.Text = "最小宽不能为空且只能输入数字！";
            }
            else if (textBox8.Text == "" || bc.yesno(textBox8.Text) == 0)
            {
                b = true;
                hint.Text = "最小长不能为空且只能输入数字！";
            }
            return b;
        }
        #endregion
        #region juage_dgv3()
        private bool juage_dgv3()
        {
            bool b = false;
            DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt3, "", "芯纸选项 IS NOT NULL");
            for (i = 0; i < dtx.Rows.Count; i++)
            {
                if (b == true)
                    break;
                for (int j = 2; j < dtx.Columns.Count -2; j++)
                {
                    if ( (dtx.Rows[i][j].ToString() != "" && bc.yesno(dtx.Rows[i][j].ToString()) == 0))
                    {
                        b = true;
                        hint.Text = string.Format("项次" + "{0}" + " 存在不为数值的栏位！" + "位置 第{1}行" + "," + "第{2}列", 
                            dtx.Rows[i]["项次"].ToString(), i + 1, j + 1);
                        break;
                    }
                }

            }
            return b;
        }
        #endregion
        public void ClearText_dgv2()
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
        }
        public void ClearText_dgv3()
        {
         
        }
        private void add()
        {
         
            bind();
        }
        private void add_dgv2()
        {
            ClearText_dgv2();
            textBox3.Text = cprinting_machine_size.GETID();
            textBox4.Focus();
            bind();
        }
        private void add_dgv3()
        {
          
            bind();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
         
            btnSave.Focus();
            if (juage())
            {

            }
            else
            {
                save();
                IFExecution_SUCCESS =cPRINT_OPTION .IFExecution_SUCCESS ;
                hint.Text = cPRINT_OPTION.ErrowInfo;
                if (IFExecution_SUCCESS)
                {
                  
                    add();
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {


                dt = bc.getdt(cPRINT_OPTION .sql +" WHERE  A.PRINT_OPTION LIKE '%%'");
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
        
            try
            {
                string id = Convert.ToString(dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM PRINT_OPTION WHERE PRINT_OPTION='" + id + "'";
                basec.getcoms(strSql);
                add();
            }
            catch (Exception)
            {


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

                dataGridView1.Focus();

                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
           /* string v1 = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
            if (v1 != "")
            {
                textBox1.Text = bc.getOnlyString(string.Format("SELECT POID FROM PRINT_OPTION WHERE PRINT_OPTION='{0}'",v1 ));
                textBox2.Text = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                textBox13.Text = dt.Rows[dataGridView1.CurrentCell.RowIndex]["修边"].ToString();
            }*/
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
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


        private void pictureBox1_Click(object sender, EventArgs e)
        {
            add_dgv2();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (juage_dgv2())
            {

            }
            else
            {
                save_dgv2();
                IFExecution_SUCCESS = cprinting_machine_size.IFExecution_SUCCESS;
                hint.Text = cprinting_machine_size.ErrowInfo;
                if (IFExecution_SUCCESS)
                {
                    add_dgv2();
                }
            }
        }
        #region save_dgv2
        private void save_dgv2()
        {
            cprinting_machine_size.PMID  = textBox3.Text;
            cprinting_machine_size.MACHINE_TYPE  = textBox4.Text;
            cprinting_machine_size.MAX_WIDTH  = textBox5.Text;
            cprinting_machine_size.MAX_LENGTH = textBox6.Text;
            cprinting_machine_size.MIN_WIDTH  = textBox7.Text;
            cprinting_machine_size.MIN_LENGTH  = textBox8.Text;
            cprinting_machine_size.PRINTING_PAPER  = textBox9.Text;
            cprinting_machine_size.MAKERID = LOGIN.EMID;
            cprinting_machine_size.save();
        }
        #endregion
        #region save_dgv3
        private void save_dgv3()
        {
            if (dt3.Rows.Count > 0)
            {
                cpaper_core_option.MAKERID = LOGIN.EMID;
                dt3 = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt3, "", "芯纸选项 IS NOT NULL");
                cpaper_core_option.save(dt3);
            }
            else
            {

            }
      
        }
        #endregion
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string id = Convert.ToString(dataGridView2[0, dataGridView2.CurrentCell.RowIndex].Value).Trim();
            try
            {
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM PRINTING_MACHINE_SIZE WHERE MACHINE_TYPE='" + id + "'";
                basec.getcoms(strSql);
                add_dgv2();
            }
            catch (Exception)
            {


            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string v1 = Convert.ToString(dataGridView2[0, dataGridView2.CurrentCell.RowIndex].Value).Trim();
            textBox3.Text = bc.getOnlyString(string.Format("SELECT PMID FROM PRINTING_MACHINE_SIZE WHERE MACHINE_TYPE='{0}'", v1));
            textBox4.Text = Convert.ToString(dataGridView2[0, dataGridView2.CurrentCell.RowIndex].Value).Trim();
            textBox5.Text = dt2.Rows[dataGridView2.CurrentCell.RowIndex]["最大宽"].ToString();
            textBox6.Text = dt2.Rows[dataGridView2.CurrentCell.RowIndex]["最大长"].ToString();
            textBox7.Text = dt2.Rows[dataGridView2.CurrentCell.RowIndex]["最小宽"].ToString();
            textBox8.Text = dt2.Rows[dataGridView2.CurrentCell.RowIndex]["最小长"].ToString();
            textBox9.Text = dt2.Rows[dataGridView2.CurrentCell.RowIndex]["印刷用纸"].ToString();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            add_dgv3();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            pictureBox5.Focus ();
            if (juage_dgv3())
            {

            }
            else
            {
                save_dgv3();
                IFExecution_SUCCESS = cpaper_core_option.IFExecution_SUCCESS;
                hint.Text = cpaper_core_option.ErrowInfo;
                if (IFExecution_SUCCESS)
                {
                    add();
                }
            }
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            string id = Convert.ToString(dataGridView3[1, dataGridView3.CurrentCell.RowIndex].Value).Trim();
            try
            {
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM PAPER_CORE_OPTION WHERE PAPER_CORE='" + id + "'";
                basec.getcoms(strSql);
                add_dgv3();
            }
            catch (Exception)
            {


            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SHOW_IMAGE show_image = new SHOW_IMAGE();
            show_image.Show();
        }

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
                    DataRow dr = dt.NewRow();
                    int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                    dr["项次"] = Convert.ToString(b1 + 1);
                    dt.Rows.Add(dr);
                    dgvStateControl(); 
                }
                
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

            //dgvfoucs();

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, this.Text);

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView3_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int a = dataGridView3.CurrentCell.ColumnIndex;
                int b = dataGridView3.CurrentCell.RowIndex;
                int c = dataGridView3.Columns.Count - 1;
                int d = dataGridView3.Rows.Count - 1;
                if (a == c && b == d)
                {
                    DataRow dr = dt3.NewRow();
                    int b1 = Convert.ToInt32(dt3.Rows[dt3.Rows.Count - 1]["项次"].ToString());
                    dr["项次"] = Convert.ToString(b1 + 1);
                    dt3.Rows.Add(dr);
                    dgvStateControl_dgv3();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

     
    }
}
