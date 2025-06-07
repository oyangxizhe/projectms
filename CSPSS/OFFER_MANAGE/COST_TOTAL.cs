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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSPSS.OFFER_MANAGE
{
    public partial class COST_TOTAL : Form
    {

    
        #region nature
        basec bc = new basec();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private bool _IF_COMPLETED;
        public bool IF_COMPLETED
        {
            set { _IF_COMPLETED = value; }
            get { return _IF_COMPLETED; }
        }
        private static string _EMID;
        public static string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private static string _ENAME;
        public static string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        private int _GET_DATA_INT;
        public int GET_DATA_INT
        {
            set { _GET_DATA_INT = value; }
            get { return _GET_DATA_INT; }

        }
        private bool _LOAD_OR_SEARCH;
        public bool LOAD_OR_SEARCH
        {
            set { _LOAD_OR_SEARCH = value; }
            get { return _LOAD_OR_SEARCH; }

        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private string _OFFER_ID;
        public string OFFER_ID
        {
            set { _OFFER_ID = value; }
            get { return _OFFER_ID; }
        }
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private bool _IF_CHECKBOX;
        public bool IF_CHECKBOX
        {
            set { _IF_CHECKBOX = value; }
            get { return _IF_CHECKBOX; }

        }
        private static bool _IF_DOUBLE_CLICK;
        public static bool IF_DOUBLE_CLICK
        {
            set { _IF_DOUBLE_CLICK = value; }
            get { return _IF_DOUBLE_CLICK; }

        }

        #endregion
        protected int M_int_judge, i;
        protected int select;

        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        COTHER_COST cother_cost = new COTHER_COST();
        CPRINT_COST_TOTAL cprint_cost_total = new CPRINT_COST_TOTAL();
        CPRINT_DIE_CUTTING cprint_die_cutting = new CPRINT_DIE_CUTTING();
        CPRINT_PORTRAY cprint_portray = new CPRINT_PORTRAY();
        CPRINT_PURCHASE cprint_purchase = new CPRINT_PURCHASE();
        CPRINT_TRANSPORT cprint_transport = new CPRINT_TRANSPORT();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable dt4 = new DataTable();
        DataTable dt5 = new DataTable();
        DataTable dt6 = new DataTable();
        DataTable dt7 = new DataTable();
        DataTable dt8 = new DataTable();
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dtt = new DataTable();
        LOADING frm_loading = new LOADING();
        PRINTING_OFFER F2 = new PRINTING_OFFER();
        public COST_TOTAL()
        {
            InitializeComponent();
        }
        public COST_TOTAL(PRINTING_OFFER FRM2)
        {
          
            F2 = FRM2;
            InitializeComponent();
        }
        private void loading(object sender, DoWorkEventArgs e)
        {
            for (i = 0; i < 6000; i++)
            {
                //MessageBox.Show(i.ToString ());
                System.Threading.Thread.Sleep(100);
                if (IF_COMPLETED)
                {
                    groupBox1.Visible = true;
                    groupBox1.Visible = true;
                    groupBox3.Visible = true;
                    groupBox11.Visible = true;
                    groupBox12.Visible = true;
                    frm_loading.Close();
                    break;
                }
            }
        }
        private void loading()
        {
   
                frm_loading.Show();
                IF_COMPLETED = false;
                groupBox1.Visible = false;
                groupBox3.Visible = false;
                groupBox11.Visible = false;
                groupBox12.Visible = false;
                /*同一个BackgroundWorker对象要实例化两次才不使用第二次点击该作业时不执持 16/01/21 start*/
                BackgroundWorker work = new BackgroundWorker();
                work.RunWorkerAsync();
                work.WorkerReportsProgress = true;
                work.DoWork += new DoWorkEventHandler(loading);
                BackgroundWorker work1 = new BackgroundWorker();
                work1.RunWorkerAsync();
                work1.WorkerReportsProgress = true;
                work1.DoWork += new DoWorkEventHandler(loading);
                /*同一个BackgroundWorker对象要实例化两次才不使用第二次点击该作业时不执持 16/01/21 end*/
            
        }
        private void COST_TOTAL_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;//避免出现线程间操作无效: 从不是创建控件“progressBar1”的线程访问它 160120
            //loading();
            bind();
            IF_COMPLETED = true;
          
        }
        private void bind()
        {
            hint.ForeColor = Color.Red;
            if (Screen.AllScreens[0].Bounds.Width == 1366 && Screen.AllScreens[0].Bounds.Height == 768 ||
               Screen.AllScreens[0].Bounds.Width == 1280 && Screen.AllScreens[0].Bounds.Height == 800)
            {

                this.AutoScroll = true;
                this.AutoScrollMinSize = new Size(1000, 700);
                groupBox12.Height = 223;
                dataGridView9.Height = 203;

                groupBox11.Height = 165;
                dataGridView1.Height = 141;
                groupBox3.Location = new Point(3, 235);

            }
            else if (Screen.AllScreens[0].Bounds.Width == 1920 && Screen.AllScreens[0].Bounds.Height == 1080)
            {

            }
            else
            {
                this.AutoScroll = true;
                this.AutoScrollMinSize = new Size(1920, 1080);
            }
            hint.Text = "";
            label1.Text = "";
            button1.BackColor = CCOLOR.lylfnp;
            button1.ForeColor = Color.White;
            try
            {
                textBox1.ScrollBars = ScrollBars.Both;
              this.Icon = Resource1.xz_200X200;
                //OFFER_ID = "1601Z004-02-ADM-A";
                //PFID =bc.getOnlyString("SELECT PFID FROM PRINTING_OFFER_MST WHERE OFFER_ID in ('1601Z004-02-ADM','1601Z004-02-ADM-A')");
                dtt = bc.getdt(cprinting_offer.sqlni  + " WHERE  A.PFID='" + PFID  + "'");
                dt = cprinting_offer.RETURN_DT_SHOW_HIDE_FORM(dtt);
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dgvStateControl();
                }
              
                label1.Text = "报价编号：" + OFFER_ID;
                label1.Font = new Font("", 16);
                label1.BackColor = CCOLOR.lylfnp;
                label1.ForeColor = Color.White;
                dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.PFID='" + PFID + "'");
                if (dtx.Rows.Count > 0)
                {

                    dtx = cprint_cost_total.RETURN_DT(dtx);
                    dtx = cprinting_offer.RETURN_COST_TOTAL_DT_FORM(dtx);
                }
                //MessageBox.Show("1");
                if (dtx.Rows.Count > 0)
                {
                    dataGridView9.DataSource = dtx;
                    dgvStateControl_dgv9();
                }
                dt2 = bc.getdt(cprint_die_cutting.sql + " WHERE A.PFID='" + PFID + "'");
                if (dt2.Rows.Count > 0)
                {
                    dataGridView2.DataSource = dt2;
                    dgvStateControl_dgv2();
                }
                dt3 = bc.getdt(cprint_portray.sql  + " WHERE A.PFID='" + PFID + "'");/*dgv3 start*/
                dt3=cprint_portray.RETURN_DT(dt3);
                if (dt3.Rows.Count > 0)
                {
                    dataGridView3.DataSource = dt3;
                    dgvStateControl_dgv3();
               
                }
                //MessageBox.Show("2");
                dt4 = cprinting_offer.RETURN_PARTS_AUXILIAR_DT(PFID, dtt);/*dgv4 start*/
                if (dt4.Rows.Count > 0)
                {
                    dataGridView4.DataSource = dt4;
                    dgvStateControl_dgv4();
                }
                //MessageBox.Show("4");
                dt5 = cprinting_offer.RETURN_PACK_MATERIAL_DT(PFID, dtt);/*dgv5 start*/
                if (dt5.Rows.Count > 0)
                {
                    dataGridView5.DataSource = dt5;
                    dgvStateControl_dgv5();
                }
                //MessageBox.Show("5");
                dt6 = cprinting_offer.RETURN_ARTIFICIAL_DT(PFID, dtt);/*dgv6 start*/
                if (dt6.Rows.Count > 0)
                {
                    dataGridView6.DataSource = dt6;
                    dgvStateControl_dgv6();
                }
               // MessageBox.Show("3");
                dt7 = bc.getdt(cprint_purchase.sql + " WHERE A.PFID='" + PFID + "'");/*dgv7 start*/
                dt7 = cprint_purchase.RETURN_HAVE_ID_DT(dt7);
                if (dt7.Rows.Count > 0)
                {
                    dataGridView7.DataSource = dt7;
                    dgvStateControl_dgv7();
                }
                dt8 = bc.getdt(cprint_transport.sql + " WHERE A.PFID='" + PFID + "'");/*dgv8 start*/
                if (dt8.Rows.Count > 0)
                {
                    dataGridView8.DataSource = dt8;
                    dgvStateControl_dgv8();
                }
               
          
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #region dgvStateControl_dgv2
        private void dgvStateControl_dgv2()
        {
            int i;
            dataGridView2.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView2.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView2.ClearSelection();
            dataGridView2.AllowUserToAddRows = false;
            int numCols1 = dataGridView2.Columns.Count;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            // dataGridView2.Columns["项次"].Width = 40;

            for (i = 0; i < numCols1; i++)
            {

                dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView2.EnableHeadersVisualStyles = false;
                dataGridView2.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView2.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView2.Columns[i].ReadOnly = true;

            }
            //dataGridView2.Columns["站别代码"].DefaultCellStyle.BackColor = Color.Yellow;


            for (i = 0; i < dataGridView2.Rows.Count; i++)
            {

                dataGridView2.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView2.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView2.Columns["项目"].Width = 70;
            dataGridView2.Columns["刀模长米"].Width = 40;
            dataGridView2.Columns["元米"].Width = 50;
            dataGridView2.Columns["圆孔个数"].Width = 40;
            dataGridView2.Columns["元个"].Width = 50;
            dataGridView2.Columns["小计"].Width = 50;
            dataGridView2.Columns["元米"].HeaderText = "元/米";
            dataGridView2.Columns["元个"].HeaderText = "元/个";
        }
        #endregion
        #region dgvStateControl_dgv3
        private void dgvStateControl_dgv3()
        {
            int i;
            dataGridView3.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView3.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView3.ClearSelection();
            dataGridView3.AllowUserToAddRows = false;
            int numCols1 = dataGridView3.Columns.Count;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            dataGridView3.Columns["写真类型"].Width = 70;
            dataGridView3.Columns["长"].Width = 40;
            dataGridView3.Columns["宽"].Width = 40;
            dataGridView3.Columns["总数量"].Width = 50;
            dataGridView3.Columns["单价"].Width = 40;
            dataGridView3.Columns["小计"].Width = 60;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView3.EnableHeadersVisualStyles = false;
                dataGridView3.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView3.Columns[i].ReadOnly = true;

            }
            //dataGridView2.Columns["站别代码"].DefaultCellStyle.BackColor = Color.Yellow;


            for (i = 0; i < dataGridView3.Rows.Count; i++)
            {

                dataGridView3.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                dataGridView3.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView3.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView3.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
        }
        #endregion
        #region dgvStateControl_dgv4
        private void dgvStateControl_dgv4()
        {
            int i;
            dataGridView4.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView4.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView4.ClearSelection();
            dataGridView4.AllowUserToAddRows = false;
            int numCols1 = dataGridView4.Columns.Count;
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
         
      
            for (i = 0; i < numCols1; i++)
            {

                dataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView4.EnableHeadersVisualStyles = false;
                dataGridView4.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView4.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView4.Columns[i].ReadOnly = true;

            }
            //dataGridView2.Columns["站别代码"].DefaultCellStyle.BackColor = Color.Yellow;


            for (i = 0; i < dataGridView4.Rows.Count; i++)
            {

                dataGridView4.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView4.Rows.Count - 1; i++)
            {
                dataGridView4.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView4.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView4.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
        }
        #endregion
        #region dgvStateControl_dgv5
        private void dgvStateControl_dgv5()
        {
            int i;
            dataGridView5.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView5.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView5.ClearSelection();
            dataGridView5.AllowUserToAddRows = false;
            int numCols1 = dataGridView5.Columns.Count;


            for (i = 0; i < numCols1; i++)
            {
                dataGridView5.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView5.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

                dataGridView5.EnableHeadersVisualStyles = false;
                dataGridView5.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView5.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView5.Columns[i].ReadOnly = true;
            }
            for (i = 0; i < dataGridView5.Rows.Count; i++)
            {

                dataGridView5.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView5.Rows.Count - 1; i++)
            {
                dataGridView5.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView5.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView5.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView5.Columns["项目"].Width = 60;
            dataGridView5.Columns["数量"].Width = 40;
            dataGridView5.Columns["长"].Width = 40;
            dataGridView5.Columns["宽"].Width = 40;
            dataGridView5.Columns["高"].Width = 40;
            dataGridView5.Columns["箱形"].Width = 100;
            dataGridView5.Columns["材质"].Width = 70;
            dataGridView5.Columns["单价"].Width = 50;
            dataGridView5.Columns["小计"].Width = 50;
        }
        #endregion
        #region dgvStateControl_dgv6
        private void dgvStateControl_dgv6()
        {
            int i;
            dataGridView6.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView6.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView6.ClearSelection();
            dataGridView6.AllowUserToAddRows = false;
            int numCols1 = dataGridView6.Columns.Count;
            dataGridView6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/

            for (i = 0; i < numCols1; i++)
            {
                dataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView6.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView6.EnableHeadersVisualStyles = false;
                dataGridView6.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView6.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView6.Columns[i].ReadOnly = true;
            }
            //dataGridView2.Columns["站别代码"].DefaultCellStyle.BackColor = Color.Yellow;


            for (i = 0; i < dataGridView6.Rows.Count; i++)
            {

                dataGridView6.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
            {
                dataGridView6.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView6.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView6.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            //dataGridView6.Columns["序号"].Width = 40;
            dataGridView6.Columns["项目"].Width = 70;
            dataGridView6.Columns["数量"].Width = 40;
            dataGridView6.Columns["单价"].Width = 50;
            dataGridView6.Columns["元套"].Width = 40;
            dataGridView6.Columns["小计"].Width = 50;
        }
        #endregion
        #region dgvStateControl_dgv7
        private void dgvStateControl_dgv7()
        {
            int i;
            dataGridView7.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView7.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView7.ClearSelection();
            dataGridView7.AllowUserToAddRows = false;
            int numCols1 = dataGridView7.Columns.Count;
            //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/

            for (i = 0; i < numCols1; i++)
            {
                dataGridView7.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView7.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView7.EnableHeadersVisualStyles = false;
                dataGridView7.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView7.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView7.Columns[i].ReadOnly = true;
            }
            for (i = 0; i < dataGridView7.Rows.Count; i++)
            {
                dataGridView7.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView7.Rows.Count - 1; i++)
            {
                dataGridView7.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView7.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView7.Columns["类型一"].Width = 70;
            dataGridView7.Columns["类型二"].Width = 70;
            dataGridView7.Columns["外购价一"].Width = 50;
            dataGridView7.Columns["外购价二"].Width = 50;
            dataGridView7.Columns["管理费一"].Width = 50;
            dataGridView7.Columns["小计一"].Width = 50;
            dataGridView7.Columns["管理费二"].Width = 50;
            dataGridView7.Columns["小计二"].Width = 50;

            dataGridView7.Columns["类型一"].HeaderText = "类型";
            dataGridView7.Columns["类型二"].HeaderText = "类型";
            dataGridView7.Columns["外购价一"].HeaderText = "外购价";
            dataGridView7.Columns["外购价二"].HeaderText = "外购价";
            dataGridView7.Columns["管理费一"].HeaderText = "管理费";
            dataGridView7.Columns["小计一"].HeaderText = "小计";
            dataGridView7.Columns["管理费二"].HeaderText = "管理费";
            dataGridView7.Columns["小计二"].HeaderText = "小计";
        }
        #endregion
        #region dgvStateControl_dgv8
        private void dgvStateControl_dgv8()
        {
            int i;
            dataGridView8.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView8.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView8.ClearSelection();
            dataGridView8.AllowUserToAddRows = false;
            int numCols1 = dataGridView8.Columns.Count;
            dataGridView8.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/

            for (i = 0; i < numCols1; i++)
            {
                dataGridView8.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView8.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView8.EnableHeadersVisualStyles = false;
                dataGridView8.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView8.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dataGridView8.Columns[i].ReadOnly = true;
            }
            for (i = 0; i < dataGridView8.Rows.Count; i++)
            {

                dataGridView8.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView8.Rows.Count - 1; i++)
            {
                dataGridView8.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView8.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            /*dataGridView8.Columns["长"].Width = 40;
            dataGridView8.Columns["宽"].Width = 40;
            dataGridView8.Columns["高"].Width = 40;
            dataGridView8.Columns["单价"].Width = 50;
            dataGridView8.Columns["总箱数"].Width = 50;
            dataGridView8.Columns["总立方数"].Width = 70;
            dataGridView8.Columns["运输方式"].Width = 120;
            dataGridView8.Columns["小计"].Width = 50;*/
        }
        #endregion
        private void btnSearch_Click(object sender, EventArgs e)
        {


        }


        private void btnAdd_Click(object sender, EventArgs e)
        {


        }
        private void btndgvInfoCopy_Click(object sender, EventArgs e)
        {

            dgvCopy(ref dataGridView9);
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
        #region ProcessCmdKey
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (ActiveControl.TabIndex == 0)
            {

            }
            else
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
            }
            return base.ProcessCmdKey(ref msg, keyData);

        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            dataGridView1.ClearSelection();//取消默认选中行
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView1.Columns["序号"].Width = 40;
            int numCols1 = dataGridView1.Columns.Count;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                if (dataGridView1.Columns[i].DataPropertyName == "部品名" || dataGridView1.Columns[i].DataPropertyName == "机器型号" ||
                    dataGridView1.Columns[i].DataPropertyName == "项目名称" || dataGridView1.Columns[i].DataPropertyName == "数量" ||
                    dataGridView1.Columns[i].DataPropertyName == "项目号" || dataGridView1.Columns[i].DataPropertyName == "报价编号" ||
                    dataGridView1.Columns[i].DataPropertyName == "报价" || dataGridView1.Columns[i].DataPropertyName == "日期" ||
                    dataGridView1.Columns[i].DataPropertyName == "序号")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                else
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }
                dataGridView1.Columns[i].ReadOnly = true;


            }
            dataGridView1.Columns["印刷选项"].Visible = false;
            dataGridView1.Columns["模切"].Visible = false;
            dataGridView1.Columns["报价数量"].Visible = false;
            dataGridView1.Columns["报出价"].Visible = false;
            dataGridView1.Columns["审核批注"].Visible = false;
            dataGridView1.Columns["打样单号"].Visible = false;
            dataGridView1.Columns["打样金额"].Visible = false;
            dataGridView1.Columns["AE"].Visible = false;
            dataGridView1.Columns["报价ID"].Visible = false;
            dataGridView1.Columns["客户"].Visible = false;
            dataGridView1.Columns["品牌"].Visible = false;
            dataGridView1.Columns["项目名称"].Visible = false;
            dataGridView1.Columns["数量"].Visible = false;
            dataGridView1.Columns["项目号"].Visible = false;
            dataGridView1.Columns["报价编号"].Visible = false;
            dataGridView1.Columns["报价"].Visible = false;
            dataGridView1.Columns["日期"].Visible = false;
            dataGridView1.Columns["序号"].Visible = false;
            dataGridView1.Columns["项目名称"].Visible = false;
            dataGridView1.Columns["数量"].Visible = false;

            dataGridView1.Columns["序号"].Visible = false;
            dataGridView1.Columns["项目名称"].Visible = false;
            dataGridView1.Columns["数量"].Visible = false;
            dataGridView1.Columns["项目号"].Visible = false;
            dataGridView1.Columns["报价编号"].Visible = false;
            dataGridView1.Columns["报价"].Visible = false;
            dataGridView1.Columns["日期"].Visible = false;
            dataGridView1.Columns["部品名"].Visible = false;
            dataGridView1.Columns["加工门幅"].Visible = false;
            dataGridView1.Columns["加工长度"].Visible = false;
            dataGridView1.Columns["部品总数"].Visible = false;
            dataGridView1.Columns["机器型号"].Visible = false;
            dataGridView1.Columns["部品单价"].Visible = false;
            dataGridView1.Columns["部品总价"].Visible = false;
            dataGridView1.Columns["面纸单价"].Visible = false;
            dataGridView1.Columns["面纸用量"].Visible = false;
            dataGridView1.Columns["面纸内耗"].Visible = false;
            dataGridView1.Columns["面纸下单"].Visible = false;
            dataGridView1.Columns["面纸外耗"].Visible = false;
            dataGridView1.Columns["面纸门幅"].Visible = false;
            dataGridView1.Columns["面纸纸长"].Visible = false;
            dataGridView1.Columns["面纸可用"].Visible = false;
            dataGridView1.Columns["面纸单个用量"].Visible = false;
            dataGridView1.Columns["面纸小计"].Visible = false;
            dataGridView1.Columns["芯纸单价"].Visible = false;
            dataGridView1.Columns["芯纸内耗"].Visible = false;
            dataGridView1.Columns["芯纸用量"].Visible = false;
            dataGridView1.Columns["芯纸门幅"].Visible = false;
            dataGridView1.Columns["芯纸纸长"].Visible = false;
            dataGridView1.Columns["芯纸可用"].Visible = false;
            dataGridView1.Columns["芯纸单个用量"].Visible = false;
            dataGridView1.Columns["芯纸小计"].Visible = false;
            dataGridView1.Columns["底纸单价"].Visible = false;
            dataGridView1.Columns["底纸用量"].Visible = false;
            dataGridView1.Columns["底纸内耗"].Visible = false;
            dataGridView1.Columns["底纸下单"].Visible = false;
            dataGridView1.Columns["底纸外耗"].Visible = false;
            dataGridView1.Columns["底纸单个用量"].Visible = false;
            dataGridView1.Columns["底纸小计"].Visible = false;
            dataGridView1.Columns["印工单色单价"].Visible = false;
            dataGridView1.Columns["超出单色单张价"].Visible = false;
            dataGridView1.Columns["CTP单张价"].Visible = false;
            dataGridView1.Columns["正面色数共计"].Visible = false;
            dataGridView1.Columns["正面CTP张数"].Visible = false;
            dataGridView1.Columns["正面纸张损耗"].Visible = false;
            dataGridView1.Columns["正面防晒合计"].Visible = false;
            dataGridView1.Columns["正面CTP价计"].Visible = false;
            dataGridView1.Columns["正面印工合计"].Visible = false;
            dataGridView1.Columns["反面色数共计"].Visible = false;
            dataGridView1.Columns["反面CTP张数"].Visible = false;
            dataGridView1.Columns["反面纸张损耗"].Visible = false;
            dataGridView1.Columns["反面防晒合计"].Visible = false;
            dataGridView1.Columns["反面CTP价计"].Visible = false;
            dataGridView1.Columns["反面印工合计"].Visible = false;
            dataGridView1.Columns["正反CTP合计"].Visible = false;
            dataGridView1.Columns["正反印工合计"].Visible = false;
            dataGridView1.Columns["表面处理单价"].Visible = false;
            dataGridView1.Columns["无印刷表面处理损耗"].Visible = false;
            dataGridView1.Columns["表面处理用量"].Visible = false;
            dataGridView1.Columns["表面加工小计"].Visible = false;
            dataGridView1.Columns["裱工单价"].Visible = false;
            dataGridView1.Columns["裱工用量"].Visible = false;
            dataGridView1.Columns["裱工小计"].Visible = false;
            dataGridView1.Columns["刀模小计"].Visible = false;
            dataGridView1.Columns["模切小计"].Visible = false;

            dataGridView1.Columns["面纸"].Visible = false;
            dataGridView1.Columns["面纸克重"].Visible = false;
            dataGridView1.Columns["芯纸"].Visible = false;
            dataGridView1.Columns["芯纸规格"].Visible = false;
            dataGridView1.Columns["底纸"].Visible = false;
            dataGridView1.Columns["底纸克重"].Visible = false;
            dataGridView1.Columns["表面加工"].Visible = false;
            dataGridView1.Columns["裱纸工艺"].Visible = false;

    
            dataGridView1.Columns["部品名"].Visible = true;
            dataGridView1.Columns["加工门幅"].Visible = true;
            dataGridView1.Columns["加工长度"].Visible = true;
            dataGridView1.Columns["部品总数"].Visible = true;
            dataGridView1.Columns["机器型号"].Visible = true;
            dataGridView1.Columns["部品单价"].Visible = true;
            dataGridView1.Columns["部品总价"].Visible = true;
            dataGridView1.Columns["面纸单价"].Visible = true;
            dataGridView1.Columns["面纸用量"].Visible = true;
            dataGridView1.Columns["面纸门幅"].Visible = true;
            dataGridView1.Columns["面纸纸长"].Visible = true;
            dataGridView1.Columns["面纸可用"].Visible = true;
            dataGridView1.Columns["面纸单个用量"].Visible = true;
            dataGridView1.Columns["面纸小计"].Visible = true;
            dataGridView1.Columns["芯纸单价"].Visible = true;
            dataGridView1.Columns["芯纸用量"].Visible = true;
            dataGridView1.Columns["芯纸门幅"].Visible = true;
            dataGridView1.Columns["芯纸纸长"].Visible = true;
            dataGridView1.Columns["芯纸可用"].Visible = true;
            dataGridView1.Columns["芯纸单个用量"].Visible = true;
            dataGridView1.Columns["芯纸小计"].Visible = true;
            dataGridView1.Columns["底纸单价"].Visible = true;
            dataGridView1.Columns["底纸用量"].Visible = true;
            dataGridView1.Columns["底纸单个用量"].Visible = true;
            dataGridView1.Columns["底纸小计"].Visible = true;
            dataGridView1.Columns["正反CTP合计"].Visible = true;
            dataGridView1.Columns["正反印工合计"].Visible = true;
            dataGridView1.Columns["表面处理单价"].Visible = true;
            dataGridView1.Columns["表面加工小计"].Visible = true;
            dataGridView1.Columns["裱工单价"].Visible = true;
            dataGridView1.Columns["裱工小计"].Visible = true;
            dataGridView1.Columns["刀模小计"].Visible = true;
            dataGridView1.Columns["模切小计"].Visible = true;

            if (checkBox2.Checked)
            {
              
              
                dataGridView1.Columns["部品名"].Visible = true;
                dataGridView1.Columns["加工门幅"].Visible = true;
                dataGridView1.Columns["加工长度"].Visible = true;
                dataGridView1.Columns["部品总数"].Visible = true;
                dataGridView1.Columns["机器型号"].Visible = true;
                dataGridView1.Columns["部品单价"].Visible = true;
                dataGridView1.Columns["部品总价"].Visible = true;
                dataGridView1.Columns["面纸单价"].Visible = true;
                dataGridView1.Columns["面纸用量"].Visible = true;
                dataGridView1.Columns["面纸内耗"].Visible = true;
                dataGridView1.Columns["面纸下单"].Visible = true;
                dataGridView1.Columns["面纸外耗"].Visible = true;
                dataGridView1.Columns["面纸门幅"].Visible = true;
                dataGridView1.Columns["面纸纸长"].Visible = true;
                dataGridView1.Columns["面纸可用"].Visible = true;
                dataGridView1.Columns["面纸单个用量"].Visible = true;
                dataGridView1.Columns["面纸小计"].Visible = true;
                dataGridView1.Columns["芯纸单价"].Visible = true;
                dataGridView1.Columns["芯纸内耗"].Visible = true;
                dataGridView1.Columns["芯纸用量"].Visible = true;
                dataGridView1.Columns["芯纸门幅"].Visible = true;
                dataGridView1.Columns["芯纸纸长"].Visible = true;
                dataGridView1.Columns["芯纸可用"].Visible = true;
                dataGridView1.Columns["芯纸单个用量"].Visible = true;
                dataGridView1.Columns["芯纸小计"].Visible = true;
                dataGridView1.Columns["底纸单价"].Visible = true;
                dataGridView1.Columns["底纸用量"].Visible = true;
                dataGridView1.Columns["底纸内耗"].Visible = true;
                dataGridView1.Columns["底纸下单"].Visible = true;
                dataGridView1.Columns["底纸外耗"].Visible = true;
                dataGridView1.Columns["底纸单个用量"].Visible = true;
                dataGridView1.Columns["底纸小计"].Visible = true;
                dataGridView1.Columns["印工单色单价"].Visible = true;
                dataGridView1.Columns["超出单色单张价"].Visible = true;
                dataGridView1.Columns["CTP单张价"].Visible = true;
                dataGridView1.Columns["正面色数共计"].Visible = true;
                dataGridView1.Columns["正面CTP张数"].Visible = true;
                dataGridView1.Columns["正面纸张损耗"].Visible = true;
                dataGridView1.Columns["正面防晒合计"].Visible = true;
                dataGridView1.Columns["正面CTP价计"].Visible = true;
                dataGridView1.Columns["正面印工合计"].Visible = true;
                dataGridView1.Columns["反面色数共计"].Visible = true;
                dataGridView1.Columns["反面CTP张数"].Visible = true;
                dataGridView1.Columns["反面纸张损耗"].Visible = true;
                dataGridView1.Columns["反面防晒合计"].Visible = true;
                dataGridView1.Columns["反面CTP价计"].Visible = true;
                dataGridView1.Columns["反面印工合计"].Visible = true;
                dataGridView1.Columns["正反CTP合计"].Visible = true;
                dataGridView1.Columns["正反印工合计"].Visible = true;
                dataGridView1.Columns["表面处理单价"].Visible = true;
                dataGridView1.Columns["无印刷表面处理损耗"].Visible = true;
                dataGridView1.Columns["表面处理用量"].Visible = true;
                dataGridView1.Columns["表面加工小计"].Visible = true;
                dataGridView1.Columns["裱工单价"].Visible = true;
                dataGridView1.Columns["裱工用量"].Visible = true;
                dataGridView1.Columns["裱工小计"].Visible = true;
                dataGridView1.Columns["刀模小计"].Visible = true;
                dataGridView1.Columns["模切小计"].Visible = true;

            }
            else
            {
               
            }
           
           
            dataGridView1.Columns["部品名"].HeaderCell.Style.BackColor = CCOLOR.CDET_WNAME;
            dataGridView1.Columns["加工门幅"].HeaderCell.Style.BackColor = CCOLOR.CPROCESSING_DOOR;
            dataGridView1.Columns["加工长度"].HeaderCell.Style.BackColor = CCOLOR.CPROCESSING_DOOR;
            dataGridView1.Columns["部品总数"].HeaderCell.Style.BackColor = CCOLOR.CPROCESSING_DOOR;
            dataGridView1.Columns["机器型号"].HeaderCell.Style.BackColor = CCOLOR.CDET_WNAME;
            dataGridView1.Columns["部品单价"].HeaderCell.Style.BackColor = CCOLOR.CDET_WNAME;
            dataGridView1.Columns["部品总价"].HeaderCell.Style.BackColor = CCOLOR.CDET_WNAME;
            dataGridView1.Columns["面纸单价"].HeaderCell.Style.BackColor = CCOLOR.CPROCESSING_DOOR;
            dataGridView1.Columns["面纸单个用量"].HeaderCell.Style.BackColor = Color.Yellow;
            dataGridView1.Columns["面纸小计"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["芯纸单价"].HeaderCell.Style.BackColor = CCOLOR.CPROCESSING_DOOR;
            dataGridView1.Columns["芯纸单个用量"].HeaderCell.Style.BackColor = Color.Yellow;
            dataGridView1.Columns["芯纸小计"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["底纸单价"].HeaderCell.Style.BackColor = CCOLOR.CPROCESSING_DOOR;
            dataGridView1.Columns["底纸单个用量"].HeaderCell.Style.BackColor = Color.Yellow;
            dataGridView1.Columns["底纸小计"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["正反CTP合计"].HeaderCell.Style.BackColor = CCOLOR.CPOSITIVE_AND_POSSITE_PRINTING;
            dataGridView1.Columns["正反印工合计"].HeaderCell.Style.BackColor = CCOLOR.CPOSITIVE_AND_POSSITE_PRINTING;
            dataGridView1.Columns["表面处理单价"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["表面加工小计"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["裱工单价"].HeaderCell.Style.BackColor = CCOLOR.CDET_WNAME;
            dataGridView1.Columns["裱工小计"].HeaderCell.Style.BackColor = CCOLOR.CDET_WNAME;
            dataGridView1.Columns["刀模小计"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["模切小计"].HeaderCell.Style.BackColor = CCOLOR.CTOTAL_TISSUE;
            dataGridView1.Columns["序号"].Width = 40;
            dataGridView1.Columns["项目号"].Width = 40;
            dataGridView1.Columns["项目名称"].Width = 40;
            dataGridView1.Columns["数量"].Width = 40;
            dataGridView1.Columns["部品名"].Width = 40;
            dataGridView1.Columns["加工门幅"].Width = 40;
            dataGridView1.Columns["加工长度"].Width = 40;
            dataGridView1.Columns["部品总数"].Width = 40;
            dataGridView1.Columns["机器型号"].Width = 40;
            dataGridView1.Columns["部品单价"].Width = 40;
            dataGridView1.Columns["部品总价"].Width = 40;
            dataGridView1.Columns["面纸单价"].Width = 40;
            dataGridView1.Columns["面纸用量"].Width = 40;
            dataGridView1.Columns["面纸门幅"].Width = 40;
            dataGridView1.Columns["面纸纸长"].Width = 40;
            dataGridView1.Columns["面纸可用"].Width = 40;
            dataGridView1.Columns["面纸单个用量"].Width = 40;
            dataGridView1.Columns["面纸小计"].Width = 40;
            dataGridView1.Columns["芯纸单价"].Width = 40;
            dataGridView1.Columns["芯纸用量"].Width = 40;
            dataGridView1.Columns["芯纸门幅"].Width = 40;
            dataGridView1.Columns["芯纸纸长"].Width = 40;
            dataGridView1.Columns["芯纸可用"].Width = 40;
            dataGridView1.Columns["芯纸单个用量"].Width = 40;
            dataGridView1.Columns["芯纸小计"].Width = 40;
            dataGridView1.Columns["底纸单价"].Width = 40;
            dataGridView1.Columns["底纸用量"].Width = 40;
            dataGridView1.Columns["底纸单个用量"].Width = 40;
            dataGridView1.Columns["底纸小计"].Width = 40;
            dataGridView1.Columns["正反CTP合计"].Width = 40;
            dataGridView1.Columns["正反印工合计"].Width = 40;
            dataGridView1.Columns["表面处理单价"].Width = 40;
            dataGridView1.Columns["表面加工小计"].Width = 40;
            dataGridView1.Columns["裱工单价"].Width = 40;
            dataGridView1.Columns["裱工小计"].Width = 40;
            dataGridView1.Columns["刀模小计"].Width = 40;
            dataGridView1.Columns["模切小计"].Width = 40;

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
        #region dgvStateControl_dgv9
        private void dgvStateControl_dgv9()
        {
            int i;
            dataGridView9.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView9.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView9.ClearSelection();
            dataGridView9.AllowUserToAddRows = false;
            int numCols1 = dataGridView9.Columns.Count;
            dataGridView9.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
            dataGridView9.Columns["元套"].HeaderText = "元/套";
            dataGridView9.Columns["序号"].Width = 40;
            dataGridView9.Columns["项目"].Width = 80;
            dataGridView9.Columns["元套"].Width = 50;
            dataGridView9.Columns["批量小计"].Width = 80;
            dataGridView9.Columns["主件用量"].Width = 60;
            dataGridView9.Columns["序号"].ReadOnly = true;
            dataGridView9.Columns["项目"].ReadOnly = true;
            dataGridView9.Columns["元套"].ReadOnly = true;
            dataGridView9.Columns["批量小计"].ReadOnly = true;

            for (i = 0; i < numCols1; i++)
            {

                dataGridView9.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView9.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView9.EnableHeadersVisualStyles = false;
                dataGridView9.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView9.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            }
            dataGridView9.Columns["元套"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView9.Columns["批量小计"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView9.Columns["主件用量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            for (i = 0; i < dataGridView9.Rows.Count; i++)
            {

                dataGridView9.Rows[i].Height = 18;
            }
            for (i = 0; i < dataGridView9.Rows.Count - 1; i++)
            {
                dataGridView9.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView9.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
            dataGridView9["主件用量", 1].ReadOnly = true;
            dataGridView9["主件用量", 3].ReadOnly = true;
            dataGridView9["主件用量", 5].ReadOnly = true;
            dataGridView9["主件用量", 7].ReadOnly = true;
            dataGridView9["主件用量", 15].ReadOnly = true;
            dataGridView9["主件用量", 17].ReadOnly = true;
            dataGridView9["主件用量", 18].ReadOnly = true;
            dataGridView9["主件用量", 19].ReadOnly = true;
        }
        #endregion
        public void WORKORDER_USE()
        {
            select = 1;

        }


        #region dataGridView1_DataSourceChanged
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {



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

        #region juage()

        private bool juage()
        {
            string v1 = "";
            bool b = false;

            if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 0].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 0].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 2].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 2].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 4].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 4].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 6].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 6].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 8].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 8].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 9].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 9].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 10].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 10].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 11].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 11].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 12].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 12].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno_HAVE_PERCENT(dataGridView9["主件用量", 14].FormattedValue.ToString()) == 0)
            {
                v1 = string.Format("序号 {0} 主件用量只能输入数值", dataGridView9["序号", 14].FormattedValue.ToString());
                MessageBox.Show(v1, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }

            return b;
        }
        #endregion
        private void COST_TOTAL_FormClosing(object sender, FormClosingEventArgs e)
        {
          
 
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_DataSourceChanged(object sender, EventArgs e)
        {

        }

        private void btnToExcel_Click_1(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 0)
            {
                bc.dgvtoExcel(dataGridView2, this.Text);
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView9_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView9.CurrentCell.RowIndex;
                int columnsindex = dataGridView9.CurrentCell.ColumnIndex;

                if ((rowsindex != 1 && rowsindex != 3 && rowsindex != 5 && rowsindex != 7 && rowsindex != 15) &&
                    dataGridView9.Columns[columnsindex].Name == "主件用量" &&
                    bc.yesno_HAVE_PERCENT(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
     
        }

        private void dataGridView9_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                cprinting_offer.RECEPTION_USE = true;
            if (!string.IsNullOrEmpty(dtx.Rows[13]["主件用量"].ToString()) && bc.yesno_HAVE_PERCENT(dtx.Rows[13]["主件用量"].ToString()) != 0)
            {
                cprinting_offer.MAIN_MANAGE = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[13]["主件用量"].ToString(), '%'));
            }
            else
            {
                cprinting_offer.MAIN_MANAGE = 0;
            }
            if (!string.IsNullOrEmpty(dtx.Rows[14]["主件用量"].ToString()) && bc.yesno_HAVE_PERCENT(dtx.Rows[14]["主件用量"].ToString()) != 0)
            {
                cprinting_offer.MAIN_PROFIT = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[14]["主件用量"].ToString(), '%'));
            }
            else
            {
                cprinting_offer.MAIN_PROFIT = 0;
            }
            if (!string.IsNullOrEmpty(dtx.Rows[16]["主件用量"].ToString()) && bc.yesno_HAVE_PERCENT(dtx.Rows[16]["主件用量"].ToString()) != 0)
            {
                cprinting_offer.MAIN_PURCHASE = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[16]["主件用量"].ToString(), '%'));
            }
            else
            {
                cprinting_offer.MAIN_PURCHASE = 0;
            }

      
          
            cprinting_offer.RETURN_BATCH_SUBTOTAL_COST_SET = 1;
            cprinting_offer.RETURN_COST_TOTAL_DT(PFID, dtt);
            int rowsindex = dataGridView9.CurrentCell.RowIndex;
            int columnsindex = dataGridView9.CurrentCell.ColumnIndex;
            dataGridView9["元套", 13].Value = (cprinting_offer.YUAN_SET_MANAGE).ToString("0.00");
            dataGridView9["元套", 14].Value = (cprinting_offer.YUAN_SET_PROFIT).ToString("0.00");

            dataGridView9["批量小计", 13].Value = (cprinting_offer.RETURN_BATCH_SUBTOTAL_MANAGE ).ToString("0");
            dataGridView9["批量小计", 14].Value = (cprinting_offer.RETURN_MAIN_DOSAGE_PROFIT ).ToString("0");

            dataGridView9["元套", 16].Value = (cprinting_offer.RETURN_YUAN_SET_PURCHASE_COST ).ToString("0.00");
            dataGridView9["批量小计", 16].Value = (cprinting_offer.RETURN_MAIN_DOSAGE_PURCHASE_COST).ToString("0");

            dataGridView9["元套", 17].Value = (cprinting_offer.YUAN_SET_NO_TAX).ToString("0.00");
            dataGridView9["批量小计", 17].Value = (cprinting_offer.BATCH_TOTAL_NO_TAX).ToString("0");

            dataGridView9["元套", 18].Value = (cprinting_offer.YUAN_SET_HAVE_TAX).ToString("0.00");
            dataGridView9["批量小计", 18].Value = (cprinting_offer.BATCH_TOTAL_HAVE_TAX).ToString("0");

            dataGridView9["元套", 19].Value = (cprinting_offer.YUAN_SET_PURCHASE_PERCENT).ToString("0.0") + "%";
            dataGridView9["主件用量", 19].Value = (cprinting_offer.MAIN_DOSAGE_PURCHASE_PERCENT).ToString("0.0") + "%";

                
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView3_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView3.Columns.Count; i++)
            {
                if (dataGridView3.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                  dataGridView3.Columns[i].DefaultCellStyle.Format = "#0.00";
                  dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void dataGridView4_DataSourceChanged(object sender, EventArgs e)
        {
            for (i = 0; i < dataGridView4.Columns.Count; i++)
            {
                if (dataGridView4.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView4.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView4.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void dataGridView5_DataSourceChanged(object sender, EventArgs e)
        {
            for (i = 0; i < dataGridView5.Columns.Count; i++)
            {
                if (dataGridView5.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView5.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView5.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void dataGridView6_DataSourceChanged(object sender, EventArgs e)
        {
            for (i = 0; i < dataGridView6.Columns.Count; i++)
            {
                if (dataGridView6.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView6.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView6.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }

        }

        private void dataGridView7_DataSourceChanged(object sender, EventArgs e)
        {
            for (i = 0; i < dataGridView7.Columns.Count; i++)
            {
                if (dataGridView7.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView7.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView7.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }

        }

        private void dataGridView8_DataSourceChanged(object sender, EventArgs e)
        {
            for (i = 0; i < dataGridView8.Columns.Count; i++)
            {
                if (dataGridView8.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView8.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView8.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            /*if (dtt.Rows.Count > 0)
            {
                if (checkBox2.Checked)
                {
                    dt = cprinting_offer.RETURN_DT_SHOW_HIDE(dtt);
                    dataGridView1.DataSource = null;//将DATAGRIDVIEW按件数据源清空，避免受之前数据源影响使加载的数据在显示时栏位顺序不对
                }
                else
                {
                    dt = cprinting_offer.RETURN_DT_TO_EXCEL(dtt);

                }

            }*/
            if (dt.Rows.Count > 0)
            {

                dataGridView1.DataSource = dt;
                dgvStateControl();
            }
            else
            {
                hint.Text = "找不到所要搜索项！";
                dataGridView1.DataSource = null;

            }
        }

        private void btnToExcel_Click_2(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, this.Text );

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
           
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Focus();
            try
            {
                if (juage())
                {
                }
                else
                {
                    if (textBox1.Text != "")
                    {
                        basec.getcoms(@"UPDATE PRINTING_OFFER_MST SET AUDIT_OPINION='" + textBox1.Text + "'   WHERE PFID='" + PFID + "'");
                    }
                    cprint_cost_total.MAKERID = LOGIN.EMID;
                    cprint_cost_total.PFID = PFID;
                    cprint_cost_total.save(dtx);
                    IFExecution_SUCCESS = true;
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            F2.bind("");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, "客户信息");

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void COST_TOTAL_Resize(object sender, EventArgs e)
        {
          
         
        }

   



 

  

      
    }
}
