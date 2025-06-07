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
using System.Threading;

namespace CSPSS.OFFER_MANAGE
{
    public partial class PRINTING_OFFERT : Form
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable dt4 = new DataTable();
        DataTable dt5 = new DataTable();
        DataTable dt6 = new DataTable();
        DataTable dt7 = new DataTable();
        DataTable dt8 = new DataTable();
        StringBuilder sqb = new StringBuilder();
        basec bc = new basec();
        #region nature
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _OFFER_ID;
        public string OFFER_ID
        {
            set { _OFFER_ID = value; }
            get { return _OFFER_ID; }
        }
        private bool _IF_COMPLETED;
        public bool IF_COMPLETED
        {
            set { _IF_COMPLETED = value; }
            get { return _IF_COMPLETED; }
        }
        private decimal _PACK_LENGTH;
        public decimal PACK_LENGTH
        {
            set { _PACK_LENGTH = value; }
            get { return _PACK_LENGTH; }

        }
        private decimal _TOTAL_BOXS_COUNT;
        public decimal TOTAL_BOXS_COUNT
        {
            set { _TOTAL_BOXS_COUNT = value; }
            get { return _TOTAL_BOXS_COUNT; }

        }
        private decimal _PACK_WIDTH;
        public decimal PACK_WIDTH
        {
            set { _PACK_WIDTH = value; }
            get { return _PACK_WIDTH; }

        }
        private decimal _PACK_HEIGHT;
        public decimal PACK_HEIGHT
        {
            set { _PACK_HEIGHT = value; }
            get { return _PACK_HEIGHT; }

        }
        private string _CUSTOMER_TYPE;
        public string CUSTOMER_TYPE
        {
            set { _CUSTOMER_TYPE = value; }
            get { return _CUSTOMER_TYPE; }
        }
        private string _EDIT;
        public string EDIT
        {
            set { _EDIT = value; }
            get { return _EDIT; }
        }
        private string _PIID;
        public string PIID
        {
            set { _PIID = value; }
            get { return _PIID; }

        }
        private string _PROJECT_ID;
        public string PROJECT_ID
        {
            set { _PROJECT_ID = value; }
            get { return _PROJECT_ID; }

        }
        private static string _GET_PROJECT_ID;
        public static string GET_PROJECT_ID
        {
            set { _GET_PROJECT_ID = value; }
            get { return _GET_PROJECT_ID; }

        }
        private static string _POSITIVE_4C;
        public static string POSITIVE_4C
        {
            set { _POSITIVE_4C = value; }
            get { return _POSITIVE_4C; }

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
        private static string _TISSUE_SPEC;
        public static string TISSUE_SPEC
        {
            set { _TISSUE_SPEC = value; }
            get { return _TISSUE_SPEC; }

        }
        private static string _WEIGHT;
        public static string WEIGHT
        {
            set { _WEIGHT = value; }
            get { return _WEIGHT; }

        }
        private static string _PRINT_OPTION;
        public static string PRINT_OPTION
        {
            set { _PRINT_OPTION = value; }
            get { return _PRINT_OPTION; }

        }
        private static string _PAPER_CORE;
        public static string PAPER_CORE
        {
            set { _PAPER_CORE = value; }
            get { return _PAPER_CORE; }

        }
        private static string _SPEC;
        public static string SPEC
        {
            set { _SPEC = value; }
            get { return _SPEC; }

        }
        private static string _POSITIVE_COLOR;
        public static string POSITIVE_COLOR
        {
            set { _POSITIVE_COLOR = value; }
            get { return _POSITIVE_COLOR; }

        }
        private string _SAMPLE_CODE;
        public string SAMPLE_CODE
        {
            set { _SAMPLE_CODE = value; }
            get { return _SAMPLE_CODE; }

        }
        private string _SAMPLE_CODE_FIRST;
        public string SAMPLE_CODE_FIRST
        {
            set { _SAMPLE_CODE_FIRST = value; }
            get { return _SAMPLE_CODE_FIRST; }

        }
        private static string _POSITIVE_SUN_SCREEN;
        public static string POSITIVE_SUN_SCREEN
        {
            set { _POSITIVE_SUN_SCREEN = value; }
            get { return _POSITIVE_SUN_SCREEN; }

        }
        private static string _SURFACE_PROCESSING;
        public static string SURFACE_PROCESSING
        {
            set { _SURFACE_PROCESSING = value; }
            get { return _SURFACE_PROCESSING; }

        }
        private static string _LAMINATING_PROCESS;
        public static string LAMINATING_PROCESS
        {
            set { _LAMINATING_PROCESS = value; }
            get { return _LAMINATING_PROCESS; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        #endregion
        private delegate bool dele(string a1, string a2);
        private delegate void delex();
        PRINTING_OFFER F1 = new PRINTING_OFFER();
        protected int M_int_judge, i;
        protected int select;
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        CTISSUE_SPEC ctissue_spec = new CTISSUE_SPEC();
        CPAPER_CORE cpaper_core = new CPAPER_CORE();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        CDIE_CUTTING_COST cdie_cutting_cost = new CDIE_CUTTING_COST();
        CPRINT_DIE_CUTTING cprint_die_cutting = new CPRINT_DIE_CUTTING();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        CPORTRAY cportray = new CPORTRAY();
        CPRINT_PORTRAY cprint_portray = new CPRINT_PORTRAY();
        CPARTS_AUXILIARY cparts_auxiliary = new CPARTS_AUXILIARY();
        CPRINT_PARTS_AUXILIARY cprint_parts_auxiliary = new CPRINT_PARTS_AUXILIARY();
        CPACK_MATERIAL cpack_material = new CPACK_MATERIAL();
        CPRINT_PACK_MATERIAL cprint_pack_material = new CPRINT_PACK_MATERIAL();
        CARTIFICIAL cartificial = new CARTIFICIAL();
        CPRINT_ARTIFICIALL cprint_artificial = new CPRINT_ARTIFICIALL();
        CPURCHASE cpurchase = new CPURCHASE();
        CPRINT_PURCHASE cprint_purchase = new CPRINT_PURCHASE();
        CTRANSPORT ctrasport = new CTRANSPORT();
        CPRINT_TRANSPORT cprint_transport = new CPRINT_TRANSPORT();
        COTHER_COST cother_cost = new COTHER_COST();
        CPRINT_COST_TOTAL cprint_cost_total = new CPRINT_COST_TOTAL();
        DataTable dtt = new DataTable();
        DataTable dtt_hide = new DataTable();
        LOADING frm_loading= new LOADING();
   
        public PRINTING_OFFERT()
        {
           
            InitializeComponent();
        }
     
        public PRINTING_OFFERT(PRINTING_OFFER FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        private void loading(object sender, DoWorkEventArgs e)
        {
          
                for (i = 0; i < 100; i++)
                {
                    //MessageBox.Show(i.ToString ());
                    System.Threading.Thread.Sleep(100);
                    if (IF_COMPLETED)
                    {
                        groupBox1.Visible = true;
                        label7.Visible = true;
                        progressBar1.Visible = true;
                        groupBox2.Visible = true;
                        groupBox3.Visible = true;
                        groupBox11.Visible = true;
                        frm_loading.Close();
               
                        break;
                    }
                }
        }
        private void loading()
        {
         
            if (ADD_OR_UPDATE == "ADD")
            {
                frm_loading.Show();
                IF_COMPLETED = false;
                groupBox1.Visible = false;
                label7.Visible = false;
                progressBar1.Visible = false;
                groupBox2.Visible = false;
                groupBox3.Visible = false;
                groupBox11.Visible = false;
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
        }
        private void PRINTING_OFFERT_Load(object sender, EventArgs e)
        {
            //label7.ForeColor = CCOLOR.XZL;
            //label7.Font = new Font("微软雅黑", 11, FontStyle.Bold);//含字体粗体的FONT有三个参数 16/01/22

            //loading();
            //pictureBox2.Visible = true;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
            comboBox2.Text = "Z";
            comboBox2.Enabled = false;
            Control.CheckForIllegalCrossThreadCalls = false;//避免出现线程间操作无效: 从不是创建控件“progressBar1”的线程访问它 160120
            if (Screen.AllScreens[0].Bounds.Width == 1366 && Screen.AllScreens[0].Bounds.Height == 768 ||
                    Screen.AllScreens[0].Bounds.Width == 1280 && Screen.AllScreens[0].Bounds.Height == 800)
            {
                groupBox11.Height = 125;
                groupBox3.Location = new Point(3, 330);
                groupBox3.Height = 375;

            }
            else if (Screen.AllScreens[0].Bounds.Width == 1920 && Screen.AllScreens[0].Bounds.Height == 1080)
            {
                this.AutoScroll = true;
                this.AutoScrollMinSize = new Size(1920, 1080);
            }
            else
            {
                this.AutoScroll = true;
                this.AutoScrollMinSize = new Size(1920, 1080);
            }
           // MessageBox.Show(Screen.AllScreens[0].Bounds.Width + "," + Screen.AllScreens[0].Bounds.Height);
         
            try
            {

                hint.ForeColor = Color.Red;
                textBox3.ReadOnly = true;
              this.Icon = Resource1.xz_200X200;

                //IDO = bc.getOnlyString("SELECT PFID FROM PRINTING_OFFER_MST WHERE OFFER_ID in ('1601Z004-02-ADM','1601Z004-02-ADM-A')");
                total1();
                bind();
                right();
                IF_COMPLETED = true;
            }
            catch (Exception EX)
            {

                MessageBox.Show(EX.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #region total1
        private void total1()
        {
           
            dt= basec.getdts(cprinting_offer.sql + " where A.PFID='" + IDO + "' ORDER BY  B.PFID ASC ");
            dt1 = bc.getdt("SELECT PRINT_OPTION FROM PRINT_OPTION");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "PRINT_OPTION", "印刷选项", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                印刷选项.Items.Clear();
                印刷选项.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    印刷选项.Items.Add(dr["VALUE"].ToString());
                }
            }
            dt1 = bc.getdt(string.Format(@"SELECT *  FROM TISSUE_SPEC_MST 
WHERE SUBSTRING(CUSTOMER_TYPE,1,1)='{0}' ORDER BY TSID ASC", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "TISSUE_SPEC", "面纸", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {

                面纸.Items.Clear();
                面纸.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    面纸.Items.Add(dr["VALUE"].ToString());
                }
            }
            dt1 = bc.getdt(string.Format("SELECT * FROM PAPER_CORE_MST WHERE SUBSTRING(CUSTOMER_TYPE,1,1)='{0}' ORDER BY PCID ASC",
                bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "PAPER_CORE", "芯纸", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                芯纸.Items.Clear();
                芯纸.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    芯纸.Items.Add(dr["VALUE"].ToString());
                }
            }
            dt1 = bc.getdt(string.Format(@"SELECT *  FROM TISSUE_SPEC_MST 
WHERE SUBSTRING(CUSTOMER_TYPE,1,1)='{0}' ORDER BY TSID ASC", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "TISSUE_SPEC", "底纸", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                底纸.Items.Clear();
                底纸.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    底纸.Items.Add(dr["VALUE"].ToString());
                }
            }
            dt1 = bc.getdt("SELECT PRIMARY_COLORS FROM PRIMARY_COLORS");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "PRIMARY_COLORS", "正面4C", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                正面4C.Items.Clear();
                正面4C.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    正面4C.Items.Add(dr["VALUE"].ToString());
                }
            }
            dt1 = bc.getdt("SELECT COLOR_PARAMETERS FROM COLOR_PARAMETERS");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "COLOR_PARAMETERS", "正面专色", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                正面专色.Items.Clear();
                正面专色.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    正面专色.Items.Add(dr["VALUE"].ToString());
                }
            }
            正面防晒.Items.Clear();
            正面防晒.Items.Add("");
            正面防晒.Items.Add("是");
            正面防晒.Items.Add("否");
            双面印刷.Items.Clear();
            双面印刷.Items.Add("");
            双面印刷.Items.Add("双异");
            双面印刷.Items.Add("单异");
            dt1 = bc.getdt("SELECT PRIMARY_COLORS FROM PRIMARY_COLORS");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "PRIMARY_COLORS", "反面4C", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                反面4C.Items.Clear();
                反面4C.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    反面4C.Items.Add(dr["VALUE"].ToString());
                }
            }
            dt1 = bc.getdt("SELECT COLOR_PARAMETERS FROM COLOR_PARAMETERS");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "COLOR_PARAMETERS", "反面专色", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                反面专色.Items.Clear();
                反面专色.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    反面专色.Items.Add(dr["VALUE"].ToString());
                }
            }
            反面防晒.Items.Clear();
            反面防晒.Items.Add("");
            反面防晒.Items.Add("是");
            反面防晒.Items.Add("否");
            dt1 = bc.getdt("SELECT SURFACE_PROCESSING FROM SURFACE_PROCESSING");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "SURFACE_PROCESSING", "表面加工", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                表面加工.Items.Clear();
                表面加工.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    表面加工.Items.Add(dr["VALUE"].ToString());
                }
            }
            表面次数.Items.Clear();
            表面次数.Items.Add("");
            表面次数.Items.Add("1");
            表面次数.Items.Add("2");
            dt1 = bc.getdt("SELECT LAMINATING_PROCESS FROM LAMINATING_PROCESS");
            dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dt, "LAMINATING_PROCESS", "裱纸工艺", 0, dt.Rows.Count);
            if (dt1.Rows.Count > 0)
            {
                裱纸工艺.Items.Clear();
                裱纸工艺.Items.Add("");
                foreach (DataRow dr in dt1.Rows)
                {
                    裱纸工艺.Items.Add(dr["VALUE"].ToString());
                }
            } 裱纸次数.Items.Clear();
            裱纸次数.Items.Add("");
            裱纸次数.Items.Add("1");
            裱纸次数.Items.Add("2");
            模切.Items.Clear();
            模切.Items.Add("");
            模切.Items.Add("是");
            模切.Items.Add("否");

            dataGridView1["项次", 0].Value = 1;
            面纸克重.ReadOnly = true;
            芯纸规格.ReadOnly = true;
            底纸克重.ReadOnly = true;
            反面4C.ReadOnly = true;
            反面专色.ReadOnly = true;
            反面防晒.ReadOnly = true;
            芯纸.ReadOnly = true;
            底纸.ReadOnly = true;
            双面印刷.ReadOnly = true;

        }
        #endregion
        #region bind
        private void bind()
        {

            #region main
            印刷选项.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            面纸.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            面纸克重.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            芯纸.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            芯纸规格.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            底纸.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            底纸克重.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            正面4C.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            正面专色.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            正面防晒.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            双面印刷.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            反面4C.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            反面专色.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            反面防晒.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            表面加工.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            表面次数.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            裱纸工艺.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            裱纸次数.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            模切.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;


            comboBox1.Focus();
         
            comboBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
            comboBox2.BackColor = CCOLOR.CUSTOMER_YELLOW;
            textBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
            textBox2.BackColor = CCOLOR.qmhs;
            textBox2.ReadOnly = true;

            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {

                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
            DataTable dtx = basec.getdts(cprinting_offer.sql + " where A.PFID='" + IDO + "' ORDER BY  B.PFID ASC ");
            if (dtx.Rows.Count > 0)
            {
               
                comboBox1.Text = dtx.Rows[0]["项目号"].ToString();
                textBox1.Text = dtx.Rows[0]["数量"].ToString();
                textBox2.Text = dtx.Rows[0]["报价编号"].ToString();
                if (bc.RETURN_UNTIL_CHAR(dtx.Rows[0]["报价编号"].ToString(), '-').Length == 7)
                {
                    comboBox2.Text = dtx.Rows[0]["报价编号"].ToString().Substring(3, 1);

                }
                else
                {
                    comboBox2.Text = dtx.Rows[0]["报价编号"].ToString().Substring(4, 1);
                }

                DataTable dt = cprinting_offer.GetTableInfo();
                foreach (DataRow dr1 in dtx.Rows)
                {
              
                    DataRow dr = dt.NewRow();
                    //dr["项次"] = dr1["项次"].ToString();
                    dr["部品名"] = dr1["部品名"].ToString();
                    dr["图纸门幅"] = dr1["图纸门幅"].ToString();
                    dr["图纸纸长"] = dr1["图纸纸长"].ToString();
                    dr["部品个数"] = dr1["部品个数"].ToString();
                    dr["拼模数"] = dr1["拼模数"].ToString();
                    dr["部品数"] = dr1["部品数"].ToString();
                    dr["印刷选项"] = dr1["印刷选项"].ToString();
                    dr["面纸"] = dr1["面纸"].ToString();
                    dr["面纸克重"] = dr1["面纸克重"].ToString();
                    dr["芯纸"] = dr1["芯纸"].ToString();
                    dr["芯纸规格"] = dr1["芯纸规格"].ToString();
                    dr["底纸"] = dr1["底纸"].ToString();
                    dr["底纸克重"] = dr1["底纸克重"].ToString();
                    dr["正面4C"] = dr1["正面4C"].ToString();
                    dr["正面专色"] = dr1["正面专色"].ToString();
                    dr["正面防晒"] = dr1["正面防晒"].ToString();
                    dr["双面印刷"] = dr1["双面印刷"].ToString();
                    dr["反面4C"] = dr1["反面4C"].ToString();
                    dr["反面专色"] = dr1["反面专色"].ToString();
                    dr["反面防晒"] = dr1["反面防晒"].ToString();
                    dr["表面加工"] = dr1["表面加工"].ToString();
                    dr["表面次数"] = dr1["表面次数"].ToString();
                    dr["裱纸工艺"] = dr1["裱纸工艺"].ToString();
                    dr["裱纸次数"] = dr1["裱纸次数"].ToString();
                    dr["模切"] = dr1["模切"].ToString();
                    dt.Rows.Add(dr);

                }

                dataGridView1.DataSource = dt;
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewComboBoxCell dgvcc = (DataGridViewComboBoxCell)dataGridView1["面纸克重", i];
                    if (dataGridView1["面纸", i].FormattedValue.ToString() != "")
                    {
                        dataGridView1["面纸克重", i].ReadOnly = false;
                        dgvcc = (DataGridViewComboBoxCell)dataGridView1["面纸克重", i];

                        dt1 = bc.getdt(string.Format(ctissue_spec.sql + @" WHERE B.TISSUE_SPEC='{0}' AND  
                    SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dataGridView1["面纸", i].Value.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                        DataTable dtx1 =bc. GET_DT_TO_DV_TO_DT(dt, "", string.Format ("面纸='{0}'",dataGridView1["面纸", i].Value.ToString()));
                        dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dtx1, "克重", "面纸克重", 0, dtx1.Rows.Count);
                        if (dt1.Rows.Count > 0)
                        {
                            dgvcc.Items.Clear();
                            dgvcc.Items.Add("");
                            foreach (DataRow dr in dt1.Rows)
                            {
                                dgvcc.Items.Add(dr["VALUE"].ToString());
                            }

                        }
                    }
                    else
                    {
                        dataGridView1["面纸克重", i].ReadOnly = true;
                    }

                    if (Convert.ToString(dataGridView1["芯纸", i].Value).Trim() != "")
                    {
                        dataGridView1["芯纸规格", i].ReadOnly = false;
                        dgvcc = (DataGridViewComboBoxCell)dataGridView1["芯纸规格", i];

                        dt1 = bc.getdt(string.Format(cpaper_core.sql + @" WHERE B.PAPER_CORE='{0}' AND  
                    SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}' ", dataGridView1["芯纸", i].Value.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                        DataTable dtx1 = bc.GET_DT_TO_DV_TO_DT(dt, "", string.Format("芯纸='{0}'", dataGridView1["芯纸", i].Value.ToString()));
                        dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dtx1, "规格", "芯纸规格", 0, dtx1.Rows.Count);
                        if (dt1.Rows.Count > 0)
                        {
                            dgvcc.Items.Clear();
                            dgvcc.Items.Add("");
                            foreach (DataRow dr in dt1.Rows)
                            {
                                dgvcc.Items.Add(dr["VALUE"].ToString());
                            }

                        }
                    }
                    else
                    {
                        dataGridView1["芯纸规格", i].ReadOnly = true;
                    }

                    if (dataGridView1["底纸", i].FormattedValue.ToString() != "")
                    {
                        dataGridView1["底纸克重", i].ReadOnly = false;
                        dgvcc = (DataGridViewComboBoxCell)dataGridView1["底纸克重", i];

                        dt1 = bc.getdt(string.Format(ctissue_spec.sql + @" WHERE B.TISSUE_SPEC='{0}' AND  
                    SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}' ", dataGridView1["底纸", i].Value.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                        DataTable dtx1 = bc.GET_DT_TO_DV_TO_DT(dt, "", string.Format("底纸='{0}'", dataGridView1["底纸", i].Value.ToString()));
                        dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dtx1, "克重", "底纸克重", 0, dtx1.Rows.Count);
                        if (dt1.Rows.Count > 0)
                        {
                            dgvcc.Items.Clear();
                            dgvcc.Items.Add("");
                            foreach (DataRow dr in dt1.Rows)
                            {
                                dgvcc.Items.Add(dr["VALUE"].ToString());
                            }

                        }
                    }
                    else
                    {
                        dataGridView1["底纸克重", i].ReadOnly = true;
                    }
                }
            }
            else
            {
                total1();

            }

            AUDIT();
            if (PROJECT_NAME != null)
            {
                comboBox1.Text = PROJECT_ID;
                textBox3.Text = PROJECT_NAME;
                textBox1.Focus();
            }
            dgvStateControl();
            #endregion
   
            try
            {
                bind_die_cutting_price();
                bind_portray();
                bind_PARTS_AUXILIARY();
                bind_pack_material();
                bind_artificial();
                bind_purchase();
                bind_transport();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        #endregion
        #region  bind_die_cutting_price
        private void bind_die_cutting_price()
        {
            DataGridViewComboBoxColumn dgvc刀模项目 = new DataGridViewComboBoxColumn();
            DataGridViewTextBoxColumn dgvc刀模长米 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc元米 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc圆孔个数 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc元个 = new DataGridViewTextBoxColumn();

            dataGridView2.Columns.Add(dgvc刀模项目);
            dataGridView2.Columns.Add(dgvc刀模长米);
            dataGridView2.Columns.Add(dgvc元米);
            dataGridView2.Columns.Add(dgvc圆孔个数);
            dataGridView2.Columns.Add(dgvc元个);


            dgvc刀模项目.Name = "项目";
            dgvc刀模长米.Name = "刀模长米";
            dgvc元米.Name = "元米";
            dgvc圆孔个数.Name = "圆孔个数";
            dgvc元个.Name = "元个";


            for (int i = 0; i < 3; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView2.Rows.Add(dgvr);
            }
            dt = bc.getdt(cprint_die_cutting.sql + string.Format(" WHERE A.PFID='{0}'", IDO));
            dtx = bc.getdt(cdie_cutting_cost.sql);
            dtx = bc.GET_DT_TO_DV_TO_DT(dtx, "", "项目 NOT IN ('按平方','圆孔')");
            dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "项目",0,dt.Rows .Count );
            DataGridViewComboBoxCell dgvcc1 = new DataGridViewComboBoxCell();
            dgvcc1.Items.Add("");
            dgvcc1.Items.Add("按平方");
            dgvcc1.Items.Add("按米计");

            if (dtx.Rows.Count > 0)
            {
                dgvc刀模项目.Items.Clear();
                dgvc刀模项目.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {
                    dgvc刀模项目.Items.Add(dr["VALUE"].ToString());
                }
            }
  
            dgvc刀模项目.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            dataGridView2["项目", 2].ReadOnly = true;
            dgvc元个.HeaderText = "元/个";
            dgvc元米.HeaderText = "元/米";
            dataGridView2["元米", 2] = dgvcc1;
            if (dt.Rows.Count > 0 && dt.Rows.Count <= 3)
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    dataGridView2["项目", i].Value = dt.Rows[i]["项目"].ToString();
                    dataGridView2["刀模长米", i].Value = dt.Rows[i]["刀模长米"].ToString();
                    
                    dataGridView2["圆孔个数", i].Value = dt.Rows[i]["圆孔个数"].ToString();
                 
                    if (dt.Rows[i]["元米"].ToString() == "按平方")
                    {
                        dataGridView2["项目", i].ReadOnly = false;
                    }
                    if (i == 1)
                    {
                        dataGridView2["元米", i].Value = dt.Rows[i]["元米"].ToString();//第一行不显示元米 元个单价 16/01/13
                        dataGridView2["元个", i].Value = dt.Rows[i]["元个"].ToString();//第一行不显示元米 元个单价 16/01/13
                    }
                    if (i == 2)
                    {
                        dataGridView2["元米", i].Value = dt.Rows[i]["元米"].ToString();//第三行要显示按米计还是按平方计，第一行是单价隐藏不显示 16/05/20
                    }
                }
            }
            dataGridView2["刀模长米", 2].ReadOnly = true;
            dgvc刀模项目.ReadOnly = true;
            dataGridView2["元个", 2].ReadOnly = true;
            dataGridView2["元米", 0].ReadOnly = true;//第一行元米 元个不能修改不能显示单价 16/01/13
            dataGridView2["元个", 0].ReadOnly = true;//第一行元米 元个不能修改不能显示单价 16/01/13
            dataGridView2["圆孔个数", 2].ReadOnly = true;
       
            dataGridView2["刀模长米", 2].Value = "刀模计价";
            dataGridView2["元个", 2].Value = "单套计";
            dgvStateControl_dgv2();
        }
        #endregion
        #region  bind_portray
        private void bind_portray()
        {
       
            DataGridViewTextBoxColumn dgvc写真类型 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc长 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc宽 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc总数量 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc单价 = new DataGridViewTextBoxColumn();
            dataGridView3.Columns.Add(dgvc写真类型);
            dataGridView3.Columns.Add(dgvc长);
            dataGridView3.Columns.Add(dgvc宽);
            dataGridView3.Columns.Add(dgvc总数量);
            dataGridView3.Columns.Add(dgvc单价);
            dgvc写真类型.Name = "写真类型";
            dgvc长.Name = "长";
            dgvc宽.Name = "宽";
            dgvc总数量.Name = "总数量";
            dgvc单价.Name = "单价";
            dt = bc.getdt(cprint_portray.sql + string.Format(" WHERE A.PFID='{0}'", IDO));
            for (int i = 0; i < 9; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView3.Rows.Add(dgvr);
            }
         
            dtx = bc.getdt(string.Format(cportray.sql + @" WHERE 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            //MessageBox.Show(bc.RETURN_CUSTOMER_TYPE(comboBox1.Text));
            dtx = bc.GET_DT_TO_DV_TO_DT(dtx, "", "写真类型 NOT IN ('批次写真运费')");
            if (dt.Rows.Count >= 7)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "写真类型", 0, 7);//此为正常写入数据库时 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "写真类型", 0, dt.Rows.Count);//此为写入数据库异常，数据没有写入数据库 160120
            }
            dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "VALUE");
            for (i = 0; i < 7; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();
                if (dtx.Rows.Count > 0)
                {
                    c1.Items.Clear();
                    c1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c1.Items.Add(dr["VALUE"].ToString());
                    }
                }
                dataGridView3["写真类型", i] = c1;
                dataGridView3["单价", i].ReadOnly = true;//含组合框的行单价不只能调BOM，不能在录入时手动修改 16/01/12
            }
        
            //dgvc写真类型.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            dgvc长.HeaderText = "长(mm)";
            dgvc宽.HeaderText = "宽(mm)";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows .Count-1 ; i++)//dt.rows.count-1是不显示最后一行，避免汇总二字写到单价字段保存出错 16/01/05
                {
                    dataGridView3["写真类型", i].Value = dt.Rows[i]["写真类型"].ToString();
                    dataGridView3["长", i].Value = dt.Rows[i]["长"].ToString();
                    dataGridView3["宽", i].Value = dt.Rows[i]["宽"].ToString();
                    dataGridView3["总数量", i].Value = dt.Rows[i]["总数量"].ToString();
                }
                for (int i = 7; i < dt.Rows .Count -1; i++)
                {
                    dataGridView3["单价", i].Value =dt.Rows[i]["单价"].ToString();//只有手输入单价的行显示单价
                }
            }
            dgvStateControl_dgv3();
        }
        #endregion
        #region  bind_PARTS_AUXILIARY
        private void bind_PARTS_AUXILIARY()
        {
            DataGridViewTextBoxColumn dgvc序号 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc配件名 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc用量 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc单价 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc单位 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc备注 = new DataGridViewTextBoxColumn();
            dataGridView4.Columns.Add(dgvc序号);
            dataGridView4.Columns.Add(dgvc配件名);
            dataGridView4.Columns.Add(dgvc用量);
            dataGridView4.Columns.Add(dgvc单价);
            dataGridView4.Columns.Add(dgvc单位);
            dataGridView4.Columns.Add(dgvc备注);
            dgvc序号.Name = "序号";
            dgvc配件名.Name = "配件名";
            dgvc用量.Name = "用量";
            dgvc单价.Name = "单价";
            dgvc单位.Name = "单位";
            dgvc备注.Name = "备注";

            for (int i = 0; i < 13; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView4.Rows.Add(dgvr);
                dataGridView4["序号", i].Value = i + 1;

            }
            dgvc序号.ReadOnly = true;
            dt = bc.getdt(cprint_parts_auxiliary.sql + string.Format(" WHERE A.PFID='{0}'", IDO));
            dtx = bc.getdt(cparts_auxiliary.sql);
            if (dt.Rows.Count >= 8)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "配件名", 0, 8);//此为正常写入数据库时 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "配件名", 0, dt.Rows .Count );//此为写入数据库异常，数据没有写入数据库 160120
            }
            for (i = 0; i < 8; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();
          
                if (dtx.Rows.Count > 0)
                {
                    c1.Items.Clear();
                    c1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c1.Items.Add(dr["VALUE"].ToString());
                    }
                }
                dataGridView4["配件名", i] = c1;
                dataGridView4["单价", i].ReadOnly = true;//含组合框的行单价不只能调BOM，不能在录入时手动修改 16/01/12
            }
    
            if (dt.Rows.Count > 0)
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    dataGridView4["配件名", i].Value = dt.Rows[i]["配件名"].ToString();
                    dataGridView4["用量", i].Value = dt.Rows[i]["用量"].ToString();
            
                    dataGridView4["单位", i].Value = dt.Rows[i]["单位"].ToString();
                    dataGridView4["备注", i].Value = dt.Rows[i]["备注"].ToString();
                }
                for (i = 8; i < dt.Rows.Count; i++)
                {
                    dataGridView4["单价", i].Value = dt.Rows[i]["单价"].ToString();//只有手输单价的行显示单价 16/01/12
                }
            }
            dgvStateControl_dgv4();
        }
        #endregion
        #region  bind_pack_material
        private void bind_pack_material()
        {
            DataGridViewTextBoxColumn dgvc序号 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc项目 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc数量 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc长 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc宽 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc高 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc箱形 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc材质 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc单价 = new DataGridViewTextBoxColumn();
            dataGridView5.Columns.Add(dgvc序号);
            dataGridView5.Columns.Add(dgvc项目);
            dataGridView5.Columns.Add(dgvc数量);
            dataGridView5.Columns.Add(dgvc长);
            dataGridView5.Columns.Add(dgvc宽);
            dataGridView5.Columns.Add(dgvc高);
            dataGridView5.Columns.Add(dgvc箱形);
            dataGridView5.Columns.Add(dgvc材质);
            dataGridView5.Columns.Add(dgvc单价);
            dgvc序号.Name = "序号";
            dgvc项目.Name = "项目";
            dgvc数量.Name = "数量";
            dgvc长.Name = "长";
            dgvc宽.Name = "宽";
            dgvc高.Name = "高";
            dgvc箱形.Name = "箱形";
            dgvc材质.Name = "材质";
            dgvc单价.Name = "单价";

            dgvc项目.ReadOnly = true;

            for (int i = 0; i < 10; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView5.Rows.Add(dgvr);
                dataGridView5["序号", i].Value = i + 1;

            }
            dataGridView5["项目", 0].Value = "最大外箱";
            dataGridView5["项目", 1].Value = "产品外箱";
            dataGridView5["项目", 2].Value = "内箱";
            dataGridView5["项目", 3].Value = "内卡";
            dataGridView5["项目", 4].Value = "辅材A";
            dataGridView5["项目", 5].Value = "辅材B";
            dataGridView5["项目", 6].Value = "辅材C";
            dataGridView5["项目", 7].Value = "辅材D";
            dataGridView5["项目", 8].Value = "辅材E";
            dataGridView5["项目", 9].Value = "辅材F";
            dgvc长.HeaderText = "长(mm)";
            dgvc宽.HeaderText = "宽(mm)";
            dgvc高.HeaderText = "高(mm)";
            dgvc序号.ReadOnly = true;
            bind_pack_material_again();
        }
        #endregion
        private void bind_pack_material_again()
        {
            //根据不同的客户类别加载不同的数据源 16/01/14
       
            for (i = 0; i < 3; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();
                c1.Items.Clear();
                c1.Items.Add("");
                c1.Items.Add("A式箱");
                c1.Items.Add("天地盖");
                dataGridView5["箱形", i] = c1;
            }
            dt = bc.getdt(cprint_pack_material.sql + string.Format(" WHERE A.PFID='{0}' ORDER BY A.PPKEY ASC", IDO));

            dtx = bc.getdt(string.Format(cpack_material.sql + @" WHERE 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            dtx = bc.GET_DT_TO_DV_TO_DT(dtx, "", "包装材质 IN ('ABF170','ABF140','ABF190','AF200','AF170','EF170')");
            if (dt.Rows.Count >= 4)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "包装材质", "材质", 0, 4);//此为正常写入数据库时 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "包装材质", "材质", 0, dt.Rows .Count );//此为写入数据库异常，数据没有写入数据库 160120
            }
            dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "VALUE");
            for (i = 0; i < 4; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();

                if (dtx.Rows.Count > 0)
                {
                    c1.Items.Clear();
                    c1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c1.Items.Add(dr["VALUE"].ToString());
                    }
                }
                dataGridView5["材质", i] = c1;
            }
            dtx = bc.getdt(string.Format(cpack_material.sql + @" WHERE 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            dtx = bc.GET_DT_TO_DV_TO_DT(dtx, "", "包装材质 NOT IN ('ABF170','ABF140','ABF190','AF200','AF170','EF170')");
            if (dt.Rows.Count >= 7)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "包装材质", "箱形", 4, 7);//此为正常写入数据库时 160120
            }
            else if (dt.Rows.Count >= 4)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "包装材质", "箱形", 4, dt.Rows.Count);//此为写入数据库异常，数据没有写入数据库 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "包装材质", "箱形", 0, dt.Rows.Count);//此为写入数据库异常，数据没有写入数据库 160120
            }
            dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "VALUE");
            for (i = 4; i < 7; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();
                if (dtx.Rows.Count > 0)
                {
                    c1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c1.Items.Add(dr["VALUE"].ToString());
                        //MessageBox.Show(dr["VALUE"].ToString());
                    }
                }
                dataGridView5["箱形", i] = c1;
            }

            //start

            if (dt.Rows.Count > 0)
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {

                    dataGridView5["项目", i].Value = dt.Rows[i]["项目"].ToString();
                    dataGridView5["数量", i].Value = dt.Rows[i]["数量"].ToString();
                    dataGridView5["长", i].Value = dt.Rows[i]["长"].ToString();
                    dataGridView5["宽", i].Value = dt.Rows[i]["宽"].ToString();
                    dataGridView5["高", i].Value = dt.Rows[i]["高"].ToString();
                    dataGridView5["箱形", i].Value = dt.Rows[i]["箱形"].ToString();
                    dataGridView5["材质", i].Value = dt.Rows[i]["材质"].ToString();
                }
                for (i = 7; i < dt.Rows.Count; i++)
                {
                    dataGridView5["单价", i].Value = dt.Rows[i]["单价"].ToString();//只有手输单价的行显示单价 16/01/12
                }
            }
            for (i = 0; i < 7; i++)
            {
                dataGridView5["单价", i].ReadOnly = true;//含组合框的行单价不只能调BOM，不能在录入时手动修改 16/01/12
            }
            dataGridView5["高", 3].ReadOnly = true;
            dataGridView5["箱形", 3].ReadOnly = true;
            dgvStateControl_dgv5();

        }
        #region  bind_artificial
        private void bind_artificial()
        {
            DataGridViewTextBoxColumn dgvc序号 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc项目 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc单价 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc数量 = new DataGridViewTextBoxColumn();
            dataGridView6.Columns.Add(dgvc序号);
            dataGridView6.Columns.Add(dgvc项目);
            dataGridView6.Columns.Add(dgvc单价);
            dataGridView6.Columns.Add(dgvc数量);
            dgvc序号.Name = "序号";
            dgvc项目.Name = "项目";
            dgvc数量.Name = "数量";
            dgvc单价.Name = "单价";

            for (int i = 0; i < 2; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView6.Rows.Add(dgvr);
                dataGridView6["序号", i].Value = i + 1;

            }
            dgvc序号.ReadOnly = true;
            bind_artificial_again();
        }
        #endregion
        #region  bind_artificial_again
        private void bind_artificial_again()
        {
            dtx = bc.getdt(string.Format(cartificial.sql + @" WHERE 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            dt = bc.getdt(cprint_artificial.sql + string.Format(" WHERE A.PFID='{0}'", IDO));
            if (dt.Rows.Count >= 1)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "纸品人工", "项目", 0, 1);//此为正常写入数据库时 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "纸品人工", "项目", 0, dt.Rows .Count );//此为写入数据库异常，数据没有写入数据库 160120
            }
            dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "VALUE");
            for (i = 0; i < 1; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();

                if (dtx.Rows.Count > 0)
                {
                    c1.Items.Clear();
                    c1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c1.Items.Add(dr["VALUE"].ToString());
                    }
                }
                dataGridView6["项目", i] = c1;
                dataGridView6["单价", i].ReadOnly = true;//含组合框的行单价不只能调BOM，不能在录入时手动修改 16/01/12
            }
            if (dt.Rows.Count > 0)
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    dataGridView6["项目", i].Value = dt.Rows[i]["项目"].ToString();
                    dataGridView6["数量", i].Value = dt.Rows[i]["数量"].ToString();

                }
                dataGridView6["单价", dt.Rows.Count - 1].Value = dt.Rows[dt.Rows.Count - 1]["单价"].ToString();//只有手输入单价的行显示单价

            }
            dgvStateControl_dgv6();
        }
        #endregion
        #region  bind_purchase
        private void bind_purchase()
        {
            DataGridViewTextBoxColumn dgvc序号 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc类型一 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc外购价一 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc类型二 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc外购价二 = new DataGridViewTextBoxColumn();
            dataGridView7.Columns.Add(dgvc序号);
            dataGridView7.Columns.Add(dgvc类型一);
            dataGridView7.Columns.Add(dgvc外购价一);
            dataGridView7.Columns.Add(dgvc类型二);
            dataGridView7.Columns.Add(dgvc外购价二);
            dgvc序号.Name = "序号";
            dgvc类型一.Name = "类型一";
            dgvc类型二.Name = "类型二";
            dgvc外购价一.Name = "外购价一";
            dgvc外购价二.Name = "外购价二";
            dataGridView7.Columns["序号"].Width = 40;
            dataGridView7.Columns["类型一"].Width = 70;
            dataGridView7.Columns["类型二"].Width = 70;
            dataGridView7.Columns["外购价一"].Width = 50;
            dataGridView7.Columns["外购价二"].Width = 50;
      
            for (int i = 0; i < 2; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView7.Rows.Add(dgvr);
                dataGridView7["序号", i].Value = i + 1;
            }
       
            dgvc序号.ReadOnly = true;
            /*绑定第一行数据源 16/01/13 start*/
            dt = bc.getdt(cprint_purchase.sql + string.Format(" WHERE A.PFID='{0}'", IDO));
            dtx = bc.getdt(cpurchase.sql);
            if (dt.Rows.Count >= 1)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "外购件类型", "类型一", 0, 1);//此为正常写入数据库时 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "外购件类型", "类型一", 0, dt.Rows .Count );//此为写入数据库异常，数据没有写入数据库 160120
            }
            for (i = 0; i <1 ; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();
                c1.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {
                    c1.Items.Add(dr["VALUE"].ToString());
                }
                dataGridView7["类型一", i] = c1;
                dtx = bc.getdt(cpurchase.sql);
                if (dt.Rows.Count >= 1)
                {
                    dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "外购件类型", "类型二", 0, 1);//此为正常写入数据库时 160120
                }
                else
                {
                    dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "外购件类型", "类型二", 0, dt.Rows.Count);//此为写入数据库异常，数据没有写入数据库 160120
                }
                if (dtx.Rows.Count > 0)
                {
                    DataGridViewComboBoxCell c2 = new DataGridViewComboBoxCell();
                    c2.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c2.Items.Add(dr["VALUE"].ToString());
                    }
                    dataGridView7["类型二", i] = c2;
                }
            }
            /*绑定第一行数据源 16/01/13 end*/
            /*绑定第二行数据源 16/01/13 start*/
            for (i = 1; i < 2; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();
                c1.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {
                    c1.Items.Add(dr["VALUE"].ToString());
                }
                dataGridView7["类型一", i] = c1;
            }
            /*绑定第二行数据源 16/01/13 end*/
            dgvc类型一.HeaderText = "类型";
            dgvc类型二.HeaderText = "类型";
            dgvc外购价一.HeaderText = "外购价";
            dgvc外购价二.HeaderText = "外购价";
            dataGridView7["类型二", 1].ReadOnly = true;
            dataGridView7["外购价二", 1].ReadOnly = true;
      
            if (dt.Rows.Count > 0)
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    dataGridView7["类型一", i].Value = dt.Rows[i]["类型一"].ToString();
                    dataGridView7["类型二", i].Value = dt.Rows[i]["类型二"].ToString();
                    dataGridView7["外购价一", i].Value = dt.Rows[i]["外购价一"].ToString();
                    dataGridView7["外购价二", i].Value = dt.Rows[i]["外购价二"].ToString();
                }

            }
      
            dgvStateControl_dgv7();
        }
        #endregion
        #region  bind_transport
        private void bind_transport()
        {
            DataGridViewTextBoxColumn dgvc长 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc宽 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc高 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc总箱数 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc总立方数 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc运输方式 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn dgvc单价 = new DataGridViewTextBoxColumn();
            //DataGridViewTextBoxColumn dgvc单价 = new DataGridViewTextBoxColumn();
            dataGridView8.Columns.Add(dgvc长);
            dataGridView8.Columns.Add(dgvc宽);
            dataGridView8.Columns.Add(dgvc高);
            dataGridView8.Columns.Add(dgvc总箱数);
            dataGridView8.Columns.Add(dgvc总立方数);
            dataGridView8.Columns.Add(dgvc运输方式);
            dataGridView8.Columns.Add(dgvc单价);
            //dataGridView8.Columns.Add(dgvc单价);
            dgvc长.Name = "长";
            dgvc宽.Name = "宽";
            dgvc高.Name = "高";
            dgvc总箱数.Name = "总箱数";
            dgvc总立方数.Name = "总立方数";
            dgvc运输方式.Name = "运输方式";
            dgvc单价.Name = "单价";

            dgvc长.ReadOnly = true;
            dgvc宽.ReadOnly = true;
            dgvc高.ReadOnly = true;
            for (int i = 0; i < 3; i++)
            {
                DataGridViewRow dgvr = new DataGridViewRow();
                dataGridView8.Rows.Add(dgvr);
            }
            dataGridView8["总箱数", 0].ReadOnly = true;
            dataGridView8["总箱数", 1].ReadOnly = true;
            bind_transport_again();
        }
        #endregion
        #region  bind_transport_again
        private void bind_transport_again()
        {
           

            dt = bc.getdt(cprint_transport.sql + string.Format(" WHERE A.PFID='{0}'", IDO));
            dtx = bc.getdt(string.Format(ctrasport.sql + @" WHERE 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            if (dt.Rows.Count >= 3)
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "物流运输", "运输方式", 0, 3);//此为正常写入数据库时 160120
            }
            else
            {
                dtx = bc.RETURN_NATURE_AND_NOW_DT(dtx, dt, "物流运输", "运输方式", 0, dt.Rows .Count );//此为写入数据库异常，数据没有写入数据库 160120
            }
            for (i = 0; i < 2; i++)
            {
                DataGridViewComboBoxCell c1 = new DataGridViewComboBoxCell();

                if (dtx.Rows.Count > 0)
                {
                    c1.Items.Clear();
                    c1.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        c1.Items.Add(dr["VALUE"].ToString());
                    }
                }
                dataGridView8["运输方式", i] = c1;
                dgv8(i, false);
            }

            if (dt.Rows.Count > 0 && dt.Rows .Count >=5)
            {
                for (i = 0; i < dt.Rows.Count - 2; i++)
                {
                    dataGridView8["总箱数", i].Value = dt.Rows[i]["总箱数"].ToString();
                    dataGridView8["运输方式", i].Value = dt.Rows[i]["运输方式"].ToString();
                    dataGridView8["总立方数", i].Value = dt.Rows[i]["总立方数"].ToString();

                }
                dataGridView8["单价", dt.Rows.Count - 2].Value = dt.Rows[dt.Rows.Count - 2]["单价"].ToString();//只有手输单价的行显示单价 16/01/12
            }
            dataGridView8["总立方数", 0].ReadOnly = true;
            dataGridView8["总立方数", 1].ReadOnly = true;
            dataGridView8["单价", 0].ReadOnly = true;
            dataGridView8["单价", 1].ReadOnly = true;
            dgvStateControl_dgv8();
        }
        #endregion
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("纸品报价新增", LOGIN.USID);
            btnAdd.Visible = false;
            btnSave.Visible = false;
            label15.Visible = false;
            label17.Visible = false;
            pictureBox1.Visible = false;
            label1.Visible = false;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            btnDel.Visible = false;
            label13.Visible = false;
            if (dtx.Rows.Count > 0)
            {
                if (dtx.Rows[0]["新增权限"].ToString() == "有权限")
                {
                    btnAdd.Visible = true;
                    label17.Visible = true;
                    btnSave.Visible = true;
                    label15.Visible = true;
        
                }
                if (dtx.Rows[0]["报价审核"].ToString() == "有权限")
                {
                    pictureBox1.Visible = true;
                    label1.Visible = true;
                }
                if (dtx.Rows[0]["修改权限"].ToString() == "有权限")
                {
                    btnSave.Visible = true;
                    label15.Visible = true;
                    EDIT = "有权限";
                }
                if (dtx.Rows[0]["删除权限"].ToString() == "有权限")
                {
                    btnDel.Visible = true;
                    label13.Visible = true;
                }
                if (dtx.Rows[0]["基本信息_采购"].ToString() == "有权限")
                {
                    button1.Visible = true;
                }
                if (dtx.Rows[0]["估计计算表"].ToString() == "有权限")
                {
                    button2.Visible = true;
                }
                if (dtx.Rows[0]["预算明细表"].ToString() == "有权限")
                {
                    button3.Visible = true;
                }
                if (dtx.Rows[0]["基本信息_AE"].ToString() == "有权限")
                {
                    button4.Visible = true;
                }
                if (dtx.Rows[0]["主件明细表"].ToString() == "有权限")
                {
                    button5.Visible = true;
                }
                if (dtx.Rows[0]["产品报价单"].ToString() == "有权限")
                {
                    button6.Visible = true;
                }
                if (dtx.Rows[0]["明细报价单"].ToString() == "有权限")
                {
                    button7.Visible = true;
                }

            }

        }
        #endregion
        #region AUDIT
        private void AUDIT()
        {
            DataTable dtx = basec.getdts(cprinting_offer.sql + " where A.PFID='" + IDO + "' ORDER BY  B.PFID ASC ");
            if (dtx.Rows.Count > 0)
            {
                if (dtx.Rows[0]["审核状态"].ToString() == "已审核")
                {
                    pictureBox1.Image = Resource1.Audit;
                    label1.Text = "已审核";
                }
                else
                {
                    pictureBox1.Image = Resource1._61;
                    label1.Text = "未审核";
                }

            }
        }
        #endregion
        public void ClearText()
        {
            comboBox1.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {

            try
            {
                bind();
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                add();
                hint.Text = "";
            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
            M_int_judge = 1;
        }
        void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            progressBar1.Maximum = 600;
            for (int i = 0; i <= 600; i++)
            {
                if (IFExecution_SUCCESS)
                {
                    progressBar1.Value = progressBar1.Maximum;
                    break;
                }
                else
                {
                    progressBar1.Value = i;
                    System.Threading.Thread.Sleep(100);//线程开始后激发该事件,在此事件里处理进度条显示效果 16/01/20
                }
                
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            IFExecution_SUCCESS = false;
            progressBar1.Value = 0;//初始化进度条 16/01/20
            hint.Text = "";
            btnSave.Focus();
            if (juage())
            {

            }
            else
            {
                backgroundWorker1.RunWorkerAsync();//线程开始开始运行 16/01/20
                backgroundWorker1.WorkerReportsProgress = true;//允许使用线程进度  16/01/20
                backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);//线程开始后激发该事件,在此事件里处理进度条显示效果 16/01/20
                save();
            }
            try
            {
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void add()
        {
            ClearText();
            IDO = cprinting_offer.GETID();
            bind();
            ADD_OR_UPDATE = "ADD";
            this.Close();
            PRINTING_OFFERT FRM = new PRINTING_OFFERT();
            FRM.IDO = cprinting_offer.GETID();
            FRM.Show();
        }
        #region save
        private void save()
        {

            btnSave.Focus();
            //dgvfoucs();
            cprinting_offer.MAKERID = LOGIN.EMID;
            OFFER_ID = textBox2.Text;
            string v1 = bc.getOnlyString("SELECT CHARGE_AUDIT_STATUS FROM PRINTING_OFFER_MST WHERE PFID='" + IDO + "'");
            if (!bc.exists(cprinting_offer.sql + " WHERE B.PFID='" + IDO + "'"))
            {

                cprinting_offer.PFID = IDO;
                cprinting_offer.PIID = bc.getOnlyString("SELECT PIID FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'"); ;
                cprinting_offer.COUNT = decimal.Parse(textBox1.Text);
                cprinting_offer.GETID_OFFER_ID("", comboBox2.Text);
                cprinting_offer.OFFER_ID = cprinting_offer.OFFER_ID;
                textBox2.Text = cprinting_offer.OFFER_ID;
                cprinting_offer.CHARGE_AUDIT_STATUS = "N";
                cprinting_offer.save(dataGridView1, comboBox2.Text);
            }
            else if (v1 != "Y")
            {
                cprinting_offer.PFID = IDO;
                cprinting_offer.OFFER_ID = textBox2.Text;
                cprinting_offer.CHARGE_AUDIT_STATUS = "N";
                cprinting_offer.PIID = bc.getOnlyString("SELECT PIID FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'"); ;
                cprinting_offer.COUNT = decimal.Parse(textBox1.Text);
                cprinting_offer.save(dataGridView1, comboBox2.Text);
            }
            else
            {
                IDO = cprinting_offer.GETID();
                cprinting_offer.PFID = IDO;
                cprinting_offer.PIID = bc.getOnlyString("SELECT PIID FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'"); ;
                cprinting_offer.COUNT = decimal.Parse(textBox1.Text);
                SAMPLE_CODE_FIRST = textBox2.Text.Substring(textBox2.Text.Length - 1, 1);
                cprinting_offer.GETID_OFFER_ID("", comboBox2.Text);
                cprinting_offer.OFFER_ID = cprinting_offer.OFFER_ID;
                textBox2.Text = cprinting_offer.OFFER_ID;
                cprinting_offer.CHARGE_AUDIT_STATUS = "N";
                cprinting_offer.save(dataGridView1, comboBox2.Text);
                AUDIT();
            }
          
            dtt = bc.getdt(cprinting_offer.sql + " WHERE B.PFID='" +IDO + "'");
            if (dtt.Rows.Count > 0)
            {
                dtt = cprinting_offer.RETURN_DT(dtt);
                dtt = cprinting_offer.bind2(dtt, 0, "");
                dtt_hide = cprinting_offer.RETURN_DT_SHOW_HIDE(dtt);
                dtt = dtt_hide;
            }
            /*保存主件汇总表 start*/
            if (dtt_hide.Rows.Count > 0)
            {
                cprinting_offer.MAKERID = LOGIN.EMID;
                cprinting_offer.PFID = IDO;
                cprinting_offer.save_print_total(dtt_hide);
            }
            /*保存主件汇总表 end*/
            /*dgv2*/
            dt2 = cprinting_offer.RETURN_DIE_CUTTING_PRICE_DT(dtt,dataGridView2);
            cprint_die_cutting.PFID = IDO;
            cprint_die_cutting.MAKERID = LOGIN.EMID;
         
            cprint_die_cutting.save(dt2);
          
            /*dgv2*/

            /*dgv3*/
            dt3 = cprinting_offer.RETURN_PORTRAY_DT(dtt, dataGridView3);
            cprint_portray.PFID = IDO;
            cprint_portray.MAKERID = LOGIN.EMID;
            cprint_portray.save(dt3);
       
            /*dgv3*/

            /*dgv4*/
        
            cprint_parts_auxiliary.PFID = IDO;
            cprint_parts_auxiliary.MAKERID = LOGIN.EMID;
            cprint_parts_auxiliary.save(dataGridView4);
            /*dgv4*/
            /*dgv5*/
            cprint_pack_material.PFID = IDO;
            cprint_pack_material.PROJECT_ID = comboBox1.Text;
            cprint_pack_material.MAKERID = LOGIN.EMID;
            cprint_pack_material.save(dataGridView5);
            /*dgv5*/
            /*dgv6*/
            cprint_artificial.PFID = IDO;
            cprint_artificial.MAKERID = LOGIN.EMID;
            cprint_artificial.PROJECT_ID = comboBox1.Text;
            cprint_artificial.save(dataGridView6);
            /*dgv6*/
            /*dgv7*/
            dt7 = cprinting_offer.RETURN_PURCHASE_DT(dtt, dataGridView7);
            cprint_purchase.PFID = IDO;
            cprint_purchase.MAKERID = LOGIN.EMID;
            cprint_purchase.save(dt7);
            /*dgv7*/
            /*dgv8*/
      
            dt8 = cprinting_offer.RETURN_TRANSPORT_DT(dtt, dataGridView8);
            //MessageBox.Show("7");
            cprint_transport.PFID = IDO;
            cprint_transport.MAKERID = LOGIN.EMID;
        
            cprint_transport.save(dt8);
            //MessageBox.Show("8");
            /*dgv8*/
        
            /*保存费用汇总表获取查询页的报出价 start*/
            dtx = cprinting_offer.RETURN_COST_TOTAL_DT(IDO, dtt);
            if (dtx.Rows.Count > 0)
            {
                //MessageBox.Show("5");
                cprint_cost_total.MAKERID = LOGIN.EMID;
                cprint_cost_total.PFID = IDO;
                cprint_cost_total.save(dtx);
            }
    
            /*保存费用汇总表获取查询页的报出价 end*/

            IFExecution_SUCCESS = cprinting_offer.IFExecution_SUCCESS;
            hint.Text = cprinting_offer.ErrowInfo;
            if (IFExecution_SUCCESS)
            {
                if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
                {

                    hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
                }
                else
                {
                    hint.Text = "";
                }
                F1.load();
             
            }
        }
        #endregion
        #region juage
        private bool juage()
        {
           
            bool b = false;
            if (bc.exists(cprinting_offer.sql + " WHERE A.PFID='" + IDO + "'") && EDIT != "有权限")
            {
                hint.Text = "本账号无修改权限！";
                b = true;
            }
            else if (IDO == "")
            {
                hint.Text = "编号不能为空！";
                b = true;
            }
            else if (comboBox1.Text == "")
            {
                hint.Text = "项目号不能为空！";
                b = true;
            }
            else if (!bc.exists(cproject_info.sql + " WHERE A.PROJECT_ID='" + comboBox1.Text + "'"))
            {

                hint.Text = "项目号不存在系统！";
                b = true;
            }
            else if (comboBox2.Text == "")
            {
                hint.Text = "报价类别不能为空！";
                b = true;
            }
            else if (juage3())
            {
                hint.Text = "报价类别不存在！";
                b = true;

            }
            else if (textBox1.Text == "")
            {
                hint.Text = "数量不能为空！";
                b = true;
            }
            else if (bc.yesno(textBox1.Text) == 0)
            {
                hint.Text = "数量只能输入数字！";
                b = true;
            }
            else if (decimal.Parse(textBox1.Text) == 0)
            {
                hint.Text = "数量需大于0！";
                b = true;
            }

            else if (juage2())
            {
                b = true;
            }
            else if (RETURN_ERROW())//juage bind2
            {

                b = true;

            }
            else if (juage_dgv2())
            {
                b = true;
            }
            else if (juage_dgv3())
            {
                b = true;
            }
            else if (juage_dgv4())
            {
                b = true;
            }
            else if (juage_dgv6())
            {
                b = true;
            }
            else if (juage_dgv7())
            {

                b = true;
            }
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE PFID='{0}'",bc.RETURN_PFID(textBox2 .Text ))))
            {
                hint.Text = string.Format("报价 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
      
            return b;
        }
        #endregion
        #region juage2()
        private bool juage2()
        {

            bool b = false;
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {

                if (JUAGE_WNAME_IF_ABOVE_ONE(dataGridView1, "部品名") == false)
                {
                    hint.Text = string.Format("至少有一项部品才能保存");
                    b = true;
                    break;
                }
                else if (dataGridView1["部品名", i].FormattedValue.ToString() == "")
                {

                }
                else if (dataGridView1["图纸门幅", i].FormattedValue.ToString() == "")
                {

                    hint.Text = string.Format("项次 {0} 图纸门幅不能为空", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView1["图纸门幅", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 图纸门幅只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if (dataGridView1["图纸纸长", i].FormattedValue.ToString() == "")
                {
                    hint.Text = string.Format("项次 {0} 图纸纸长不能为空", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;

                    break;
                }
                else if (bc.yesno(dataGridView1["图纸纸长", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 图纸纸长只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if (dataGridView1["部品个数", i].FormattedValue.ToString() == "")
                {
                    hint.Text = string.Format("项次 {0} 部品个数不能为空", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;

                    break;
                }
                else if (bc.yesno(dataGridView1["部品个数", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 部品个数只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;

                    break;
                }
                else if (dataGridView1["拼模数", i].FormattedValue.ToString() == "")
                {
                    hint.Text = string.Format("项次 {0} 拼模数不能为空", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;

                    break;
                }
                else if (bc.yesno(dataGridView1["拼模数", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 拼模数只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;

                    break;
                }
                else if (bc.yesno(dataGridView1["面纸克重", i].FormattedValue.ToString()) == 0)
                {

                    hint.Text = string.Format("项次 {0} 面纸克重只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView1["底纸克重", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 底纸克重只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView1["表面次数", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 表面次数只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView1["裱纸次数", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 裱纸次数只能输入数字", dataGridView1["项次", i].FormattedValue.ToString());
                    b = true;
                    break;
                }
                else if ((dataGridView1["印刷选项", i].FormattedValue.ToString() != "单纸双异" &&
                    dataGridView1["印刷选项", i].FormattedValue.ToString() != "双纸画异") && dataGridView1["反面4C", i].FormattedValue.ToString() != "")
                {

                    MessageBox.Show(string.Format("项次 {0} 印刷选项为单纸双异或是双纸画异时 反面4C栏位才可输入数据",
                        dataGridView1["项次", i].FormattedValue.ToString()), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    /*hint.Text = string.Format("项次 {0} 印刷选项为单纸双异或是双纸画异时 反面4C栏位才可输入数据", 
                    dataGridView1["项次", i].FormattedValue.ToString());*/
                    b = true;
                    break;


                }
                else if ((dataGridView1["印刷选项", i].FormattedValue.ToString() != "单纸双异" &&
                    dataGridView1["印刷选项", i].FormattedValue.ToString() != "双纸画异") && dataGridView1["反面专色", i].FormattedValue.ToString() != "")
                {
                    MessageBox.Show(string.Format("项次 {0} 印刷选项为单纸双异或是双纸画异时  反面专色栏位才可输入数据",
                    dataGridView1["项次", i].FormattedValue.ToString()), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    /* hint.Text = string.Format("项次 {0} 印刷选项为单纸双异或是双纸画异时  反面专色栏位才可输入数据",
                         dataGridView1["项次", i].FormattedValue.ToString());*/
                    b = true;
                    break;


                }
                else if ((dataGridView1["印刷选项", i].FormattedValue.ToString() != "单纸双异" &&
                    dataGridView1["印刷选项", i].FormattedValue.ToString() != "双纸画异") && dataGridView1["反面防晒", i].FormattedValue.ToString() != "")
                {
                    MessageBox.Show(string.Format("项次 {0} 印刷选项为单纸双异或是双纸画异时 反面防晒栏位才可输入数据",
                   dataGridView1["项次", i].FormattedValue.ToString()), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    /*hint.Text = string.Format("项次 {0} 印刷选项为单纸双异或是双纸画异时 反面防晒栏位才可输入数据", 
                        dataGridView1["项次", i].FormattedValue.ToString());*/
                    b = true;
                    break;
                }
                else if ((dataGridView1["印刷选项", i].FormattedValue.ToString() == "单纸双同" ||
                    dataGridView1["印刷选项", i].FormattedValue.ToString() == "单纸双异") &&
                    (dataGridView1["芯纸", i].FormattedValue.ToString() != "" || dataGridView1["芯纸规格", i].FormattedValue.ToString() != "" ||
                    dataGridView1["底纸", i].FormattedValue.ToString() != "" || dataGridView1["底纸克重", i].FormattedValue.ToString() != ""))
                {
                    MessageBox.Show(string.Format("项次 {0} 印刷选项不为单纸双同或是单纸双异时 芯纸 芯纸规格 底纸 底纸克重栏位才可输入数据",
               dataGridView1["项次", i].FormattedValue.ToString()), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    /*hint.Text = string.Format("项次 {0} 印刷选项不为单纸双同或是单纸双异时 芯纸 芯纸规格 底纸 底纸克重栏位才可输入数据",
                        dataGridView1["项次", i].FormattedValue.ToString());*/
                    b = true;
                    break;
                }


                else if ((dataGridView1["印刷选项", i].FormattedValue.ToString() != "双纸画异" &&
                    dataGridView1["印刷选项", i].FormattedValue.ToString() != "单纸双异") &&
                    dataGridView1["双面印刷", i].FormattedValue.ToString() != "")
                {
                    MessageBox.Show(string.Format("项次 {0} 印刷选项为双纸画异或单纸双异时 双面印刷栏位才可输入数据",
               dataGridView1["项次", i].FormattedValue.ToString()), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    /*hint.Text = string.Format("项次 {0} 印刷选项不为单纸双同或是单纸双异时 芯纸 芯纸规格 底纸 底纸克重栏位才可输入数据",
                        dataGridView1["项次", i].FormattedValue.ToString());*/
                    b = true;
                    break;
                }

            }



            return b;
        }
        #endregion
        #region juage3
        private bool juage3()
        {

            List<string> list = new List<string>();
            list.Add("Z");
            list.Add("M");
            list.Add("Y");
            list.Add("J");
            list.Add("F");
            DataTable dtt = new DataTable();
            dtt.Columns.Add("VALUE", typeof(string));
            for (int i = 0; i < list.Count; i++)
            {
                DataRow dr = dtt.NewRow();
                dr["VALUE"] = list[i];
                dtt.Rows.Add(dr);
            }
            dtt = bc.GET_DT_TO_DV_TO_DT(dtt, "", "VALUE='" + comboBox2.Text + "'");
            bool b = false;
            if (dtt.Rows.Count > 0)
            {

            }
            else
            {
                b = true;
            }
            return b;
        }
        #endregion
        #region juage_dgv2()

        private bool juage_dgv2()
        {
            bool b = false;

            for (i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {

                if (bc.yesno(dataGridView2["刀模长米", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 刀模长米只能输入数字", (i + 1).ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView2["元米", 1].FormattedValue.ToString()) == 0)//元米只需判断第二行就好，第一行是不显示的，只是调用 16/01/13
                {
                    hint.Text = string.Format("项次 {0} 元/米只能输入数字", (i + 1).ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView2["圆孔个数", i].FormattedValue.ToString()) == 0)
                {
                    hint.Text = string.Format("项次 {0} 圆孔个数只能输入数字", (i + 1).ToString());
                    b = true;
                    break;
                }
                else if (bc.yesno(dataGridView2["元个", 1].FormattedValue.ToString()) == 0)//元米只需判断第二行就好，第一行是不显示的，只是调用 16/01/13
                {
                    hint.Text = string.Format("项次 {0} 元/个只能输入数字", (i + 1).ToString());
                    b = true;
                    break;
                }
                else if (dataGridView2["项目", i].FormattedValue.ToString() != "" && dataGridView2["元米", 2].FormattedValue.ToString() == "按平方")
                {
                    hint.Text = string.Format("刀模计价使用项目时需选择按米计", "");
                    b = true;
                    break;
                }
            }
            return b;
        }
        #endregion
        #region juage_dgv3()

        private bool juage_dgv3()
        {

            bool b = false;
            for (i = 0; i < 7; i++)
            {
                if (dataGridView3[0, i].FormattedValue.ToString() == "")
                {

                }
                else if (dataGridView3["长", i].FormattedValue.ToString() == "" || dataGridView3["宽", i].FormattedValue.ToString() == "" ||
                    dataGridView3["总数量", i].FormattedValue.ToString() == "" || juage_dgv3_IF_HAVE_PRICE(dataGridView3[0, i].FormattedValue.ToString(),i+1))
                {
                    b = true;
                    hint.Text = string.Format("第 {0} 行写真类型不为空时需输入长 宽 总数量 且属性管理里要存在预设的单价", (i + 1).ToString());
                    break;

                }
            }
            if (b)
            {
            }
            else
            {
            
                for (i = 7; i < dataGridView3.Rows.Count; i++)
                {
                    if (dataGridView3[0, i].FormattedValue.ToString() == "")
                    {

                    }
                    else if (dataGridView3["长", i].FormattedValue.ToString() == "" || dataGridView3["宽", i].FormattedValue.ToString() == "" ||
                        dataGridView3["总数量", i].FormattedValue.ToString() == "" || dataGridView3[0, i].FormattedValue.ToString() == "")
                    {
                        b = true;
                        hint.Text = string.Format("第 {0} 行写真类型不为空时需输入长 宽 总数量 单价", (i + 1).ToString());
                        break;

                    }
                }
            }
            return b;
        }
        #endregion
        #region juage_dgv4()
        private bool juage_dgv4()
        {
            bool b = false;
            for (i = 0; i < 8; i++)
            {
                if (dataGridView4["配件名", i].FormattedValue.ToString() == "")
                {

                }
                else if (dataGridView4["用量", i].FormattedValue.ToString() == "" || dataGridView4["单位", i].FormattedValue.ToString() == "" || 
                    juage_dgv4_IF_HAVE_PRICE(dataGridView4["配件名", i].FormattedValue.ToString(), i + 1) )
                {
                    b = true;
                    hint.Text = string.Format("序号 {0} 配件名不为空时需输入用量 单价 单位", (i + 1).ToString());
                    break;

                }
            }
            for (i = 8; i < dataGridView4.Rows.Count; i++)
            {
                if (dataGridView4["配件名", i].FormattedValue.ToString() == "")
                {

                }
                else if (dataGridView4["用量", i].FormattedValue.ToString() == "" || dataGridView4["单价", i].FormattedValue.ToString() == "" ||
                    dataGridView4["单位", i].FormattedValue.ToString() == "")
                {
                    b = true;
                    hint.Text = string.Format("序号 {0} 配件名不为空时需输入用量 单价 单位", (i + 1).ToString());
                    break;

                }
            }
            return b;
        }
        #endregion
        #region juage_dgv6()
        private bool juage_dgv6()
        {
            bool b = false;
            for (i = 0; i < 1; i++)
            {
                if (dataGridView6["项目", i].FormattedValue.ToString() == "")
                {

                }
                else if (dataGridView6["数量", i].FormattedValue.ToString() == "" || 
                    juage_dgv6_IF_HAVE_PRICE(dataGridView6["项目", i].FormattedValue.ToString(), i + 1))
                {
                    b = true;
                    hint.Text = string.Format("人工 序号 {0} 项目不为空时需输入数单价 数量", (i + 1).ToString());
                    break;

                }
            }
            for (i = 1; i < dataGridView6.Rows.Count; i++)
            {
                if (dataGridView6["项目", i].FormattedValue.ToString() == "")
                {

                }
                else if (dataGridView6["数量", i].FormattedValue.ToString() == "" || dataGridView6["单价", i].FormattedValue.ToString() == "")
                {
                    b = true;
                    hint.Text = string.Format("人工 序号 {0} 项目不为空时需输入数单价 数量", (i + 1).ToString());
                    break;

                }
            }
            return b;
        }
        #endregion
        #region juage_dgv7()
        private bool juage_dgv7()
        {
            bool b = false;
            if (dataGridView7.Rows.Count >= 2)
            {
                for (i = 0; i < 1; i++)
                {
                    if (dataGridView7["类型一", i].FormattedValue.ToString() == "")
                    {

                    }
                    else if (dataGridView7["外购价一", i].FormattedValue.ToString() == "")
                    {
                        b = true;
                        hint.Text = string.Format("序号 {0} 类型不为空时需输入外购价", (i + 1).ToString());
                        break;

                    }
                    else if (dataGridView7["类型二", i].FormattedValue.ToString() == "")
                    {

                    }
                    else if (dataGridView7["外购价二", i].FormattedValue.ToString() == "")
                    {
                        b = true;
                        hint.Text = string.Format("序号 {0} 类型不为空时需输入外购价", (i + 1).ToString());
                        break;

                    }
                }
                for (i = 1; i < 2; i++)
                {
                    if (dataGridView7["类型一", i].FormattedValue.ToString() == "")
                    {

                    }
                    else if (dataGridView7["外购价一", i].FormattedValue.ToString() == "")
                    {
                        b = true;
                        hint.Text = string.Format("序号 {0} 类型不为空时需输入外购价", (i + 1).ToString());
                        break;

                    }

                }
            }
            return b;
        }
        #endregion
        #region juage_dgv3_IF_HAVE_PRICE()
        private bool juage_dgv3_IF_HAVE_PRICE(string VALUE,int row)
        {
            bool b = false;
            dtx = bc.getdt(cportray.sql + string.Format(@" WHERE A.PORTRAY_TYPE='{0}' AND 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{1}'", VALUE , CUSTOMER_TYPE));
            if (dtx.Rows.Count > 0)
            {
                if (string.IsNullOrEmpty(VALUE ))
                {
                    hint.Text = string.Format("第 {0} 行预先设好的单价存在空值", (row).ToString());
                    b = true;
                }
            }
            else
            {
                hint.Text = string.Format("第 {0} 行不存在预先设好的单价", (row).ToString());
                b = true;

            }
            return b;
        }
        #endregion
        #region juage_dgv4_IF_HAVE_PRICE()
        private bool juage_dgv4_IF_HAVE_PRICE(string VALUE, int row)//判断是否存在预先设好的单价
        {
            bool b = false;
   
            dtx = bc.getdt(cparts_auxiliary.sql + string.Format(" WHERE A.PARTS_AUXILIARY='{0}'",
                      VALUE ));
            if (dtx.Rows.Count > 0)
            {
                if (string.IsNullOrEmpty(VALUE))
                {
                    hint.Text = string.Format("第 {0} 行预先设好的单价存在空值", (row).ToString());
                    b = true;
                }
            }
            else
            {
                hint.Text = string.Format("第 {0} 行不存在预先设好的单价", (row).ToString());
                b = true;

            }
            return b;
        }
        #endregion
        #region juage_dgv6_IF_HAVE_PRICE()
        private bool juage_dgv6_IF_HAVE_PRICE(string VALUE, int row)//判断是否存在预先设好的单价
        {
            bool b = false;
            dtx = bc.getdt(cartificial.sql + string.Format(" WHERE A.ARTIFICIAL='{0}' AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{1}'",
                        VALUE, bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
            if (dtx.Rows.Count > 0)
            {
                if (string.IsNullOrEmpty(VALUE))
                {
                    hint.Text = string.Format("第 {0} 行预先设好的单价存在空值", (row).ToString());
                    b = true;
                }
            }
            else
            {
                hint.Text = string.Format("第 {0} 行不存在预先设好的单价", (row).ToString());
                b = true;

            }
            return b;
        }
        #endregion
        #region JUAGE_WNAME_IF_ABOVE_ONE
        private bool JUAGE_WNAME_IF_ABOVE_ONE(DataGridView dgv, string COLUMN_NAME)
        {
            bool b = false;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv[COLUMN_NAME, i].FormattedValue.ToString() != "")
                {
                    b = true;
                }
            }
            return b;
        }
        #endregion
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
            dataGridView1.ClearSelection();//加载不选中第一列
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
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            }


            //dataGridView1.Columns["站别代码"].DefaultCellStyle.BackColor = Color.Yellow;

            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["部品数"].ReadOnly = true;
            dataGridView1.Columns["项次"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
          
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
            dataGridView2.Columns["项目"].Width = 80;
            dataGridView2.Columns["刀模长米"].Width = 60;
            dataGridView2.Columns["元米"].Width = 70;
            dataGridView2.Columns["圆孔个数"].Width = 40;
            dataGridView2.Columns["元个"].Width = 50;
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
            for (i = 0; i < numCols1; i++)
            {
                dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView3.EnableHeadersVisualStyles = false;
                dataGridView3.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView3.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
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
            /*dataGridView4.Columns["序号"].Width = 40;
            dataGridView4.Columns["配件名"].Width = 100;
            dataGridView4.Columns["用量"].Width = 40;
            dataGridView4.Columns["单价"].Width = 40;
            dataGridView4.Columns["单位"].Width = 50;
            dataGridView4.Columns["备注"].Width = 80;*/


            for (i = 0; i < numCols1; i++)
            {

                dataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView4.EnableHeadersVisualStyles = false;
                dataGridView4.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView4.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;


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
            dataGridView5.Columns["序号"].Width = 40;
            dataGridView5.Columns["项目"].Width = 60;
            dataGridView5.Columns["数量"].Width = 40;
            dataGridView5.Columns["长"].Width = 40;
            dataGridView5.Columns["宽"].Width = 40;
            dataGridView5.Columns["高"].Width = 40;
            dataGridView5.Columns["箱形"].Width = 100;
            dataGridView5.Columns["材质"].Width = 70;
            dataGridView5.Columns["单价"].Width = 50;
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
            //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/

            for (i = 0; i < numCols1; i++)
            {
                dataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView6.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView6.EnableHeadersVisualStyles = false;
                dataGridView6.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView6.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
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
            dataGridView6.Columns["序号"].Width = 40;
            dataGridView6.Columns["项目"].Width = 70;
            dataGridView6.Columns["数量"].Width = 40;
            dataGridView6.Columns["单价"].Width = 50;
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
            };


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
            };


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


        }
        #endregion
        #region dataGridView1_CellValidating
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                int columnsindex = dataGridView1.CurrentCell.ColumnIndex;

                if (dataGridView1.Columns[columnsindex].Name == "图纸门幅" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView1.Columns[columnsindex].Name == "图纸纸长" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (dataGridView1.Columns[columnsindex].Name == "部品个数" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (dataGridView1.Columns[columnsindex].Name == "拼模数" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (dataGridView1.Columns[columnsindex].Name == "部品数" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (dataGridView1.Columns[columnsindex].Name == "面纸克重" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (dataGridView1.Columns[columnsindex].Name == "底纸克重" && bc.yesno(e.FormattedValue.ToString()) == 0)
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
        #region RETURN_ERROW
        private bool RETURN_ERROW()  //JUAGE KEY IN INFOMATION
        {
            //本客户未设置客户类别 或 项目号没有设置品牌 1909
            //ErrowInfo = string.Format("项次 {0} 面纸暂无", i + 1); 3034
            //判断芯纸为PVC或PET加工门幅不能超915在NO.3632行 16/01/06
            //判断芯纸为KT板时加工门幅不能超1200在NO.3642行 16/01/06
            //判断芯纸为双灰板时加工门幅及加工纸长不能超过889X1194  3654，3659 16/01/06
            //判断芯纸为AD板时加工门幅及加工纸长不能超过1220X2440   3671，3676 16/01/06
            bool b = false;
            DataTable dtt = cprinting_offer.GetTableInfo_show_all();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                DataRow dr1 = dtt.NewRow();
                dr1["部品名"] = dataGridView1["部品名", i].FormattedValue.ToString();
                dr1["项次"] = dataGridView1["项次", i].FormattedValue.ToString();
                dr1["图纸门幅"] = dataGridView1["图纸门幅", i].FormattedValue.ToString();
                dr1["图纸纸长"] = dataGridView1["图纸纸长", i].FormattedValue.ToString();
                if (!string.IsNullOrEmpty(dataGridView1["部品个数", i].FormattedValue.ToString()))
                {
                    dr1["部品个数"] = dataGridView1["部品个数", i].FormattedValue.ToString();
                }
                if (!string.IsNullOrEmpty(dataGridView1["拼模数", i].FormattedValue.ToString()))
                {
                    dr1["拼模数"] = dataGridView1["拼模数", i].FormattedValue.ToString();
                }
                dr1["印刷选项"] = dataGridView1["印刷选项", i].FormattedValue.ToString();
                dr1["面纸"] = dataGridView1["面纸", i].FormattedValue.ToString();
                dr1["芯纸"] = dataGridView1["芯纸", i].FormattedValue.ToString();
                dr1["芯纸规格"] = dataGridView1["芯纸规格", i].FormattedValue.ToString();
                dr1["客户"] = bc.RETURN_CNAME(comboBox1.Text);
                dr1["品牌"] = bc.getOnlyString("SELECT BRAND FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'");
                dtt.Rows.Add(dr1);
            }
            DataTable dtx = cprinting_offer.bind2(dtt, 1, textBox1.Text);
            if (dtx.Rows.Count > 0)
            {

                for (i = 0; i < dtx.Rows.Count; i++)
                {
                    /*StringBuilder sqb = new StringBuilder();
                    sqb.AppendFormat("印刷选项：{0},", dtx.Rows[i]["印刷选项"].ToString());
                    sqb.AppendFormat("机器型号：{0},", dtx.Rows[i]["机器型号"].ToString());
                    sqb.AppendFormat("加工门幅：{0},", dtx.Rows[i]["加工门幅"].ToString());
                    sqb.AppendFormat("加工长度：{0},", dtx.Rows[i]["加工长度"].ToString());
                    sqb.AppendFormat("部品总价：{0},", dtx.Rows[i]["部品总价"].ToString());
               
                    //MessageBox.Show(sqb.ToString ());*/
                    if (dtx.Rows[i]["部品总价"].ToString() == "四开小")
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 印刷最短边小于300", i+1);
                        break;
                    }
                    else if (dtx.Rows[i]["部品总价"].ToString() == "对开小")
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 印刷最短边小于350", i + 1);
                        break;
                    }
                    else if (dtx.Rows[i]["部品总价"].ToString() == "全开小")
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 印刷最短边小于500", i + 1);
                        break;
                    }
                    else if (dtx.Rows[i]["部品总价"].ToString() == "大全开小")
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 印刷最短边小于350", i + 1);
                        break;
                    }
                    else if (dtx.Rows[i]["部品总价"].ToString() == "超出大全开")
                    {
                        b = true;
                        hint.Text = string.Format("项次 {0} 加工门幅大于1200且加工纸长大于1620", i + 1);
                        break;
                    }
                    else if (cprinting_offer.ErrowInfo != null)
                    {
                        b = true;
                        //hint.Text = cprinting_offer.ErrowInfo;//3056 芯纸暂无
                        MessageBox.Show(cprinting_offer.ErrowInfo, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;

                    }
                }
            }
            return b;

        }
        #endregion
        #region dataGridView1_CellEndEdit
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
           
            try
            {

                hint.Text = "";
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                int columnsindex = dataGridView1.CurrentCell.ColumnIndex;

                if (textBox1.Text == "")
                {
                    hint.Text = "数量不能为空";
                }
                else
                {
                    hint.Text = "";
                }
                for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                    dataGridView1.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                    i = i + 1;
                }
                decimal d1 = 0, d2 = 0;
                d1 = 0;
                d2 = 0;
                if (dataGridView1["部品个数", rowsindex].FormattedValue.ToString() != "" &&
                    dataGridView1["拼模数", rowsindex].FormattedValue.ToString() != "")
                {
                    d1 = decimal.Parse(dataGridView1["部品个数", rowsindex].Value.ToString());
                    d2 = decimal.Parse(dataGridView1["拼模数", rowsindex].Value.ToString());
                    if (d2 != 0)
                    {
                        dataGridView1["部品数", rowsindex].Value = d1 / d2;
                    }
                }
                dt = basec.getdts(cprinting_offer.sql + " where A.PFID='" + IDO + "' ORDER BY  B.PFID ASC ");
                if (dataGridView1["面纸", rowsindex].FormattedValue.ToString() != "")
                {
                    //MessageBox.Show("1");
                    dataGridView1["面纸克重", rowsindex].ReadOnly = false;
                    DataGridViewComboBoxCell dgvcc = (DataGridViewComboBoxCell)dataGridView1["面纸克重", rowsindex];

                    dt1 = bc.getdt(string.Format(ctissue_spec.sql + @" WHERE B.TISSUE_SPEC='{0}' AND 
                    SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dataGridView1["面纸", rowsindex].Value.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                    DataTable dtx1 = bc.GET_DT_TO_DV_TO_DT(dt, "", string.Format("面纸='{0}'", dataGridView1["面纸", rowsindex ].Value.ToString()));
                    dt1 = bc.RETURN_NATURE_AND_NOW_DT(dt1, dtx1, "克重", "面纸克重", 0, dtx1.Rows.Count);
                    if (dt1.Rows.Count > 0)
                    {
                        dgvcc.Items.Clear();
                        dgvcc.Items.Add("");
                        foreach (DataRow dr in dt1.Rows)
                        {
                            dgvcc.Items.Add(dr["VALUE"].ToString());
                        }

                    }

                }
                else
                {
                    //MessageBox.Show("2");
                    dataGridView1["面纸克重", rowsindex].ReadOnly = true;
                }

                if (Convert.ToString(dataGridView1["芯纸", rowsindex].Value).Trim() != "")
                {
                    dataGridView1["芯纸规格", rowsindex].ReadOnly = false;
                    DataGridViewComboBoxCell dgvcc = (DataGridViewComboBoxCell)dataGridView1["芯纸规格", rowsindex];

                    dt1 = bc.getdt(string.Format(cpaper_core.sql + @" WHERE B.PAPER_CORE='{0}' AND 
                    SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dataGridView1["芯纸", rowsindex].Value.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                    dt1 = bc.RETURN_NOHAVE_REPEAT_DT(dt1, "规格");
                    if (dt1.Rows.Count > 0)
                    {
                        dgvcc.Items.Clear();
                        dgvcc.Items.Add("");
                        foreach (DataRow dr in dt1.Rows)
                        {

                            dgvcc.Items.Add(dr["VALUE"].ToString());
                        }

                    }
                }
                else
                {
                    dataGridView1["芯纸规格", rowsindex].ReadOnly = true;
                }
                if (dataGridView1["底纸", rowsindex].FormattedValue.ToString() != "")
                {
                    dataGridView1["底纸克重", rowsindex].ReadOnly = false;
                    DataGridViewComboBoxCell dgvcc = (DataGridViewComboBoxCell)dataGridView1["底纸克重", rowsindex];

                    dt1 = bc.getdt(string.Format(ctissue_spec.sql + @" WHERE B.TISSUE_SPEC='{0}' AND 
                    SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dataGridView1["底纸", rowsindex].Value.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                    if (dt1.Rows.Count > 0)
                    {
                        dgvcc.Items.Clear();
                        dgvcc.Items.Add("");
                        foreach (DataRow dr in dt1.Rows)
                        {
                            dgvcc.Items.Add(dr["克重"].ToString());
                        }

                    }
                }
                else
                {
                    dataGridView1["底纸克重", rowsindex].ReadOnly = true;
                }
                if (dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() == "双纸画异" ||
                    dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() == "单纸双异")
                {
                    dataGridView1["双面印刷", rowsindex].ReadOnly = false;
                }
                else
                {
                    dataGridView1["双面印刷", rowsindex].ReadOnly = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
          

        }
        #endregion
        #region dataGridView1_CellClick
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
            try
            {
                int rowsindex = dataGridView1.CurrentCell.RowIndex;
                int columnsindex = dataGridView1.CurrentCell.ColumnIndex;
                if (dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() != "单纸双异" &&
                    dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() != "双纸画异")
                {

                    dataGridView1["反面4C", rowsindex].ReadOnly = true;
                    dataGridView1["反面专色", rowsindex].ReadOnly = true;
                    dataGridView1["反面防晒", rowsindex].ReadOnly = true;
                 
                }
                else
                {
                 
                    dataGridView1["反面4C", rowsindex].ReadOnly = false;
                    dataGridView1["反面专色", rowsindex].ReadOnly = false;
                    dataGridView1["反面防晒", rowsindex].ReadOnly = false;
                }
                if (dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() == "不印刷")
                {
                    dataGridView1["正面4C", rowsindex].ReadOnly = true;
                    dataGridView1["正面专色", rowsindex].ReadOnly = true;
                    dataGridView1["正面防晒", rowsindex].ReadOnly = true;
                }
                else
                {
                    dataGridView1["正面4C", rowsindex].ReadOnly = false;
                    dataGridView1["正面专色", rowsindex].ReadOnly = false;
                    dataGridView1["正面防晒", rowsindex].ReadOnly = false;
                }
                if (dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() == "单纸双同" ||
                 dataGridView1["印刷选项", rowsindex].FormattedValue.ToString() == "单纸双异")
                {

                    dataGridView1["芯纸", rowsindex].ReadOnly = true;
                    dataGridView1["芯纸规格", rowsindex].ReadOnly = true;
                    dataGridView1["底纸", rowsindex].ReadOnly = true;
                    dataGridView1["底纸克重", rowsindex].ReadOnly = true;


                }
                else
                {
                    dataGridView1["芯纸", rowsindex].ReadOnly = false;
                    dataGridView1["芯纸规格", rowsindex].ReadOnly = false;
                    dataGridView1["底纸", rowsindex].ReadOnly = false;
                    dataGridView1["底纸克重", rowsindex].ReadOnly = false;


                }
                if (dataGridView1.Columns[columnsindex].Name == "面纸克重")
                {

                    if (dataGridView1["面纸", rowsindex].FormattedValue.ToString() != "")
                    {
                        dataGridView1["面纸克重", rowsindex].ReadOnly = false;
                    }
                    else
                    {
                        dataGridView1["面纸克重", rowsindex].ReadOnly = true;
                    }
                }
                if (dataGridView1.Columns[columnsindex].Name == "芯纸规格")
                {

                    if (dataGridView1["芯纸", rowsindex].FormattedValue.ToString() != "")
                    {
                        dataGridView1["芯纸规格", rowsindex].ReadOnly = false;
                    }
                    else
                    {
                        dataGridView1["芯纸规格", rowsindex].ReadOnly = true;
                    }
                }
                if (dataGridView1.Columns[columnsindex].Name == "底纸克重")
                {

                    if (dataGridView1["底纸", rowsindex].FormattedValue.ToString() != "")
                    {
                        dataGridView1["底纸克重", rowsindex].ReadOnly = false;
                    }
                    else
                    {
                        dataGridView1["底纸克重", rowsindex].ReadOnly = true;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
    
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {

                dataGridView1[0, i].Value = i + 1;
                if (dataGridView1["拼模数", i].FormattedValue.ToString() == "")
                {
                    dataGridView1["拼模数", i].Value = 1;
                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Height = 18;
                }
            }
            for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].DefaultCellStyle.BackColor = CCOLOR.GLS;
                dataGridView1.Rows[i + 1].DefaultCellStyle.BackColor = CCOLOR.YG;
                i = i + 1;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
    
            audit();
        }
        #region audit
        private void audit()
        {

            try
            {
                SAMPLE_CODE = bc.getOnlyString("SELECT SAMPLE_CODE FROM EMPLOYEEINFO WHERE EMID='" + LOGIN.EMID + "'");
                SAMPLE_CODE_FIRST = SAMPLE_CODE.Substring(0, 1);
                DataTable dtx = bc.getdt(cprinting_offer.sql + " WHERE B.OFFER_ID='" + textBox2.Text + "'");
                string vdate=DateTime .Now .ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
                if (dtx.Rows.Count > 0)
                {
                    if (juage())
                    {

                    }
                    else
                    {

                        if (label1.Text == "未审核")
                        {
                            basec.getcoms(@"UPDATE PRINTING_OFFER_MST SET CHARGE_AUDIT_STATUS='Y',EDIT_TIME='"+vdate +"',OFFER_ID='" + textBox2.Text + "-" + SAMPLE_CODE_FIRST +
                                "'   WHERE PFID='" + IDO + "'");
                            textBox2.Text = bc.getOnlyString("SELECT OFFER_ID FROM PRINTING_OFFER_MST WHERE PFID='" + IDO + "'");
                            pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/audit.png"));
                            label1.Text = "已审核";
                            COST_TOTAL FRM = new COST_TOTAL(F1);
                            FRM.Text = "费用汇总统计表";
                            FRM.OFFER_ID = textBox2.Text;
                            FRM.PFID = bc.getOnlyString("SELECT PFID FROM PRINTING_OFFER_MST WHERE OFFER_ID='" + textBox2.Text + "'");
                            FRM.Show();
                            this.WindowState = FormWindowState.Minimized;
                            F1.bind("");
                        }
                        else
                        {
                            basec.getcoms("UPDATE PRINTING_OFFER_MST SET CHARGE_AUDIT_STATUS='N' ,EDIT_TIME='" + vdate + "',OFFER_ID='" + textBox2.Text.Substring(0, (textBox2.Text).Length - 2)
                                + "' WHERE PFID='" + IDO + "'");
                            textBox2.Text = bc.getOnlyString("SELECT OFFER_ID FROM PRINTING_OFFER_MST WHERE PFID='" + IDO + "'");
                            pictureBox1.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/61.png"));
                            label1.Text = "未审核";
                            basec.getcoms(@"UPDATE PRINTING_OFFER_MST SET AUDIT_OPINION=''   WHERE PFID='" + IDO + "'");
                            F1.bind("");

                        }
                    }
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }


        }
        #endregion
        private void comboBox1_DropDown(object sender, EventArgs e)
        {
           
            PROJECT_INFO FRM = new PROJECT_INFO();
            FRM.WindowState = FormWindowState.Normal;
            FRM.PRINTING_OFFER();
            FRM.ShowDialog();
            this.comboBox1.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox1.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox1.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox1.Text = GET_PROJECT_ID;
              
            }
            /*try
            {
                dtx = bc.getdt(cproject_info.sql + " WHERE DateDiff(day,A.DATE,getdate()) >-1 and DateDiff(day,A.DATE,getdate()) <+20");
                if (dtx.Rows.Count > 0)
                {
                    comboBox1.Items.Clear();
                    foreach (DataRow dr in dtx.Rows)
                    {
                        comboBox1.Items.Add(dr["项目号"].ToString());

                    }
                  
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }*/

        }
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dtx = bc.getdt(cproject_info.sql + " WHERE A.PROJECT_ID='" + comboBox1.Text + "'");
                if (dtx.Rows.Count > 0)
                {
                    textBox3.Text = dtx.Rows[0]["项目名称"].ToString();
                    CUSTOMER_TYPE = bc.RETURN_CUSTOMER_TYPE(comboBox1.Text);
                    comboBox2.Focus();
                    total1();
                    /*写真 start*/
                    for (i = 0; i < 7; i++)
                    {
                        bind_portray_again(i);//根据不同的客户类别加载不同写真的数据源 先加载单击时才有数据源16/01/14
                    }
                    /*写真 end*/
                }
                /* 包装 start*/
                bind_pack_material_again();
                /* 包装 end*/
                /* 运输 start*/
                bind_transport_again();
                /* 运输 end*/
                /* 人工 start*/
                bind_artificial_again();
                /* 人工 end*/
              
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                hint.Text = "";
                int rowsindex = dataGridView2.CurrentCell.RowIndex;
                int columnsindex = dataGridView2.CurrentCell.ColumnIndex;

                if (rowsindex == 0 && columnsindex == 0)
                {
                    dtx = bc.getdt(cdie_cutting_cost.sql + string.Format(@" WHERE A.DIE_CUTTING='{0}' 
", dataGridView2["项目", rowsindex].FormattedValue.ToString()));
                    if (dataGridView2["项目", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                    {

                        //dataGridView2["元米", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();

                    }
                    else
                    {
                        dataGridView2["元米", rowsindex].Value = "";
                    }
                    dtx = bc.getdt(cdie_cutting_cost.sql + string.Format(@" WHERE A.DIE_CUTTING='{0}' 
", "圆孔"));
                    if (dataGridView2["项目", rowsindex].FormattedValue.ToString() != "" &&
                        dataGridView2["圆孔个数", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                    {
                        //dataGridView2["元个", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();
                    }
                    else
                    {
                        dataGridView2["元个", rowsindex].Value = "";
                    }

                }

                if (dataGridView2["元米", 2].FormattedValue.ToString() == "按米计")
                {
                    dataGridView2["项目", 0].ReadOnly = false;
                    dataGridView2["项目", 1].ReadOnly = false;
                }
                else
                {

                    dataGridView2["项目", 0].ReadOnly = true;
                    dataGridView2["项目", 1].ReadOnly = true;
                }
            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView2.CurrentCell.RowIndex;
                int columnsindex = dataGridView2.CurrentCell.ColumnIndex;
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    rowsindex = i;

                    if (dataGridView2["刀模长米", rowsindex].FormattedValue.ToString() != "" && bc.yesno(dataGridView2["刀模长米", rowsindex].FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;

                        MessageBox.Show("刀模长米 只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;

                    }
                    else if (dataGridView2["元米", rowsindex].FormattedValue.ToString() != "" && bc.yesno(dataGridView2["刀模长米", rowsindex].FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;
                        MessageBox.Show("元米 只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                    }
                    if (dataGridView2["圆孔个数", rowsindex].FormattedValue.ToString() != "" && bc.yesno(dataGridView2["刀模长米", rowsindex].FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;
                        MessageBox.Show("圆孔个数 只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                    }
                    else if (dataGridView2["元个", rowsindex].FormattedValue.ToString() != "" && bc.yesno(dataGridView2["刀模长米", rowsindex].FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;
                        MessageBox.Show("元个 只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                    }

                }

            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Focus();
        }
        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
         
            try
            {



                hint.Text = "";
                int rowsindex = dataGridView3.CurrentCell.RowIndex;
                int columnsindex = dataGridView3.CurrentCell.ColumnIndex;
                if ((rowsindex == 0 || rowsindex == 1 || rowsindex == 2 || rowsindex == 3) && dataGridView3.Columns[columnsindex].Name == "写真类型")
                {
                    dtx = bc.getdt(cportray.sql + string.Format(@" WHERE A.PORTRAY_TYPE='{0}' AND 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{1}'", dataGridView3["写真类型", rowsindex].FormattedValue.ToString(), CUSTOMER_TYPE));
                    if (dataGridView3["写真类型", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                    {
                        //dataGridView3["单价", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();
                    }
                    else
                    {
                        dataGridView3["单价", rowsindex].Value = "";
                    }
                    /*sqb = new StringBuilder();
                    sqb.AppendFormat("列名：{0}, ", dataGridView3.Columns[columnsindex].Name.ToString());
                    sqb.AppendFormat("列值：{0}, ", dataGridView3["写真类型", rowsindex].FormattedValue.ToString());
                    sqb.AppendFormat("当前行索引：{0}, ", rowsindex);
                    MessageBox.Show(sqb.ToString());*/
                }
            

            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void dataGridView3_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView3.CurrentCell.RowIndex;
                int columnsindex = dataGridView3.CurrentCell.ColumnIndex;
                if (dataGridView3.Columns[columnsindex].Name == "长" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView3.Columns[columnsindex].Name == "宽" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (dataGridView3.Columns[columnsindex].Name == "总数量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView3.Columns[columnsindex].Name == "单价" && bc.yesno(e.FormattedValue.ToString()) == 0)
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

 

        private void dataGridView4_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView4.CurrentCell.RowIndex;
                int columnsindex = dataGridView4.CurrentCell.ColumnIndex;

                if (dataGridView4.Columns[columnsindex].Name == "用量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView4.Columns[columnsindex].Name == "单价" && bc.yesno(e.FormattedValue.ToString()) == 0)
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

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                hint.Text = "";
                int rowsindex = dataGridView4.CurrentCell.RowIndex;
                int columnsindex = dataGridView4.CurrentCell.ColumnIndex;
                if (rowsindex == 0 || rowsindex == 1 || rowsindex == 2 || rowsindex == 3 || rowsindex == 4 || rowsindex == 5 || rowsindex == 6 || rowsindex == 7)
                {
               
                    dtx = bc.getdt(cparts_auxiliary.sql + string.Format(" WHERE A.PARTS_AUXILIARY='{0}'",
                        dataGridView4["配件名", rowsindex].FormattedValue.ToString()));
                    if (dataGridView4["配件名", rowsindex].FormattedValue.ToString() != "")
                    {
                        if (dtx.Rows.Count > 0 && dataGridView4.Columns[columnsindex].Name == "配件名")
                        {
                            //dataGridView4["单价", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();
                            dataGridView4["单位", rowsindex].Value = dtx.Rows[0]["单位"].ToString();
                        }
                    }
                    else
                    {
                        dataGridView4["单价", rowsindex].Value = "";
                        dataGridView4["单位", rowsindex].Value = "";
                    }
                    /*sqb = new StringBuilder();
                    sqb.AppendFormat("列名：{0}, ", dataGridView4.Columns[columnsindex].Name.ToString());
                    sqb.AppendFormat("列值：{0}, ", dataGridView4["配件名", rowsindex].FormattedValue.ToString());
                    sqb.AppendFormat("当前行索引：{0}, ", rowsindex);
                    MessageBox.Show(sqb.ToString());*/
                }
            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void dataGridView5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                hint.Text = "";
                int rowsindex = dataGridView5.CurrentCell.RowIndex;
                int columnsindex = dataGridView5.CurrentCell.ColumnIndex;
                if ((rowsindex == 4 || rowsindex == 5 || rowsindex == 6) && dataGridView5.Columns[columnsindex].Name == "箱形")
                {
     
                    dtx = bc.getdt(cpack_material.sql + string.Format(" WHERE A.PACK_MATERIAL='{0}' AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{1}'",
                        dataGridView5["箱形", rowsindex].FormattedValue.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                    if (dataGridView5["箱形", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                    {
                        //dataGridView5["单价", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();
                        dataGridView5["材质", rowsindex].Value = dtx.Rows[0]["单位"].ToString();
                    }
                    else
                    {
                        dataGridView5["单价", rowsindex].Value = "";
                        dataGridView5["材质", rowsindex].Value = "";
                    }
                }
                rowsindex = dataGridView5.CurrentCell.RowIndex;
                columnsindex = dataGridView5.CurrentCell.ColumnIndex;
                if ((rowsindex == 0 || rowsindex == 1 || rowsindex == 2 || rowsindex == 3) && dataGridView5.Columns[columnsindex].Name == "材质")
                {
                    dtx = bc.getdt(cpack_material.sql + string.Format(" WHERE A.PACK_MATERIAL='{0}' AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{1}'",
                        dataGridView5["材质", rowsindex].FormattedValue.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));

                    if (dataGridView5["数量", rowsindex].FormattedValue.ToString() != "" && dataGridView5["长", rowsindex].FormattedValue.ToString() != ""
                        && dataGridView5["宽", rowsindex].FormattedValue.ToString() != "" && dataGridView5["高", rowsindex].FormattedValue.ToString() != ""
                        && dataGridView5["箱形", rowsindex].FormattedValue.ToString() != "" && dataGridView5["材质", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                    {
                        //dataGridView5["单价", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();
                    }
                    else
                    {
                        dataGridView5["单价", rowsindex].Value = "";
                    }
                    if (rowsindex == 3)
                    {

                        if (dataGridView5["数量", rowsindex].FormattedValue.ToString() != "" && dataGridView5["长", rowsindex].FormattedValue.ToString() != ""
                            && dataGridView5["宽", rowsindex].FormattedValue.ToString() != "" && dataGridView5["材质", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                        {
                           // dataGridView5["单价", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();
                        }
                        else
                        {
                            dataGridView5["单价", rowsindex].Value = "";
                        }
                    }
                }
                dgv8(rowsindex, false);
            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void dataGridView5_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView5.CurrentCell.RowIndex;
                int columnsindex = dataGridView5.CurrentCell.ColumnIndex;

                if (dataGridView5.Columns[columnsindex].Name == "数量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (rowsindex == 0 || rowsindex == 1 || rowsindex == 2 || rowsindex == 3)
                {

                    if (dataGridView5.Columns[columnsindex].Name == "长" && bc.yesno(e.FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;
                        MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }
                    else if (dataGridView5.Columns[columnsindex].Name == "宽" && bc.yesno(e.FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;
                        MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }
                    else if (dataGridView5.Columns[columnsindex].Name == "高" && bc.yesno(e.FormattedValue.ToString()) == 0)
                    {
                        e.Cancel = true;
                        MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }

                }
                else if (dataGridView5.Columns[columnsindex].Name == "单价" && bc.yesno(e.FormattedValue.ToString()) == 0)
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

        private void dataGridView6_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView6.CurrentCell.RowIndex;
                int columnsindex = dataGridView6.CurrentCell.ColumnIndex;

                if (dataGridView6.Columns[columnsindex].Name == "数量" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView6.Columns[columnsindex].Name == "单价" && bc.yesno(e.FormattedValue.ToString()) == 0)
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

        private void dataGridView6_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                hint.Text = "";
                int rowsindex = dataGridView6.CurrentCell.RowIndex;
                int columnsindex = dataGridView6.CurrentCell.ColumnIndex;
                if (rowsindex == 0 && dataGridView6.Columns[columnsindex].Name == "项目")
                {
                    dtx = bc.getdt(cartificial.sql + string.Format(" WHERE A.ARTIFICIAL='{0}' AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{1}'",
                        dataGridView6["项目", rowsindex].FormattedValue.ToString(), bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                    if (dataGridView6["项目", rowsindex].FormattedValue.ToString() != "" && dtx.Rows.Count > 0)
                    {

                        //dataGridView6["单价", rowsindex].Value = dtx.Rows[0]["未税单价"].ToString();

                    }
                    else
                    {
                        dataGridView6["单价", rowsindex].Value = "";

                    }
                }

            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        private void dataGridView7_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView7.CurrentCell.RowIndex;
                int columnsindex = dataGridView7.CurrentCell.ColumnIndex;

                if (dataGridView7.Columns[columnsindex].Name == "外购价一" && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView7.Columns[columnsindex].Name == "外购价二" && bc.yesno(e.FormattedValue.ToString()) == 0)
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

        private void dataGridView7_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView8_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                int rowsindex = dataGridView8.CurrentCell.RowIndex;
                int columnsindex = dataGridView8.CurrentCell.ColumnIndex;

                if (dataGridView8.Columns[columnsindex].Name == "总箱数" && bc.yesno(e.FormattedValue.ToString()) == 0)
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
        private void dgv8(int rows, bool b)
        {

            try
            {
                PACK_LENGTH = 0;
                PACK_WIDTH = 0;
                PACK_HEIGHT = 0;
                TOTAL_BOXS_COUNT = 0;
                decimal d1 = 0, d3 = 0, d4 = 0;
                hint.Text = "";
                int rowsindex = dataGridView8.CurrentCell.RowIndex;
                int columnsindex = dataGridView8.CurrentCell.ColumnIndex;

                if (b)
                {

                }
                else
                {
                    rowsindex = rows;
                }

                if ((rowsindex == 0 || rowsindex == 1 || rowsindex == 2) && dataGridView5.Rows.Count > 0)
                {

                    if (dataGridView5["长", rowsindex].FormattedValue.ToString() != "")
                    {
                        dataGridView8["长", rowsindex].Value = dataGridView5["长", rowsindex].FormattedValue.ToString();
                        PACK_LENGTH = decimal.Parse(dataGridView5["长", rowsindex].FormattedValue.ToString());
                    }
                    else
                    {
                        dataGridView8["长", rowsindex].Value = "";
                    }
                    if (dataGridView5["宽", rowsindex].FormattedValue.ToString() != "")
                    {
                        dataGridView8["宽", rowsindex].Value = dataGridView5["宽", rowsindex].FormattedValue.ToString();
                        PACK_WIDTH = decimal.Parse(dataGridView5["宽", rowsindex].FormattedValue.ToString());
                    }
                    else
                    {
                        dataGridView8["宽", rowsindex].Value = "";
                    }
                    if (dataGridView5["高", rowsindex].FormattedValue.ToString() != "")
                    {
                        dataGridView8["高", rowsindex].Value = dataGridView5["高", rowsindex].FormattedValue.ToString();
                        PACK_HEIGHT = decimal.Parse(dataGridView5["高", rowsindex].FormattedValue.ToString());
                    }
                    else
                    {
                        dataGridView8["高", rowsindex].Value = "";
                    }
                    if (rowsindex == 0 || rowsindex == 1)
                    {
                        if (dataGridView5["数量", rowsindex].FormattedValue.ToString() != "" && !string.IsNullOrEmpty(textBox1.Text))
                        {
                            dataGridView8["总箱数", rowsindex].Value = decimal.Parse(dataGridView5["数量", rowsindex].FormattedValue.ToString()) * decimal.Parse(textBox1.Text);
                        }
                        else
                        {
                            dataGridView8["总箱数", rowsindex].Value = "";
                        }
                        TOTAL_BOXS_COUNT = decimal.Parse(dataGridView8["总箱数", rowsindex].Value.ToString());
                    }

                }

                if ((rowsindex == 0 || rowsindex == 1) && dataGridView8.Rows.Count > 0)
                {
                    if (PACK_LENGTH != 0 && PACK_WIDTH != 0 && PACK_HEIGHT != 0 && TOTAL_BOXS_COUNT != 0)
                    {

                        d1 = (PACK_LENGTH + 10) / 1000 * (PACK_WIDTH + 10) / 1000 *
                            (PACK_HEIGHT + 10) / 1000 * TOTAL_BOXS_COUNT;
                        dataGridView8["总立方数", rowsindex].Value = d1.ToString("0.00");
                    }
                    if (dataGridView8["运输方式", rowsindex].FormattedValue.ToString() != "" &&
                        !string.IsNullOrEmpty(dataGridView8["总立方数", rowsindex].FormattedValue.ToString()) &&
                        decimal.Parse(dataGridView8["总立方数", rowsindex].FormattedValue.ToString()) != 0 &&
           (rowsindex == 0 || rowsindex == 1))
                    {

                        DataTable dtx1 = bc.getdt(@"
SELECT * FROM TRANSPORT A WHERE A.TRANSPORT='" + dataGridView8["运输方式", rowsindex].FormattedValue.ToString() + "'AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='" +
                                                   bc.RETURN_CUSTOMER_TYPE(comboBox1.Text) + "' ");
                        if (dtx1.Rows.Count > 0)
                        {

                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                            {
                                d4 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                            }
                            if (d1 < 50)
                            {
                                if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE_ONE"].ToString()))
                                {
                                    d3 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE_ONE"].ToString());
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE_TWO"].ToString()))
                                {
                                    d3 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE_TWO"].ToString());
                                }
                            }
                            //dataGridView8["单价", rowsindex].Value = (d3 / (1 + d4 / 100)).ToString("0.00");

                        }

                    }

                }

            }
            catch (Exception)
            {

                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }


        }
        private void dataGridView8_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                dgv8(0, true);
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                for (i = 0; i < dataGridView8.Rows.Count; i++)
                {
                    dgv8(i, false);
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" +IDO  + "'  ORDER BY E.PFKEY ASC");
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_FOR_BASEINFO_PURCHASE(dtt, "xxx报价纸品基本信息FOR_核价采购",
                        System.IO.Path.GetFullPath("xxx报价纸品基本信息FOR_核价采购.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" + IDO + "' ORDER BY E.PFKEY ASC");
                dtt = cprinting_offer.RETURN_DT_SHOW_HIDE_FORM(dtt);
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_FOR_NUCLEAR_PRICE(dtt, "xxx报价纸品估计计算表FOR_核价",
                        System.IO.Path.GetFullPath("xxx报价纸品估计计算表FOR_核价.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" + IDO + "' ORDER BY E.PFKEY ASC");
                dtt = cprinting_offer.RETURN_DT_SHOW_HIDE_FORM(dtt);
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_FOR_NUCLEAR_PURCHASE(dtt, "xxx报价纸品预算明细表FOR_采购",
                        System.IO.Path.GetFullPath("xxx报价纸品预算明细表FOR_采购.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" + IDO + "' ORDER BY E.PFKEY ASC");
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_FOR_BASEINFO_AE(dtt, "xxx报价纸品基本信息表FOR_AE",
                        System.IO.Path.GetFullPath("xxx报价纸品基本信息表FOR_AE.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" + IDO + "'");
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_FOR_MAIN_DETAIL(dtt, "xxx报价纸品主件明细表FOR_AE",
                        System.IO.Path.GetFullPath("xxx报价纸品主件明细表FOR_AE.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" + IDO + "'");
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_OFFER_FOR_AE_1(dtt, "产品报价单AE",
                        System.IO.Path.GetFullPath("产品报价单AE.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                dtt = bc.getdt(cprinting_offer.sqlni + " WHERE A.PFID='" + IDO + "'");
                if (dtt.Rows.Count > 0)
                {
                    cprinting_offer.ExcelPrint_OFFER_FOR_AE_2(dtt, "含明细报价单AE",
                        System.IO.Path.GetFullPath("含明细报价单AE.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    basec.getcoms("DELETE PRINTING_OFFER_MST WHERE PFID='" + IDO + "'");
                    basec.getcoms("DELETE PRINTING_OFFER_DET WHERE PFID='" + IDO + "'");
                    add();
                    hint.Text = "";
                    F1.load();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //根据不同的客户类别加载不同的数据源 16/01/14
            try
            {
                hint.Text = "";
                int rowsindex = dataGridView3.CurrentCell.RowIndex;
                int columnsindex = dataGridView3.CurrentCell.ColumnIndex;
                bind_portray_again(rowsindex);
                
            }
            catch (Exception)
            {

            }
                   

        }
        private void bind_portray_again(int row)
        {
            //根据不同的客户类别加载不同的数据源 16/01/14
            try
            {
                DataGridViewComboBoxCell dgvcc = (DataGridViewComboBoxCell)dataGridView3["写真类型", row];
                dtx = bc.getdt(string.Format(cportray.sql + @" WHERE 
SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(comboBox1.Text)));
                //MessageBox.Show(bc.RETURN_CUSTOMER_TYPE(comboBox1.Text));
                dtx = bc.GET_DT_TO_DV_TO_DT(dtx, "", "写真类型 NOT IN ('批次写真运费')");
                dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtx, "写真类型");
                if (dtx.Rows.Count > 0)
                {
                    dgvcc.Items.Clear();
                    dgvcc.Items.Add("");
                    foreach (DataRow dr in dtx.Rows)
                    {
                        dgvcc.Items.Add(dr["VALUE"].ToString());
                    }

                }
            }
            catch (Exception)
            {

            }

        }
        private void dataGridView3_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
         //写真绑定的数据源项在重新绑定的数据源中不存在，用此事件不做系统出错提示 16/01/14
        }

        private void dataGridView4_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //写真绑定的数据源项在重新绑定的数据源中不存在，用此事件不做系统出错提示 16/01/14
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView7_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView8_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView7_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //写真绑定的数据源项在重新绑定的数据源中不存在，用此事件不做系统出错提示 16/01/14
        }

        private void dataGridView5_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //写真绑定的数据源项在重新绑定的数据源中不存在，用此事件不做系统出错提示 16/01/14
        }

        private void dataGridView8_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //写真绑定的数据源项在重新绑定的数据源中不存在，用此事件不做系统出错提示 16/01/14
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        
        }


        private void progressBar1_SizeChanged(object sender, EventArgs e)
        {
         
        }

        private void button8_Click(object sender, EventArgs e)
        {
        
   
        }

 

 

    }
}
