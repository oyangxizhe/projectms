using System;
using System.Collections.Generic;

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
    public partial class PN_PRODUCTION_INSTRUCTIONST : Form
    {
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable notice_employee = new DataTable();

        basec bc=new basec ();
        CMATERIAL_PRICE cmaterial_price = new CMATERIAL_PRICE();
        CAUDIT_LIST caudit_list = new CAUDIT_LIST();
        #region nature
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _INITIAL_OR_OTHER;
        public string INITIAL_OR_OTHER
        {
            set { _INITIAL_OR_OTHER = value; }
            get { return _INITIAL_OR_OTHER; }
        }
        private int _EDIT_TIMES;
        public int EDIT_TIMES
        {
            set { _EDIT_TIMES = value; }
            get { return _EDIT_TIMES; }
        }
        private static string _GET_PROJECT_ID;
        public static string GET_PROJECT_ID
        {
            set { _GET_PROJECT_ID = value; }
            get { return _GET_PROJECT_ID; }
        }
        private string _WATER_MARK_CONTENT;
        public string WATER_MARK_CONTENT
        {
            set { _WATER_MARK_CONTENT = value; }
            get { return _WATER_MARK_CONTENT; }

        }
        private string _OLD_FILE_NAME;
        public string OLD_FILE_NAME
        {
            set { _OLD_FILE_NAME = value; }
            get { return _OLD_FILE_NAME; }

        }
        private string _NEW_FILE_NAME;
        public string NEW_FILE_NAME
        {
            set { _NEW_FILE_NAME = value; }
            get { return _NEW_FILE_NAME; }

        }
        private string _AE_MAKERID_PAPER_PRODUCTION;
        public string AE_MAKERID_PAPER_PRODUCTION
        {
            set { _AE_MAKERID_PAPER_PRODUCTION = value; }
            get { return _AE_MAKERID_PAPER_PRODUCTION; }

        }
        private string _PROJECT_ID;
        public string PROJECT_ID
        {
            set { _PROJECT_ID = value; }
            get { return _PROJECT_ID; }

        }
        private static string _EMPLOYEE_ID;
        public static string EMPLOYEE_ID
        {
            set { _EMPLOYEE_ID = value; }
            get { return _EMPLOYEE_ID; }

        }
        private string _EDIT;
        public string EDIT
        {
            set { _EDIT = value; }
            get { return _EDIT; }
        }
        private static string _ENAME;
        public static string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private static string _CUID;
        public static string CUID
        {
            set { _CUID = value; }
            get { return _CUID; }
        }
        private static string _CNAME;
        public static string CNAME
        {
            set { _CNAME = value; }
            get { return _CNAME; }
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
  
        private static string _CO_WAREID;
        public static string CO_WAREID
        {
            set { _CO_WAREID = value; }
            get { return _CO_WAREID; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        private string _INITIAL_MAKERID;
        public string INITIAL_MAKERID
        {
            set { _INITIAL_MAKERID = value; }
            get { return _INITIAL_MAKERID; }

        }
        private string _PAPER_PRODUCTION_AUDIT_MAKERID;
        public string PAPER_PRODUCTION_AUDIT_MAKERID
        {
            set { _PAPER_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _PAPER_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _WOOD_IRON_PRODUCTION_AUDIT_MAKERID;
        public string WOOD_IRON_PRODUCTION_AUDIT_MAKERID
        {
            set { _WOOD_IRON_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _WOOD_IRON_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _ACRYLIC_PRODUCTION_AUDIT_MAKERID;
        public string ACRYLIC_PRODUCTION_AUDIT_MAKERID
        {
            set { _ACRYLIC_PRODUCTION_AUDIT_MAKERID = value; }
            get { return _ACRYLIC_PRODUCTION_AUDIT_MAKERID; }

        }
        private string _PAPER_PLAN_AUDIT_MAKERID;
        public string PAPER_PLAN_AUDIT_MAKERID
        {
            set { _PAPER_PLAN_AUDIT_MAKERID = value; }
            get { return _PAPER_PLAN_AUDIT_MAKERID; }

        }
        private string _WOOD_IRON_PLAN_AUDIT_MAKERID;
        public string WOOD_IRON_PLAN_AUDIT_MAKERID
        {
            set { _WOOD_IRON_PLAN_AUDIT_MAKERID = value; }
            get { return _WOOD_IRON_PLAN_AUDIT_MAKERID; }
        }
        private string _STRUCTURE_AUDIT_MAKERID;
        public string STRUCTURE_AUDIT_MAKERID
        {
            set { _STRUCTURE_AUDIT_MAKERID = value; }
            get { return _STRUCTURE_AUDIT_MAKERID; }

        }
        private string _PLANE_AUDIT_MAKERID;
        public string PLANE_AUDIT_MAKERID
        {
            set { _PLANE_AUDIT_MAKERID = value; }
            get { return _PLANE_AUDIT_MAKERID; }

        }
        private string _PAPER_PURCHASE_AUDIT_MAKERID;
        public string PAPER_PURCHASE_AUDIT_MAKERID
        {
            set { _PAPER_PURCHASE_AUDIT_MAKERID = value; }
            get { return _PAPER_PURCHASE_AUDIT_MAKERID; }

        }
        private string _WOOD_IRON_PURCHASE_AUDIT_MAKERID;
        public string WOOD_IRON_PURCHASE_AUDIT_MAKERID
        {
            set { _WOOD_IRON_PURCHASE_AUDIT_MAKERID = value; }
            get { return _WOOD_IRON_PURCHASE_AUDIT_MAKERID; }
        }
        #endregion
        CFileInfo cfileinfo = new CFileInfo();
        CEMPLOYEE_INFO cemployee_info = new CEMPLOYEE_INFO();
        private  delegate bool dele(string a1,string a2);
        private delegate void delex();
        PN_PRODUCTION_INSTRUCTIONS F1 = new PN_PRODUCTION_INSTRUCTIONS();
        protected int M_int_judge, i;
        protected int select;
        CPN_PRODUCTION_INSTRUCTIONS cPN_PRODUCTION_INSTRUCTIONS = new CPN_PRODUCTION_INSTRUCTIONS();
        CPROCESSING_TECHNOLOGY cprocessing_technology = new CPROCESSING_TECHNOLOGY();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
        CNO_PAPER_OFFER cno_paper_offer = new CNO_PAPER_OFFER();
        CSAMPLE_RELY_LIST csample_rely_list = new CSAMPLE_RELY_LIST();
        CPROJECT_INFO cproject_info = new CPROJECT_INFO();
        public static List<CEMPLOYEE_INFO> list1 = new List<CEMPLOYEE_INFO>();
        StringBuilder sqb = new StringBuilder();
        CNOTICE_LIST cnotice_list = new CNOTICE_LIST();
        string sql = @"
INSERT INTO REMIND
(
RIID,
NOTICE_MAKERID,
RECEIVE_STATUS,
NOTICE_OR_AUDIT,
DATE
) 
VALUES
(
@RIID,
@NOTICE_MAKERID,
@RECEIVE_STATUS,
@NOTICE_OR_AUDIT,
@DATE
)
";
          public PN_PRODUCTION_INSTRUCTIONST(PN_PRODUCTION_INSTRUCTIONS  FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        public PN_PRODUCTION_INSTRUCTIONST()
        {
            InitializeComponent();
        }
      
        private void PN_PRODUCTION_INSTRUCTIONST_Load(object sender, EventArgs e)
        {
            
            right();
           dtx=bc.getdt(@"
SELECT * FROM [OFFER].[dbo].[EmployeeInfo] 
WHERE (Position LIKE '%AE%' OR Position  LIKE '%核价%' OR Position LIKE '%总经理%' OR Position  LIKE '%财务%' OR ENAME='系统管理')  
AND  EMID='" + LOGIN.EMID + "'");
            if (dtx.Rows.Count>0)
            {
                label30.Visible = true;
                label47.Visible = true;
                label48.Visible = true;
                textBox5.Visible = true;
                textBox28.Visible = true;
                textBox29.Visible = true;
            }
            else
            {
                label30.Visible = false;
                label47.Visible = false;
                label48.Visible = false;
                textBox5.Visible = false;
                textBox28.Visible = false;
                textBox29.Visible = false;
            }
       

            DataGridViewCheckBoxColumn dgvc1 = new DataGridViewCheckBoxColumn();
            dgvc1.Name = "复选框";
            dataGridView1.Columns.Add(dgvc1);
            DataGridViewTextBoxColumn dgvc2 = new DataGridViewTextBoxColumn();
            dgvc2.Name = "文件名";
            dataGridView1.Columns.Add(dgvc2);
            DataGridViewImageColumn dgvc3 = new DataGridViewImageColumn();

            dgvc3.Name = "缩略图";
            dataGridView1.Columns.Add(dgvc3);
            DataGridViewTextBoxColumn dgvc4 = new DataGridViewTextBoxColumn();
            dgvc4.Name = "索引";
            dgvc4.Visible = false;
            dataGridView1.Columns.Add(dgvc4);
            DataGridViewTextBoxColumn dgvc5 = new DataGridViewTextBoxColumn();
            dgvc5.Name = "新文件名";
            dgvc5.Visible = false;
            dataGridView1.Columns.Add(dgvc5);
            label59.Text = "";
            label61.Text = "";
            label63.Text = "";
            label65.Text = "";

          
            try
            {
                label52.Text = "";
                label53.Visible = false;
                label55.Visible = false;
                label56.Visible = false;
                label57.Visible = false;

                progressBar1.Visible = false;
                ClearText();
                label29.Text = "";
              
                //label40.Text = "";
                //label40.Font = new Font("微软黑体", 12, FontStyle.Bold);
                //label40.ForeColor = CCOLOR.yanghong;
                label5.Text = "下单日期 " + DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
                label6.Text = "修改日期 " + DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
                comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
                comboBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
                comboBox2.BackColor = CCOLOR.CUSTOMER_YELLOW;
                comboBox3.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox4.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox5.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox6.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox7.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox25.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox26.BackColor = CCOLOR.CUSTOMER_YELLOW;
                dt1 = bc.getdt("SELECT * FROM ORDER_TYPE");
                if (dt1.Rows.Count > 0)
                {
                    comboBox2.Items.Clear();
                    comboBox2.Items.Add("");
                    foreach (DataRow dr in dt1.Rows)
                    {

                        comboBox2.Items.Add(dr["ORDER_TYPE"].ToString());
                    }
                }
                comboBox1.Focus();

                //IDO = "SR15120263";//样板计费调用方法 RETURN_SAMPLE_PRICE()
                //right();
                #region load_bind
                if ((Screen.AllScreens[0].Bounds.Width == 1366 && Screen.AllScreens[0].Bounds.Height == 768) ||
                     (Screen.AllScreens[0].Bounds.Width == 1280 && Screen.AllScreens[0].Bounds.Height == 800))
                {
                    this.AutoScroll = true;
                    this.AutoScrollMinSize = new Size(900, 900);

                }
                else
                {
                    this.AutoScroll = true;
                    this.AutoScrollMinSize = new Size(1900, 1000);
                }
                textBox22.ScrollBars = ScrollBars.Both;
                textBox27.ScrollBars = ScrollBars.Both;
                textBox27.BorderStyle = BorderStyle.FixedSingle;

                //this.Icon = this.Icon = Resource1.xz_200X200;
                dateTimePicker1.CustomFormat = "yyyy/MM/dd";
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                comboBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
                textBox2.BackColor = CCOLOR.CUSTOMER_YELLOW;
                comboBox13.BackColor = CCOLOR.CUSTOMER_YELLOW;
                if (PROJECT_ID != null)
                {
                    comboBox1.Text = PROJECT_ID;
                    textBox2.Text = PROJECT_NAME;
                }
                #endregion
                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("生产指示书", LOGIN.USID);
            btnAdd.Visible = false;
            btnSave.Visible = false;
            btnDel.Visible = false;
            label15.Visible = false;
            label17.Visible = false;
            label36.Visible = false;
   
            if (dtx.Rows.Count > 0)
            {

                if (dtx.Rows[0]["新增权限"].ToString() == "有权限")
                {
                    btnAdd.Visible = true;
                    btnSave.Visible = true;
                    label15.Visible = true;
                    label17.Visible = true;
                }
                if (dtx.Rows[0]["删除权限"].ToString() == "有权限")
                {
                    btnDel.Visible = true;
                    label36.Visible = true;
                }
                if (dtx.Rows[0]["修改权限"].ToString() == "有权限")
                {
                    btnSave.Visible = true;
                    label15.Visible = true;
                    EDIT = "有权限";
                }
            
            }
        }
        #endregion
        #region clearText
        public void ClearText()
        {
            label5.Text ="下单日期 "+ DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
            label6.Text = "修改日期 " + DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox3.Text = "";
            comboBox1.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            DateTime date1 = Convert.ToDateTime( DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/"));
            textBox6.Text = "";
            comboBox2.Text ="";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";
            textBox24.Text = "";
            textBox25.Text = "";
            textBox26.Text = "";
            textBox27.Text = "";
            label25.Text = "";
            label26.Text = "";
            label27.Text = "";
            label28.Text = "";
            label29.Text = "";
            button1.Text = "待接收确认";
            button2.Text = "待接收确认";
            button3.Text = "待接收确认";
            button4.Text = "待接收确认";
            button5.Text = "待接收确认";
            button6.Text = "待接收确认";
            button7.Text = "待接收确认";
            button8.Text = "待接收确认";
            button9.Text = "待接收确认";
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            comboBox4.Text = "";
            comboBox5.Text = "";
            comboBox13.Text = "";
            label59.Text = "";
            label65.Text = "";
        }
        #endregion
     
        #region bind
        private void bind()
        {
          
            //listBox1.Items.Clear();
            EDIT_TIMES = 0;
            label7.Text = "第" + EDIT_TIMES + "次修改";
            /*DataTable dtt = bc.getdt(cnotice_list.sql);
            if (dtt.Rows.Count > 0)
            {
                foreach (DataRow dr in dtt.Rows )
                {
                    listBox1.Items.Add(dr["员工工号"].ToString() + " " + dr["员工姓名"].ToString());
                }
            }*/
            label25.Text = "";
            label26.Text = "";
            label27.Text = "";
            label28.Text = "";
            //hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
            hint_bind();
            dt= basec.getdts(cPN_PRODUCTION_INSTRUCTIONS.sql +" WHERE A.PNID='"+IDO+"'");

            pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/send_audit_1.png"));
           
                
            if (dt.Rows.Count > 0)
            {
                if (cPN_PRODUCTION_INSTRUCTIONS.JUAGE_IF_AUDIT_END(dt.Rows[0]["编号"].ToString()))
                {
                    pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/end_audit.png"));
                }
                else  if (dt.Rows[0]["签核状态"].ToString() == "已送签")
                {
                    pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/wait_audit.png"));
                    
                }
           
                EDIT_TIMES = Convert.ToInt32(dt.Rows[0]["修改次数"].ToString());
                label7.Text = "第" + EDIT_TIMES + "次修改";
                textBox1.Text = dt.Rows[0]["订单编号"].ToString();
           
                 
                comboBox3.Text = dt.Rows[0]["项目号"].ToString();
                
                textBox3.Text = dt.Rows[0]["项目名称"].ToString();
                textBox2.Text = dt.Rows[0]["品号"].ToString();
                comboBox1.Text = dt.Rows[0]["报价编号"].ToString();
                textBox4.Text = dt.Rows[0]["生产数量"].ToString();
                textBox5.Text = dt.Rows[0]["含税单价"].ToString();
                dateTimePicker1.Text = dt.Rows[0]["交货日期"].ToString();
                textBox6.Text = dt.Rows[0]["交货批次"].ToString();
                textBox7.Text = dt.Rows[0]["交货地点"].ToString();
                comboBox2.Text = dt.Rows[0]["订单类型"].ToString();
                if (string.IsNullOrEmpty(dt.Rows[0]["项目号"].ToString()))
                {
                    textBox8.Text = dt.Rows[0]["导入的客户"].ToString();
                    textBox9.Text = dt.Rows[0]["导入的品牌"].ToString();
                    textBox10.Text = dt.Rows[0]["导入的AE"].ToString();
                }
                else
                {
                    textBox8.Text = dt.Rows[0]["客户名称"].ToString();
                    textBox9.Text = dt.Rows[0]["品牌"].ToString();
                    textBox10.Text = dt.Rows[0]["AE01"].ToString();
                }
               
                textBox13.Text = dt.Rows[0]["AE02"].ToString();
                textBox16.Text = dt.Rows[0]["AE03"].ToString();
                textBox11.Text = dt.Rows[0]["平面01"].ToString();
                textBox14.Text = dt.Rows[0]["平面02"].ToString();
                textBox17.Text = dt.Rows[0]["平面03"].ToString();
                textBox12.Text = dt.Rows[0]["结构01"].ToString();
                textBox15.Text = dt.Rows[0]["结构02"].ToString();
                textBox18.Text = dt.Rows[0]["结构03"].ToString();
                textBox19.Text = dt.Rows[0]["包装方式"].ToString();
                textBox20.Text = dt.Rows[0]["外箱材质"].ToString();
                textBox21.Text = dt.Rows[0]["长"].ToString();
                textBox22.Text = dt.Rows[0]["宽"].ToString();
                textBox23.Text = dt.Rows[0]["高"].ToString();
                textBox24.Text = dt.Rows[0]["外箱重量"].ToString();
                textBox25.Text = dt.Rows[0]["说明书尺寸"].ToString();
                textBox26.Text = dt.Rows[0]["说明书要求"].ToString();
                textBox27.Text = dt.Rows[0]["生产注意事项"].ToString();
                comboBox13.Text = dt.Rows[0]["报价"].ToString();
                if (dt.Rows[0]["是否需纸品生产签核"].ToString() == "是")
                {
                    checkBox1.Checked = true;
                }
                else
                {
                    checkBox1.Checked = false;
                }
                if (dt.Rows[0]["是否需木铁生产签核"].ToString() == "是")
                {
                    checkBox2.Checked = true;
                }
                else
                {
                    checkBox2.Checked = false;
                }
                if (dt.Rows[0]["是否需亚克力生产签核"].ToString() == "是")
                {
                    checkBox3.Checked = true;
                }
                else
                {
                    checkBox3.Checked = false;
                }

                if (dt.Rows[0]["是否需纸品计划签核"].ToString() == "是")
                {
                    checkBox4.Checked = true;
                }
                else
                {
                    checkBox4.Checked = false;
                }
                if (dt.Rows[0]["是否需木铁计划签核"].ToString() == "是")
                {
                    checkBox5.Checked = true;
                }
                else
                {
                    checkBox5.Checked = false;
                }
                if (dt.Rows[0]["是否需结构设计签核"].ToString() == "是")
                {
                    checkBox6.Checked = true;
                    comboBox4.Text = dt.Rows[0]["指定结构设计"].ToString() +"-"+ dt.Rows[0]["指定结构设计工号"].ToString();
                }
                else
                {
                    checkBox6.Checked = false;
                }
                if (dt.Rows[0]["是否需平面设计签核"].ToString() == "是")
                {
                    checkBox7.Checked = true;
                    comboBox5.Text = dt.Rows[0]["指定平面设计"].ToString() + "-" + dt.Rows[0]["指定平面设计工号"].ToString();
                }
                else
                {
                    checkBox7.Checked = false;
                }

                if (dt.Rows[0]["是否需纸品采购签核"].ToString() == "是")
                {
                    checkBox8.Checked = true;
                }
                else
                {
                    checkBox8.Checked = false;
                }
                if (dt.Rows[0]["是否需木铁采购签核"].ToString() == "是")
                {
                    checkBox9.Checked = true;
                }
                else
                {
                    checkBox9.Checked = false;
                }
                if (dt.Rows[0]["纸品生产签核状态"].ToString() == "已签核")
                {
                    button1.Text = "已接收";
                    label25.Text = dt.Rows[0]["纸品生产"].ToString();
                }
                else
                {
                    button1.Text = "待接收确认";
                    label25.Text = "";
                }

                if (dt.Rows[0]["木铁生产签核状态"].ToString() == "已签核")
                {
                    button2.Text = "已接收";
                    label26.Text = dt.Rows[0]["木铁生产"].ToString();
                }
                else
                {
                    button2.Text = "待接收确认";
                    label26.Text = "";
                }

                if (dt.Rows[0]["亚克力生产签核状态"].ToString() == "已签核")
                {
                    button3.Text = "已接收";
                    label27.Text = dt.Rows[0]["亚克力生产"].ToString();
                }
                else
                {
                    button3.Text = "待接收确认";
                    label27.Text = "";
                }

                if (dt.Rows[0]["纸品计划签核状态"].ToString() == "已签核")
                {
                    button4.Text = "已接收";
                    label28.Text = dt.Rows[0]["纸品计划"].ToString();

                }
                else
                {
                    button4.Text = "待接收确认";
                    label28.Text = "";
                }
                if (dt.Rows[0]["木铁计划签核状态"].ToString() == "已签核")
                {
                    button5.Text = "已接收";
                    label29.Text = dt.Rows[0]["木铁计划"].ToString();
                }
                else
                {
                    button5.Text = "待接收确认";
                    label29.Text = "";
                }

                if (dt.Rows[0]["结构设计签核状态"].ToString() == "已签核")
                {
                    button6.Text = "已接收";
                    label59.Text = dt.Rows[0]["结构设计"].ToString();
                }
                else
                {
                    button6.Text = "待接收确认";
                    label59.Text = "";
                }

                if (dt.Rows[0]["平面设计签核状态"].ToString() == "已签核")
                {
                    button7.Text = "已接收";
                    label61.Text = dt.Rows[0]["平面设计"].ToString();
                }
                else
                {
                    button7.Text = "待接收确认";
                    label61.Text = "";
                }

                if (dt.Rows[0]["纸品采购签核状态"].ToString() == "已签核")
                {
                    button8.Text = "已接收";
                    label63.Text = dt.Rows[0]["纸品采购"].ToString();

                }
                else
                {
                    button8.Text = "待接收确认";
                    label63.Text = "";
                }
                if (dt.Rows[0]["木铁采购签核状态"].ToString() == "已签核")
                {
                    button9.Text = "已接收";
                    label65.Text = dt.Rows[0]["木铁采购"].ToString();
                }
                else
                {
                    button9.Text = "待接收确认";
                    label65.Text = "";
                }
                label5.Text = "下单日期 " + dt.Rows[0]["下单日期"].ToString();
                label6.Text = "修改日期 " + dt.Rows[0]["修改日期"].ToString();
               
            }
            dgvStateControl();
            bind2();
        }
        #endregion
        private void hint_bind()
        {
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
            dataGridView1.AllowUserToAddRows = false;
            if (dataGridView1.Rows.Count > 0)
            {
             
                dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
                int numCols2 = dataGridView1.Columns.Count;
                dataGridView1.Columns["复选框"].Width = 50;
                dataGridView1.Columns["文件名"].Width = 130;
                dataGridView1.Columns["索引"].Width = 130;

                for (i = 0; i < numCols2; i++)
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
                dataGridView1.Columns["文件名"].ReadOnly = true;
                dataGridView1.Columns["索引"].ReadOnly = true;
            }
  
        }
        #endregion

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
                string INITIAL_MAKERID = "";
                dt = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
                if (dt.Rows.Count > 0)
                {
                    INITIAL_MAKERID =dt.Rows [0]["制单人编号"].ToString ();
                    if (EDIT != "有权限" && LOGIN.EMID != INITIAL_MAKERID)
                    {
                        hint.Text = "本账号无修改权限！";
                        return;
                    }
                }
                if (juage())
                {

                    IFExecution_SUCCESS = false;
                }
                else if (checkBox1.Checked == false && checkBox2.Checked == false && checkBox3.Checked == false &&
                checkBox4.Checked == false && checkBox5.Checked == false && checkBox6.Checked == false && 
                checkBox7.Checked == false &&checkBox8.Checked == false && checkBox9.Checked == false)
                {
                    IFExecution_SUCCESS = false;
                    hint.Text = string.Format("至少要选择一种签核人");
                }
                else
                {

                    save();
                    if (IFExecution_SUCCESS == true && ADD_OR_UPDATE == "ADD")
                    {
                        //add();
                    }
                    bind();
                    F1.bind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #region SQlcommandE
        protected void SQlcommandE(string sql,string EMID,string NOTICE_OR_AUDIT)
        {

            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcon.Open();
            sqlcom.Parameters.Add("@RIID", SqlDbType.VarChar, 20).Value = IDO;
            sqlcom.Parameters.Add("@NOTICE_MAKERID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@RECEIVE_STATUS", SqlDbType.VarChar, 20).Value = "N";
            sqlcom.Parameters.Add("@NOTICE_OR_AUDIT", SqlDbType.VarChar, 20).Value = NOTICE_OR_AUDIT;
            sqlcom.Parameters.Add("@Date", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        private void add()
        {
            ClearText();
            IDO = cPN_PRODUCTION_INSTRUCTIONS.GETID();
            IFExecution_SUCCESS = false;
            bind();
            ADD_OR_UPDATE = "ADD";
        }

        private void save()
        {

            if (EDIT_TIMES != 0)//送签次数不为0表示已经送签过，此时要清空之签核记录，需重新送签 160928
            {
                basec.getcoms(@"
UPDATE PN_PRODUCTION_INSTRUCTIONS SET 
PAPER_PRODUCTION_AUDIT_STATUS='N',
PAPER_PRODUCTION_AUDIT_MAKERID='',
WOOD_IRON_PRODUCTION_AUDIT_STATUS='N',
WOOD_IRON_PRODUCTION_AUDIT_MAKERID='',
ACRYLIC_PRODUCTION_AUDIT_STATUS='N',
ACRYLIC_PRODUCTION_AUDIT_MAKERID='',
PAPER_PLAN_AUDIT_STATUS='N',
PAPER_PLAN_AUDIT_MAKERID='',
WOOD_IRON_PLAN_AUDIT_STATUS='N',
WOOD_IRON_PLAN_AUDIT_MAKERID='',
STRUCTURE_AUDIT_STATUS='N',
STRUCTURE_AUDIT_MAKERID='',
PLANE_AUDIT_STATUS='N',
PLANE_AUDIT_MAKERID='',
PAPER_PURCHASE_AUDIT_STATUS='N',
PAPER_PURCHASE_AUDIT_MAKERID='',
WOOD_IRON_PURCHASE_AUDIT_STATUS='N',
WOOD_IRON_PURCHASE_AUDIT_MAKERID='',
AUDIT_STATUS=''
WHERE PNID='" + IDO + "'");//清空之前的签核记录 160928
            }
            btnSave.Focus();
            cPN_PRODUCTION_INSTRUCTIONS.EMID = LOGIN.EMID;
            cPN_PRODUCTION_INSTRUCTIONS.PNID = IDO;
            cPN_PRODUCTION_INSTRUCTIONS.PFID = cno_paper_offer.RETURN_PFID_NPID(comboBox1.Text);
            cPN_PRODUCTION_INSTRUCTIONS.PIID = bc.getOnlyString("SELECT PIID FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox3 .Text  + "'");
            cPN_PRODUCTION_INSTRUCTIONS.ORDER_DATE = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
            cPN_PRODUCTION_INSTRUCTIONS.EDIT_DATE = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
            cPN_PRODUCTION_INSTRUCTIONS.EDIT_TIMES = EDIT_TIMES.ToString();
            string yyMM = dateTimePicker1.Text.Substring(2, 2) + dateTimePicker1.Text.Substring(5, 2);
          
            cPN_PRODUCTION_INSTRUCTIONS.IF_AUDIT_PRICE = comboBox13.Text;
            cPN_PRODUCTION_INSTRUCTIONS.WAREID = textBox2.Text;
            cPN_PRODUCTION_INSTRUCTIONS.PRODUCTION_COUNT = textBox4.Text;
            cPN_PRODUCTION_INSTRUCTIONS.HAVE_TAX_UNIT_PRICE = textBox5.Text;
            cPN_PRODUCTION_INSTRUCTIONS.DELIVERY_DATE = dateTimePicker1.Text;
            cPN_PRODUCTION_INSTRUCTIONS.DELIVERY_BATCH = textBox6.Text;
            cPN_PRODUCTION_INSTRUCTIONS.DELIVERY_PLACE = textBox7.Text;
            cPN_PRODUCTION_INSTRUCTIONS.ORDER_TYPE = comboBox2.Text;
            cPN_PRODUCTION_INSTRUCTIONS.PACKING_METHOD = textBox19.Text;
            cPN_PRODUCTION_INSTRUCTIONS.OUTSIDE_BOX_MATERIAL = textBox20.Text;
            cPN_PRODUCTION_INSTRUCTIONS.OUTSIDE_BOX_LONG = textBox21.Text;
            cPN_PRODUCTION_INSTRUCTIONS.OUTSIDE_BOX_WIDTH = textBox22.Text;
            cPN_PRODUCTION_INSTRUCTIONS.OUTSIDE_BOX_HEIGHT = textBox23.Text;
            cPN_PRODUCTION_INSTRUCTIONS.OUTSIDE_BOX_WEIGHT = textBox24.Text;
            cPN_PRODUCTION_INSTRUCTIONS.INSTRUCTION_SIZE = textBox25.Text;
            cPN_PRODUCTION_INSTRUCTIONS.INSTRUCTION_REQUIRE = textBox26.Text;
            cPN_PRODUCTION_INSTRUCTIONS.MATTERS_NEEDING_ATTENTION = textBox27.Text;
            about_audit();
            cPN_PRODUCTION_INSTRUCTIONS.PAPER_PRODUCTION_AUDIT_MAKERID = "";
            cPN_PRODUCTION_INSTRUCTIONS.WOOD_IRON_PRODUCTION_AUDIT_MAKERID = "";
            cPN_PRODUCTION_INSTRUCTIONS.ACRYLIC_PRODUCTION_AUDIT_MAKERID = "";
            cPN_PRODUCTION_INSTRUCTIONS.PAPER_PLAN_AUDIT_MAKERID = "";
            cPN_PRODUCTION_INSTRUCTIONS.WOOD_IRON_PLAN_AUDIT_MAKERID = "";
            cPN_PRODUCTION_INSTRUCTIONS.PAPER_PRODUCTION_AUDIT_STATUS = "N";
            cPN_PRODUCTION_INSTRUCTIONS.WOOD_IRON_PRODUCTION_AUDIT_STATUS = "N";
            cPN_PRODUCTION_INSTRUCTIONS.ACRYLIC_PRODUCTION_AUDIT_STATUS = "N";
            cPN_PRODUCTION_INSTRUCTIONS.PAPER_PLAN_AUDIT_STATUS = "N";
            cPN_PRODUCTION_INSTRUCTIONS.WOOD_IRON_PLAN_AUDIT_STATUS = "N";
            cPN_PRODUCTION_INSTRUCTIONS.SUBMIT_MAKERID = "";
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_PAPER_PRODUCTION_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PRODUCTION_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_WOOD_IRON_PRODUCTION_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PRODUCTION_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_ACRYLIC_PRODUCTION_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + ACRYLIC_PRODUCTION_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_PAPER_PLAN_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PLAN_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_WOOD_IRON_PLAN_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PLAN_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_STRUCTURE_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + STRUCTURE_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_PLANE_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PLANE_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_PAPER_PURCHASE_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PURCHASE_AUDIT_MAKERID + "'");
            cPN_PRODUCTION_INSTRUCTIONS.APPOINT_WOOD_IRON_PURCHASE_AUDIT_MAKERID = bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PURCHASE_AUDIT_MAKERID + "'");

         
            cPN_PRODUCTION_INSTRUCTIONS.save(yyMM,comboBox2.Text );
            IFExecution_SUCCESS = cPN_PRODUCTION_INSTRUCTIONS.IFExecution_SUCCESS;
            hint.Text = cPN_PRODUCTION_INSTRUCTIONS.ErrowInfo;
            try
            {

            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region about_notice
        private void  about_notice()
        {
            if (juage())
            {
            }
            else
            {
                if (checkBox1.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PRODUCTION_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox2.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PRODUCTION_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox3.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + ACRYLIC_PRODUCTION_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox4.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PLAN_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox5.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PLAN_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox6.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + STRUCTURE_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox7.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PLANE_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox8.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PURCHASE_AUDIT_MAKERID + "'"), "AUDIT");
                }
                if (checkBox9.Checked)
                {
                    SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PURCHASE_AUDIT_MAKERID + "'"), "AUDIT");
                }
            }
        }
        #endregion
        #region about_audit
        private void about_audit()
        {
            if (checkBox1.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PRODUCTION_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PRODUCTION_AUDIT = "N";
            }
            if (checkBox2.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PRODUCTION_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PRODUCTION_AUDIT = "N";
            }
            if (checkBox3.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_ACRYLIC_PRODUCTION_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_ACRYLIC_PRODUCTION_AUDIT = "N";
            }
            if (checkBox4.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PLAN_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PLAN_AUDIT = "N";
            }
            if (checkBox5.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PLAN_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PLAN_AUDIT = "N";
            }
            if (checkBox6.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_STRUCTURE_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_STRUCTURE_AUDIT = "N";
            }
            if (checkBox7.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PLANE_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PLANE_AUDIT = "N";
            }
            if (checkBox8.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PURCHASE_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PURCHASE_AUDIT = "N";
            }
            if (checkBox9.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PURCHASE_AUDIT = "Y";
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PURCHASE_AUDIT = "N";
            }
        }
        #endregion
        #region about_notice_old
        private void about_notice_old()
        {

            if (checkBox1.Checked)
            {

                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PRODUCTION_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PRODUCTION_AUDIT_MAKERID + "'"), "AUDIT");

            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PRODUCTION_AUDIT = "N";
            }

            if (checkBox2.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PRODUCTION_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PRODUCTION_AUDIT_MAKERID + "'"), "AUDIT");
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PRODUCTION_AUDIT = "N";
            }
            if (checkBox3.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_ACRYLIC_PRODUCTION_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + ACRYLIC_PRODUCTION_AUDIT_MAKERID + "'"), "AUDIT");
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_ACRYLIC_PRODUCTION_AUDIT = "N";
            }
            if (checkBox4.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PLAN_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PLAN_AUDIT_MAKERID + "'"), "AUDIT");

            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PLAN_AUDIT = "N";
            }
            if (checkBox5.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PLAN_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PLAN_AUDIT_MAKERID + "'"), "AUDIT");
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PLAN_AUDIT = "N";
            }
            if (checkBox6.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_STRUCTURE_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + STRUCTURE_AUDIT_MAKERID + "'"), "AUDIT");
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_STRUCTURE_AUDIT = "N";
            }
            if (checkBox7.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PLANE_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PLANE_AUDIT_MAKERID + "'"), "AUDIT");
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PLANE_AUDIT = "N";
            }
            if (checkBox8.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PURCHASE_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PURCHASE_AUDIT_MAKERID + "'"), "AUDIT");

            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_PAPER_PURCHASE_AUDIT = "N";
            }
            if (checkBox9.Checked)
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PURCHASE_AUDIT = "Y";
                SQlcommandE(sql, bc.getOnlyString("SELECT EMID FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PURCHASE_AUDIT_MAKERID + "'"), "AUDIT");
            }
            else
            {
                cPN_PRODUCTION_INSTRUCTIONS.IF_WOOD_IRON_PURCHASE_AUDIT = "N";
            }

        }
        #endregion
        #region juage
        private bool juage()
        {
            PAPER_PRODUCTION_AUDIT_MAKERID = "";
            WOOD_IRON_PRODUCTION_AUDIT_MAKERID = "";
            ACRYLIC_PRODUCTION_AUDIT_MAKERID = "";
            PAPER_PLAN_AUDIT_MAKERID = "";
            WOOD_IRON_PLAN_AUDIT_MAKERID = "";
            STRUCTURE_AUDIT_MAKERID = "";
            PLANE_AUDIT_MAKERID = "";
            PAPER_PURCHASE_AUDIT_MAKERID = "";
            WOOD_IRON_PURCHASE_AUDIT_MAKERID = "";
            dtx = cno_paper_offer.RETURN_PFID_NPID_DT(comboBox3.Text );
            dtx = bc.GET_DT_TO_DV_TO_DT(dtx, "", "报价编号='"+comboBox1 .Text +"'");
            decimal d1 = 0, d2 = 0;
            DataTable  dt = basec.getdts(caudit_list.sql);
            if (dt.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt.Rows[0]["纸品生产工号"].ToString()))
                {
                    PAPER_PRODUCTION_AUDIT_MAKERID =  dt.Rows[0]["纸品生产工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["木铁生产工号"].ToString()))
                {
                    WOOD_IRON_PRODUCTION_AUDIT_MAKERID =  dt.Rows[0]["木铁生产工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["亚克力生产工号"].ToString()))
                {
                    ACRYLIC_PRODUCTION_AUDIT_MAKERID = dt.Rows[0]["亚克力生产工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["纸品计划工号"].ToString()))
                {
                    PAPER_PLAN_AUDIT_MAKERID =  dt.Rows[0]["纸品计划工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["木铁计划工号"].ToString()))
                {
                    WOOD_IRON_PLAN_AUDIT_MAKERID = dt.Rows[0]["木铁计划工号"].ToString();
                }
                if (comboBox4.Text != "" )
                {
                    STRUCTURE_AUDIT_MAKERID = bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox4.Text, '-');
                }
                if (comboBox5.Text != "" )
                {
                    PLANE_AUDIT_MAKERID = bc.RETURN_FROM_RIGHT_UNTIL_CHAR(comboBox5.Text, '-');
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["纸品采购工号"].ToString()))
                {
                    PAPER_PURCHASE_AUDIT_MAKERID = dt.Rows[0]["纸品采购工号"].ToString();
                }
                if (!string.IsNullOrEmpty(dt.Rows[0]["木铁采购工号"].ToString()))
                {
                    WOOD_IRON_PURCHASE_AUDIT_MAKERID =dt.Rows[0]["木铁采购工号"].ToString();
                }

            }
            if (dtx.Rows.Count > 0)
            {
              
                if (!string.IsNullOrEmpty(dtx.Rows[0]["数量"].ToString()))
                {
                    d1 = decimal.Parse(dtx.Rows[0]["数量"].ToString());
                }
                if (!string.IsNullOrEmpty(dtx.Rows[0]["报出价"].ToString()))
                {
                    d2= decimal.Parse(dtx.Rows[0]["报出价"].ToString());
                }
            }
           bool b = false;

           //IDO = "PN16060001";
            if (IDO == null)
            {
                hint.Text = "编号不能为空";
                b = true;
            }
            else if (comboBox1.Text == "" && comboBox13.Text =="已报价")
            {
                hint.Text = "报价编号不能为空";
                b = true;
            }
            else if (comboBox3.Text == "")
            {
                hint.Text = "项目号不能为空";
                b = true;
            }
            else if (!bc.exists ("SELECT * FROM PROJECT_INFO WHERE PROJECT_ID='"+comboBox3.Text +"'"))
            {
                hint.Text = string .Format ("项目号: {0} 不存在系统中",comboBox3.Text );
                b = true;
            }
            else if (comboBox13.Text == "")
            {
                hint.Text = "报价不能为空";
                b = true;
            }
            else if (textBox2.Text == "")
            {
                hint.Text = "品号不能为空";
                b = true;
            }
            else if (comboBox13.Text != "已报价" && comboBox13.Text !="待报价")
            {
                hint.Text = "报价只能为已报价或待报价";
                b = true;
            }
            else if (comboBox13.Text =="已报价" && !cPN_PRODUCTION_INSTRUCTIONS .RETURN_OFFER_ID_IF_EXISTS(comboBox1.Text ))
            {
                b = true;
                hint.Text = string.Format("报价编号不存在系统中");

            }
            else if (textBox4.Text == "")
            {
                hint.Text = "数量不能为空";
                b = true;
            }
            else if (bc.yesno(textBox4.Text) == 0)
            {
                hint.Text = "数量只能输入数字";
                b = true;
            }
            else if (comboBox13.Text == "已报价" && decimal.Parse(textBox4.Text) < d1)
            {

                MessageBox.Show("已报价时生产数量" + textBox4.Text + "未大于等于报价数量" + d1.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (comboBox13.Text  =="已报价" &&  textBox5.Text=="")
            {
                hint.Text = "已报价时含税单价不能为空";
                b = true;
            }
            else if (comboBox13.Text == "已报价" && decimal.Parse(textBox5.Text) < d2)
            {
                MessageBox.Show("已报价时含税单价"+textBox5.Text +"未大于等于报出价"+d2.ToString (),"提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                b = true;
            }
            else if (bc.yesno(textBox5.Text) == 0)
            {
                hint.Text = "含税单价只能输入数字";
                b = true;
            }
            else if (textBox6.Text == "")
            {
                hint.Text = "交货批次不能为空";
                b = true;
            }
            else if (textBox7.Text == "")
            {
                hint.Text = "交货地点不能为空";
                b = true;
            }
            else if (comboBox2.Text == "")
            {
                hint.Text = "订单类型不能为空";
                b = true;
            }
            else if (textBox21.Text != "" && bc.yesno(textBox21.Text) == 0)
            {
                hint.Text = "长只能输入数字";
                b = true;
            }
            else if (textBox22.Text != "" && bc.yesno(textBox22.Text) == 0)
            {
                hint.Text = "宽只能输入数字";
                b = true;
            }
            else if (textBox23.Text != "" && bc.yesno(textBox23.Text) == 0)
            {
                hint.Text = "高只能输入数字";
                b = true;
            }
            else if (textBox24.Text != "" && bc.yesno(textBox24.Text) == 0)
            {
                hint.Text = "外箱重量只能输入数字";
                b = true;
            }
            else if (textBox25.Text == "")
            {
                hint.Text = "说明书尺寸不能为空";
                b = true;
            }
            else if (textBox26.Text == "")
            {
                hint.Text = "说明书要求不能为空";
                b = true;
            }
            else if (checkBox1.Checked && PAPER_PRODUCTION_AUDIT_MAKERID == "")
            {
                hint.Text = "选中纸品生产签核时，需指定工号";
                b = true;
            }
            else if (checkBox1.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PRODUCTION_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定纸品生产工号不存在系统";
                b = true;
            }
            else if (checkBox2.Checked && WOOD_IRON_PRODUCTION_AUDIT_MAKERID == "")
            {
                hint.Text = "选中木铁生产签核时，需指定工号";
                b = true;
            }
            else if (checkBox2.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PRODUCTION_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定木铁生产工号不存在系统";
                b = true;
            }
            else if (checkBox3.Checked && ACRYLIC_PRODUCTION_AUDIT_MAKERID == "")
            {
                hint.Text = "选中亚克力生产签核时，需指定工号";
                b = true;
            }
            else if (checkBox3.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + ACRYLIC_PRODUCTION_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定亚克力生产工号不存在系统";
                b = true;
            }
            else if (checkBox4.Checked && PAPER_PLAN_AUDIT_MAKERID == "")
            {
                hint.Text = "选中纸品计划签核时，需指定工号";
                b = true;
            }
            else if (checkBox4.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PLAN_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定纸品计划工号不存在系统";
                b = true;
            }

            else if (checkBox5.Checked && WOOD_IRON_PLAN_AUDIT_MAKERID == "")
            {
                hint.Text = "选中木铁计划签核时，需指定工号";
                b = true;
            }
            else if (checkBox5.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PLAN_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定木铁计划工号不存在系统";
                b = true;
            }
            else if (checkBox6.Checked && STRUCTURE_AUDIT_MAKERID == "")
            {
                hint.Text = "选中结构设计签核时，需指定工号";
                b = true;
            }
            else if (checkBox6.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + STRUCTURE_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定结构设计工号不存在系统";
                b = true;
            }
            else if (checkBox7.Checked && PLANE_AUDIT_MAKERID == "")
            {
                hint.Text = "选中平面设计签核时，需指定工号";
                b = true;
            }
            else if (checkBox7.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PLANE_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定平面设计工号不存在系统";
                b = true;
            }
            else if (checkBox8.Checked && PAPER_PURCHASE_AUDIT_MAKERID == "")
            {
                hint.Text = "选中纸品采购签核时，需指定工号";
                b = true;
            }
            else if (checkBox8.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + PAPER_PURCHASE_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定纸品采购工号不存在系统";
                b = true;
            }

            else if (checkBox9.Checked && WOOD_IRON_PURCHASE_AUDIT_MAKERID == "")
            {
                hint.Text = "选中木铁采购签核时，需指定工号";
                b = true;
            }
            else if (checkBox9.Checked && !bc.exists("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='" + WOOD_IRON_PURCHASE_AUDIT_MAKERID + "'"))
            {
                hint.Text = "指定木铁采购工号不存在系统";
                b = true;
            }

           /*else if (juage5())
           {

            b = true;
            }*/
         
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE PNID='{0}'",bc.RETURN_PNID(textBox2 .Text ))))
            {
                hint.Text = string.Format("尺寸 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
           return b;
        }
        #endregion
        
        #region juage5

        private bool juage5()
        {
            bool b = false;
            if (cPN_PRODUCTION_INSTRUCTIONS.JUAGE_IF_AUDIT_END(IDO))
            {

            }
            else if (textBox9.Text != "" && bc.yesno(textBox9.Text) == 0)
            {

                b = true;
                hint.Text = string.Format("金属值只能输入数字");
            }
            else if (textBox11.Text != "" && bc.yesno(textBox11.Text) == 0)
            {
                b = true;
                hint.Text = string.Format("塑料值只能输入数字");
            }
            return b;
        }
        #endregion
        private bool juage6()
        {
            bool b = false;
            dtx = basec.getdts(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            string v1 = "";
            if (dtx.Rows.Count > 0)
            {
                v1 = dtx.Rows[0]["是否提交"].ToString();
            }
            if (v1 == "已提交")
            {
                hint.Text = "此生产指示书已提交不允许修改";
                b = true;
            }
    
            return b;
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
   
        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }
  
        private void btnSearch_Click(object sender, EventArgs e)
        {
            bind();
            
        }

        private void btnupload_Click(object sender, EventArgs e)
        {

            DataTable dty = bc.getdt("SELECT * FROM WAREFILE WHERE WAREID='" + textBox1.Text + "'" );
            if (juage())
            {

            }
            else if (dty.Rows.Count.ToString() == "6")
            {

                hint.Text = "最多只能上传三张图片";
            }
            else
            {
                uploadfile();
            }
            try
            {
          
            
            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region uploadfile
        private void uploadfile()
        {
            int i = 0;
            label53.Visible = false;
            label55.Visible = false;
            label56.Visible = false;
            label57.Visible = false;
            progressBar1.Visible = false;
            /*  string v2 = bc.getOnlyString("SELECT EDIT FROM RIGHTLIST WHERE USID='" + LOGIN.USID + "' AND NODE_NAME='传单作业'");
              if (v2 != "Y" && ADD_OR_UPDATE == "UPDATE")
              {
                  hint.Text = "您没有修改权限不能修改上传";
              }
              else*/
            label52.Text = "";
        
                OpenFileDialog openf = new OpenFileDialog();
             
                if (openf.ShowDialog() == DialogResult.OK)
                {
                   
                    Random ro = new Random();
                    string stro = ro.Next(80, 10000000).ToString() + "-";
                    string NeWAREID = DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString() + stro;

                    cfileinfo.SERVER_IP_OR_DOMAIN = bc.RETURN_SERVER_IP_OR_DOMAIN();
                    WATER_MARK_CONTENT = "";//水印内容
                    //cfileinfo.UploadImage(openf.FileName, Path.GetFileName(openf.FileName), textBox1 .Text );
                    //this.UploadFile(openf.FileName, System.IO.Path.GetFileName(openf.FileName), "File/", textBox1.Text);

                    string v21 = bc.FROM_RIGHT_UNTIL_CHAR(Path.GetFileName(openf.FileName), 46);
                    OLD_FILE_NAME = Path.GetFileName(openf.FileName);
                    NEW_FILE_NAME = NeWAREID + Path.GetFileName(openf.FileName);
              
                    //如果上传的是图片文件
                    if (v21 == "jpeg" || v21 == "jpg" || v21 == "JPG" || v21 == "png" || v21 == "bmp" || v21 == "gif")
                    {
                        //裁切小图
                        cfileinfo.MakeThumbnail(openf.FileName, "d:\\" + Path.GetFileName(openf.FileName), 80, 80, "Cut");
                        //小图加水印
                        cfileinfo.ADD_WATER_MARK("d:\\" + Path.GetFileName(openf.FileName), "d:\\80X80" + NeWAREID + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        //原图加水印
                        cfileinfo.ADD_WATER_MARK(openf.FileName, "d:\\INITIAL" + NeWAREID + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        INITIAL_OR_OTHER = "INITIAL";
                        label5.Text = "";
                        //上传原图
                        i = Upload_Request("http://"+bc.RETURN_SERVER_IP_OR_DOMAIN() +"/webuploadfile/default.aspx", "D:\\INITIAL" + NeWAREID + System.IO.Path.GetFileName(openf.FileName),
                                "INITIAL" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1 .Text  );

                        //上传80X80的缩略图
                        INITIAL_OR_OTHER = "80X80";
                        i = Upload_Request("http://"+bc.RETURN_SERVER_IP_OR_DOMAIN() +"/webuploadfile/default.aspx", "D:\\80X80" + NeWAREID + System.IO.Path.GetFileName(openf.FileName),
                                "80X80" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1 .Text  );


                        //删除本地临时水印图及剪切图
                        if (File.Exists("d:\\80X80" + NeWAREID + Path.GetFileName(openf.FileName)))
                        {
                            File.Delete("d:\\80X80" + NeWAREID + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\" + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\INITIAL" + NeWAREID + Path.GetFileName(openf.FileName));
                        }
                        if (i == 1)
                        {
                            label52.Text = "成功上传";
                        }
                        else
                        {
                            label52.Text = "上传失败";
                        }

                        bind2();
                    }
                    else
                    {

                        MessageBox.Show("只能上传图片格式为jpeg/jpg/png/bmp/gif", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        /*label53.Visible = true;避免上传了非指定的图片格式，如Ico,所以限制上传图片格式170329
                        label55.Visible = true;
                        label56.Visible = true;
                        label57.Visible = true;
                        progressBar1.Visible = true;
                        //上传的是非图片文件
                        INITIAL_OR_OTHER = "INITIAL";
                        i = Upload_Request("http://"+bc.RETURN_SERVER_IP_OR_DOMAIN()+"/webuploadfile/default.aspx", openf.FileName,
                                                     "INITIAL" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1 .Text  );*/
                    }
                  
                }
            

        }
        #endregion
        #region HttpWebRequst_uploadfile
        /// <summary>
        /// 将本地文件上传到指定的服务器(HttpWebRequest方法)
        /// </summary>
        /// <param name="address">文件上传到的服务器</param>
        /// <param name="fileNamePath">要上传的本地文件（全路径）</param>
        /// <param name="saveName">文件上传后的名称</param>
        /// <param name="progressBar">上传进度条</param>
        /// <returns>成功返回1，失败返回0</returns>
        /// 
        #region Upload_Request
        public int Upload_Request(string address, string fileNamePath, string saveName, ProgressBar progressBar, string WAREID)
        {
            int returnValue = 0;
            // 要上传的文件

            FileStream fs = new FileStream(fileNamePath, FileMode.Open, FileAccess.Read);
            BinaryReader r = new BinaryReader(fs);
            //时间戳
            string strBoundary = "----------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + strBoundary + "\r\n");
            //请求头部信息
            StringBuilder sb = new StringBuilder();
            sb.Append("--");
            sb.Append(strBoundary);
            sb.Append("\r\n");
            sb.Append("Content-Disposition: form-data; name=\"");
            sb.Append("file");
            sb.Append("\"; filename=\"");
            sb.Append(saveName);
            sb.Append("\"");
            sb.Append("\r\n");
            sb.Append("Content-Type: ");
            sb.Append("application/octet-stream");
            sb.Append("\r\n");
            sb.Append("\r\n");
            string strPostHeader = sb.ToString();


            byte[] postHeaderBytes = Encoding.UTF8.GetBytes(strPostHeader);
            // 根据uri创建HttpWebRequest对象
            HttpWebRequest httpReq = (HttpWebRequest)WebRequest.Create(new Uri(address));
            httpReq.Method = "POST";
            //对发送的数据不使用缓存
            httpReq.AllowWriteStreamBuffering = false;
            //设置获得响应的超时时间（300秒）
            httpReq.Timeout = 300000;
            httpReq.ContentType = "multipart/form-data; boundary=" + strBoundary;
            long length = fs.Length + postHeaderBytes.Length + boundaryBytes.Length;
            long fileLength = fs.Length;
            httpReq.ContentLength = length;
            if (fileLength / 1048576.0 > 2.5)
            {
               
                label52.Visible = false;
                label53.Visible = false;
                label55.Visible = false;
                label56.Visible = false;
                label57.Visible = false;
                progressBar1.Visible = false;
                MessageBox.Show("上传的图片长度为:" + (fileLength / 1048576.0).ToString("F2") + "M" + " 已经大于允许上传的2.5M");
            }
            else 
            {
            try
            {
                progressBar.Maximum = int.MaxValue;
                progressBar.Minimum = 0;
                progressBar.Value = 0;
                //每次上传4k
                int bufferLength = 4096;
                byte[] buffer = new byte[bufferLength];
                //已上传的字节数
                long offset = 0;
                //开始上传时间
                DateTime startTime = DateTime.Now;
                int size = r.Read(buffer, 0, bufferLength);
               
                Stream postStream = httpReq.GetRequestStream();
                //发送请求头部消息
                postStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
                while (size > 0)
                {
                    postStream.Write(buffer, 0, size);
                    offset += size;
                    progressBar.Value = (int)(offset * (int.MaxValue / length));
                    TimeSpan span = DateTime.Now - startTime;
                    double second = span.TotalSeconds;
                    label53.Text = "已用时：" + second.ToString("F2") + "秒";

                    if (second > 0.001)
                    {
                        label55.Text = "平均速度：" + (offset / 1024 / second).ToString("0.00") + "KB/秒";
                    }
                    else
                    {
                        label55.Text = "正在连接…";
                    }
                    label56.Text = "已上传：" + (offset * 100.0 / length).ToString("F2") + "%";
                    label57.Text = (offset / 1048576.0).ToString("F2") + "M/" + (fileLength / 1048576.0).ToString("F2") + "M";
                    Application.DoEvents();
                    size = r.Read(buffer, 0, bufferLength);
                }
                //添加尾部的时间戳
                postStream.Write(boundaryBytes, 0, boundaryBytes.Length);
                postStream.Close();

                string year = DateTime.Now.ToString("yy");
                string month = DateTime.Now.ToString("MM");
                string day = DateTime.Now.ToString("dd");
                string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                string v1 = bc.numYMD(20, 12, "000000000001", "SELECT * FROM WAREFILE", "FLKEY", "FL");
                string newFileName, uriString;
                newFileName = System.IO.Path.GetFileName(saveName);
                uriString = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN()  + "/uploadfile/" + newFileName;


                String sql = @"
INSERT INTO  WAREFILE 
(
FLKEY,
WAREID,
OLD_FILE_NAME,
NEW_FILE_NAME,
PATH,
INITIAL_OR_OTHER,
DATE,
YEAR,
MONTH,
DAY
) 
VALUES
(
@FLKEY,
@WAREID,
@OLD_FILE_NAME,
@NEW_FILE_NAME,
@PATH,
@INITIAL_OR_OTHER,
@DATE,
@YEAR,
@MONTH,
@DAY

)";
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                sqlcom.Parameters.Add("@FLKEY", SqlDbType.VarChar, 20).Value = v1;
                sqlcom.Parameters.Add("@WAREID", SqlDbType.VarChar, 20).Value = IDO;
                sqlcom.Parameters.Add("@OLD_FILE_NAME", SqlDbType.VarChar, 100).Value = OLD_FILE_NAME;
                sqlcom.Parameters.Add("@NEW_FILE_NAME", SqlDbType.VarChar, 100).Value = NEW_FILE_NAME;
                sqlcom.Parameters.Add("@PATH", SqlDbType.VarChar, 100).Value = uriString;
                sqlcom.Parameters.Add("@INITIAL_OR_OTHER", SqlDbType.VarChar, 100).Value = INITIAL_OR_OTHER;
                sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcon.Open();
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();


                //获取服务器端的响应
                WebResponse webRespon = httpReq.GetResponse();
                Stream s = webRespon.GetResponseStream();
                StreamReader sr = new StreamReader(s);
                //读取服务器端返回的消息
                String sReturnString = sr.ReadLine();
                s.Close();
                sr.Close();
                if (sReturnString == "Success")
                {
                    returnValue = 1;
                }
                else if (sReturnString == "Error")
                {
                    returnValue = 0;
                }
            }
            catch
            {
                returnValue = 0;
            }
            finally
            {
                fs.Close();
                r.Close();
            }
            }
            return returnValue;
        }
        #endregion
        #endregion
        #region bind2
        private void bind2()
        {
           
            dt3 = bc.getdt(@"
SELECT cast(0   as   bit)   as   复选框,
OLD_FILE_NAME AS 文件名,NEW_FILE_NAME AS 新文件名,FLKEY AS 索引,
PATH FROM WAREFILE WHERE WAREID='"+IDO +"'  AND INITIAL_OR_OTHER='80X80'");


            dataGridView1.Rows.Clear();//在下一次增加行前需清空上一次产生的行，否则显示行数不正常
            for (int i = 0; i < dt3.Rows.Count; i++)
            {
                
                DataGridViewRow dgr = new DataGridViewRow();
                dataGridView1.Rows.Add(dgr);
                dataGridView1["复选框", i].Value = false;
                dataGridView1["文件名", i].Value = dt3.Rows[i]["文件名"].ToString();
                dataGridView1["缩略图", i].Value = Image.FromStream(System.Net.WebRequest.Create(dt3.Rows[i]["PATH"].ToString()).GetResponse().GetResponseStream());
                dataGridView1["索引", i].Value = dt3.Rows[i]["索引"].ToString();
              
            }
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Height = 80;
            }
            this.WindowState = FormWindowState.Maximized;
            Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");

            dgvStateControl();
        }
        #endregion
        #region btndelfile
        private void btndelfile_Click(object sender, EventArgs e)
        {
          
            try
            {
                /*string v21 = bc.getOnlyString("SELECT EDIT FROM RIGHTLIST WHERE USID='" + LOGIN.USID + "' AND NODE_NAME='传单作业'");
                if (v21 != "Y" && ADD_OR_UPDATE == "UPDATE")
                {
                    hint.Text = "您没有修改权限不能删除文件";
                }
                else if (vou.CheckIfALLOW_SAVEOR_DELETE(textBox1.Text, LOGIN.USID))
                {
                    hint.Text = vou.ErrowInfo;
                }
                else
                {
                

                }*/
                if (MessageBox.Show("确定要删除该文件吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    if (dt3.Rows.Count > 0)
                    {

                        for (int i = 0; i < dt3.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
                            {

                                string v2 = dt3.Rows[i]["索引"].ToString();
                                string v4 = dt3.Rows[i]["新文件名"].ToString();
                                bc.getcom(@"INSERT INTO SERVER_DELETE_FILE(FLKEY,NEW_FILE_NAME) VALUES ('" + v2 + "','" + v4 + "')");
                                bc.getcom("DELETE WAREFILE WHERE NEW_FILE_NAME='" +v4 + "'");
                              
                            }
                        }
                        bind2();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
   

        #region  ACTIVE_DISPALY_INF
        private void  ACTIVE_DISPALY_INFO()
        {
            right();
            if (cPN_PRODUCTION_INSTRUCTIONS.JUAGE_IF_AUDIT_END(IDO))
            {
               
                INITIAL_MAKERID = bc.getOnlyString("SELECT MAKERID FROM PN_PRODUCTION_INSTRUCTIONS WHERE PNID='" + IDO + "'");
                SQlcommandE(sql, INITIAL_MAKERID, "NOTICE");
                dtx = bc.getdt("SELECT * FROM NOTICE_LIST");
                if (dtx.Rows.Count > 0)
                {
                    foreach (DataRow dr in dtx.Rows)
                    {
                        SQlcommandE(sql, dr["EMID"].ToString(), "NOTICE");
                    }
                }
                bind();
            }
        }
        #endregion

        private void pictureBox4_Click(object sender, EventArgs e)
        {
         
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS .sql  + " WHERE A.PNID='" + IDO  + "'");
            if (dtx.Rows.Count > 0)
            {


                EDIT_TIMES = 1 + EDIT_TIMES;//送签次数加一
                basec.getcoms(@"
UPDATE PN_PRODUCTION_INSTRUCTIONS SET 
EDIT_TIMES='" + EDIT_TIMES.ToString() +
          "',AUDIT_STATUS='SEND' WHERE PNID='" + IDO + "'");
                if (!bc.exists("SELECT * FROM PN_PRODUCTION_INSTRUCTIONS WHERE PNID='" + IDO + "'"))//如果不存在此单号就新增通知签核 160928
                {
                    about_notice();
                }
                else
                {

                    about_notice();//再次通知重新签核 160928
                    bind();
                }
            }
            else
            {
                hint.Text = "先保存单据才能送签";
            }

        }


        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
        




            try
            {
                DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
                if (juage())
                {
                }
                else if (dtx.Rows.Count > 0)
                {

                    cPN_PRODUCTION_INSTRUCTIONS.ExcelPrint(dtx, "生产指示书", System.IO.Path.GetFullPath("生产指示书.xlsx"));
                }
                else
                {
                    hint.Text = "先保存单据才能导出";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }



         
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    if (bc.exists("SELECT * FROM INVENTORY_MST WHERE PNID='"+IDO+"'"))
                    {
                        MessageBox.Show("此订单编号已经存在库存维护作业中，不允许删除","",MessageBoxButtons .OK,MessageBoxIcon.Warning);
                    }
                    else
                    {
                        basec.getcoms("DELETE PN_PRODUCTION_INSTRUCTIONS WHERE PNID='" + IDO + "'");
                        basec.getcoms("DELETE REMIND WHERE RIID='"+IDO +"'");/*同时删除通信息避免无法更新状态一直弹出通知170309*/
                        add();
                        F1.load();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button1.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PAPER_PRODUCTION_AUDIT_STATUS='Y',
PAPER_PRODUCTION_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label25.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button1.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PAPER_PRODUCTION_AUDIT_STATUS='N',
PAPER_PRODUCTION_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label25.Text = "";
                        button1.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
          
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button2.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET WOOD_IRON_PRODUCTION_AUDIT_STATUS='Y',WOOD_IRON_PRODUCTION_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label26.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button2.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET WOOD_IRON_PRODUCTION_AUDIT_STATUS='N',WOOD_IRON_PRODUCTION_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label26.Text = "";
                        button2.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button3.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET ACRYLIC_PRODUCTION_AUDIT_STATUS='Y',ACRYLIC_PRODUCTION_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label27.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button3.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET ACRYLIC_PRODUCTION_AUDIT_STATUS='N',ACRYLIC_PRODUCTION_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label27.Text = "";
                        button3.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button4.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PAPER_PLAN_AUDIT_STATUS='Y',PAPER_PLAN_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label28.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button4.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PAPER_PLAN_AUDIT_STATUS='N',PAPER_PLAN_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label28.Text = "";
                        button4.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }
      

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
          
            try
            {
                DataTable dtx1 = cno_paper_offer.RETURN_PFID_NPID_DT(comboBox3.Text );
                comboBox1.Items.Clear();
                if (dtx1.Rows.Count > 0)
                {
                   
                    foreach (DataRow dr in dtx1.Rows)
                    {
                        comboBox1.Items.Add(dr["报价编号"].ToString());
                    }
                }

            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
       
            try
            {
                dtx = cno_paper_offer.RETURN_PFID_NPID_DT_FROM_OFFER_ID(comboBox1.Text);
                if (dtx.Rows.Count > 0)
                {
                    textBox28.Text = dtx.Rows[0]["数量"].ToString();
                    textBox29.Text = dtx.Rows[0]["报出价"].ToString();
                }
                else
                {
                    textBox28.Text = "";
                    textBox29.Text = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
      
        }
        private void cleartext2()
        {
            
            comboBox1.Text = "";
            textBox3.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox13.Text = "";
            textBox16.Text = "";
            textBox11.Text = "";
            textBox14.Text = "";
            textBox17.Text = "";
            textBox12.Text = "";
            textBox18.Text = "";
            textBox18.Text = "";
            dateTimePicker1.Text = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");

        }
        private void button5_Click(object sender, EventArgs e)
        {
          
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button5.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET WOOD_IRON_PLAN_AUDIT_STATUS='Y',WOOD_IRON_PLAN_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label29.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button5.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET WOOD_IRON_PLAN_AUDIT_STATUS='N',WOOD_IRON_PLAN_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label29.Text = "";
                        button5.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void comboBox3_DropDown(object sender, EventArgs e)
        {
            CSPSS.OFFER_MANAGE.PROJECT_INFO FRM = new OFFER_MANAGE.PROJECT_INFO();
            FRM.WindowState = FormWindowState.Normal;
            FRM.PN_PRODUCTION_INSTRUCTIONS();
            FRM.ShowDialog();
            this.comboBox3.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox3.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox3.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox3.Text = GET_PROJECT_ID;

            }
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dtt = bc.getdt(cproject_info.sql + " WHERE A.PROJECT_ID='"+comboBox3.Text +"'");
                if (dtt.Rows.Count > 0)
                {
                    comboBox3.Text = dtt.Rows[0]["项目号"].ToString();
                    textBox3.Text = dtt.Rows[0]["项目名称"].ToString();
                    textBox9.Text = dtt.Rows[0]["品牌"].ToString();
                    textBox10.Text = dtt.Rows[0]["AE01"].ToString();
                    textBox13.Text = dtt.Rows[0]["AE02"].ToString();
                    textBox16.Text = dtt.Rows[0]["AE03"].ToString();
                    textBox11.Text = dtt.Rows[0]["平面01"].ToString();
                    textBox14.Text = dtt.Rows[0]["平面02"].ToString();
                    textBox17.Text = dtt.Rows[0]["平面03"].ToString();
                    textBox12.Text = dtt.Rows[0]["结构01"].ToString();
                    textBox15.Text = dtt.Rows[0]["结构02"].ToString();
                    textBox18.Text = dtt.Rows[0]["结构03"].ToString();
                    comboBox3.Text = dtt.Rows[0]["项目号"].ToString();
                    textBox8.Text = dtt.Rows[0]["客户名称"].ToString();
                    //label40.Text = dtt.Rows[0]["审核状态"].ToString();
                }
                else
                {
                    cleartext2();
                }
            }
             catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
            
           
        }
        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            try
            {
                int i = dataGridView1.CurrentCell.RowIndex;

                if (dataGridView1.CurrentCell.ColumnIndex == 1)
                {
                    SaveFileDialog sfl = new SaveFileDialog();
                    sfl.FileName = dt3.Rows[dataGridView1.CurrentCell.RowIndex]["文件名"].ToString();
                    sfl.DefaultExt = "jpg";
                    sfl.Filter = "(*.jpg)|*.jpg";
                    if (sfl.ShowDialog() == DialogResult.OK)
                    {
                        sqb = new StringBuilder();
                        sqb.AppendFormat("SELECT PATH FROM WAREFILE WHERE ");
                        sqb.AppendFormat(" NEW_FILE_NAME='{0}'", dt3.Rows[i]["新文件名"].ToString());
                        sqb.AppendFormat(" AND INITIAL_OR_OTHER='INITIAL'");
                        WebClient wclient = new WebClient();
                        string v1 = bc.getOnlyString(sqb.ToString ());
                        wclient.DownloadFile(v1, sfl.FileName);

                        /*DataTable dt3x = bc.getdt("SELECT * FROM WAREFILE WHERE FLKEY='" + dt3.Rows[dataGridView1.CurrentCell.RowIndex]["索引"].ToString() + "'");
                        Byte[] byte2 = (byte[])dt3x.Rows[0]["IMAGE_DATA"];
                        System.IO.File.WriteAllBytes(sfl.FileName, byte2);*/
                        hint.Text = "已下载";
                    }
                }
         
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }



   

  



        private void button6_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button6.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET STRUCTURE_AUDIT_STATUS='Y',STRUCTURE_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label59.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button6.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET STRUCTURE_AUDIT_STATUS='N',STRUCTURE_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label59.Text = "";
                        button6.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button7.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PLANE_AUDIT_STATUS='Y',PLANE_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label61.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button7.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PLANE_AUDIT_STATUS='N',PLANE_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label59.Text = "";
                        button7.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button8.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PAPER_PURCHASE_AUDIT_STATUS='Y',PAPER_PURCHASE_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label63.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button8.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET PAPER_PURCHASE_AUDIT_STATUS='N',PAPER_PURCHASE_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label63.Text = "";
                        button8.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cPN_PRODUCTION_INSTRUCTIONS.sql + " WHERE A.PNID='" + IDO + "'");
            if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button9.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET WOOD_IRON_PURCHASE_AUDIT_STATUS='Y',WOOD_IRON_PURCHASE_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE PNID='" + IDO + "'");
                        label65.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button9.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE PN_PRODUCTION_INSTRUCTIONS SET WOOD_IRON_PURCHASE_AUDIT_STATUS='N',WOOD_IRON_PURCHASE_AUDIT_MAKERID='' WHERE PNID='" + IDO + "'");
                        label65.Text = "";
                        button9.Text = "待接收确认";
                        F1.bind();
                    }
                    ACTIVE_DISPALY_INFO();
                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox4_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.POSITION = "结构设计";
            FRM.PN_PRODUCTION_INSTRUCTIONST_USE();
            FRM.ShowDialog();
            this.comboBox4.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox4.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox4.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox4.Text = ENAME + "-" + EMPLOYEE_ID;
            }
        }

        private void comboBox5_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.POSITION = "平面设计";
            FRM.PN_PRODUCTION_INSTRUCTIONST_USE();
            FRM.ShowDialog();
            this.comboBox5.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox5.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox5.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {
                comboBox5.Text = ENAME+"-"+EMPLOYEE_ID;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opfv = new OpenFileDialog();
                if (opfv.ShowDialog() == DialogResult.OK)
                {
                    /*DataSet ds = new DataSet();
                    string tablename = ExcelToCSHARP.GetExcelFirstTableName(opfv .FileName );
                    ds = ExcelToCSHARP.importExcelToDataSet(opfv .FileName , tablename);
                    DataTable dt = ds.Tables[0];
                    dataGridView1.DataSource = dt;*/
                    string path = opfv.FileName;
                    ExcelToCSHARP etc = new ExcelToCSHARP();
                    etc.EMID = LOGIN.EMID;
                    etc.showdata(path);
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
