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

namespace CSPSS.OFFER_MANAGE
{
    public partial class SAMPLE_RELY_LISTT : Form
    {
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt3 = new DataTable();
        basec bc=new basec ();
        CMATERIAL_PRICE cmaterial_price = new CMATERIAL_PRICE();
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
        private string _WATER_MARK_CONTENT;
        public string WATER_MARK_CONTENT
        {
            set { _WATER_MARK_CONTENT = value; }
            get { return _WATER_MARK_CONTENT; }

        }
         private string _AE_MAKERID_ONE;
        public string AE_MAKERID_ONE
        {
            set { _AE_MAKERID_ONE = value; }
            get { return _AE_MAKERID_ONE; }

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
        #endregion
        CFileInfo cfileinfo = new CFileInfo();
        CEMPLOYEE_INFO cemployee_info = new CEMPLOYEE_INFO();
        private  delegate bool dele(string a1,string a2);
        private delegate void delex();
        SAMPLE_RELY_LIST F1 = new SAMPLE_RELY_LIST();
        protected int M_int_judge, i;
        protected int select;
        CSAMPLE_RELY_LIST cSAMPLE_RELY_LIST = new CSAMPLE_RELY_LIST();
        CPROCESSING_TECHNOLOGY cprocessing_technology = new CPROCESSING_TECHNOLOGY();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        StringBuilder sqb = new StringBuilder();
          public SAMPLE_RELY_LISTT(SAMPLE_RELY_LIST  FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        public SAMPLE_RELY_LISTT()
        {
            InitializeComponent();
        }
      
        private void SAMPLE_RELY_LISTT_Load(object sender, EventArgs e)
        {
            try
            {
               
                //IDO = "SR15120263";//样板计费调用方法 RETURN_SAMPLE_PRICE()
                right();
              
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
                label52.Text = "";
                label53.Text = "";
                label57.Text = "";
                label55.Text = "";
                label56.Text = "";
                progressBar1.Visible = false;
                textBox1.ReadOnly = true;
                textBox4.ScrollBars = ScrollBars.Both;
                textBox6.ScrollBars = ScrollBars.Both;
                label6.Font = new Font("", 9, FontStyle.Bold);
                label7.Font = new Font("", 9, FontStyle.Bold);
                label8.Font = new Font("", 9, FontStyle.Bold);
                label10.Font = new Font("", 9, FontStyle.Bold);
                label31.Font = new Font("", 9, FontStyle.Bold);
                checkedListBox1.MultiColumn = true;
                checkedListBox1.Height = 20;
                checkedListBox1.ColumnWidth = 100;
                checkedListBox1.CheckOnClick = true;
                checkedListBox2.MultiColumn = true;
                checkedListBox2.Height = 20;
                checkedListBox2.ColumnWidth = 100;
                checkedListBox2.CheckOnClick = true;
                checkedListBox3.MultiColumn = true;
                checkedListBox3.Height = 20;
                checkedListBox3.ColumnWidth = 100;
                checkedListBox3.CheckOnClick = true;
                checkedListBox4.MultiColumn = true;
                checkedListBox4.Height = 20;
                checkedListBox4.ColumnWidth = 100;
                checkedListBox4.CheckOnClick = true;
                checkedListBox5.MultiColumn = true;
                checkedListBox5.Height = 20;
                checkedListBox5.ColumnWidth = 100;
                checkedListBox5.CheckOnClick = true;
                dt = bc.getdt(cprocessing_technology.sql + " WHERE B.MATERIAL_TYPE='画面'");
                if (dt.Rows.Count > 0)
                {
                    checkedListBox1.Items.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        checkedListBox1.Items.Add(dr["工艺"].ToString());
                    }
                }
                dt = bc.getdt(cprocessing_technology.sql + " WHERE B.MATERIAL_TYPE='纸品'");
                if (dt.Rows.Count > 0)
                {
                    checkedListBox2.Items.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        checkedListBox2.Items.Add(dr["工艺"].ToString());
                    }
                }
                dt = bc.getdt(cprocessing_technology.sql + " WHERE B.MATERIAL_TYPE='金属'");
                if (dt.Rows.Count > 0)
                {
                    checkedListBox3.Items.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        checkedListBox3.Items.Add(dr["工艺"].ToString());
                    }
                }
                dt = bc.getdt(cprocessing_technology.sql + " WHERE B.MATERIAL_TYPE='亚克力'");
                if (dt.Rows.Count > 0)
                {
                    checkedListBox4.Items.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        checkedListBox4.Items.Add(dr["工艺"].ToString());
                    }
                }
                dt = bc.getdt(cprocessing_technology.sql + " WHERE B.MATERIAL_TYPE='木'");
                if (dt.Rows.Count > 0)
                {
                    checkedListBox5.Items.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        checkedListBox5.Items.Add(dr["工艺"].ToString());
                    }
                }
                dt = bc.getdt("SELECT * FROM SAMPLE_TECHNOLOGY WHERE SRID='" + IDO + "' AND MATERIAL_TYPE='画面'");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        for (int i = 0; i < checkedListBox1.Items.Count; i++)
                        {
                            if (dr["TECHNOLOGY"].ToString() == checkedListBox1.Items[i].ToString())
                            {
                                checkedListBox1.SetItemChecked(i, true);
                                break;

                            }
                        }
                    }
                }
                dt = bc.getdt("SELECT * FROM SAMPLE_TECHNOLOGY WHERE SRID='" + IDO + "' AND MATERIAL_TYPE='纸品'");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        for (int i = 0; i < checkedListBox2.Items.Count; i++)
                        {
                            if (dr["TECHNOLOGY"].ToString() == checkedListBox2.Items[i].ToString())
                            {
                                checkedListBox2.SetItemChecked(i, true);
                                break;

                            }
                        }
                    }
                }
                dt = bc.getdt("SELECT * FROM SAMPLE_TECHNOLOGY WHERE SRID='" + IDO + "' AND MATERIAL_TYPE='金属'");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        for (int i = 0; i < checkedListBox3.Items.Count; i++)
                        {
                            if (dr["TECHNOLOGY"].ToString() == checkedListBox3.Items[i].ToString())
                            {
                                checkedListBox3.SetItemChecked(i, true);
                                break;

                            }
                        }
                    }
                }
                dt = bc.getdt("SELECT * FROM SAMPLE_TECHNOLOGY WHERE SRID='" + IDO + "' AND MATERIAL_TYPE='亚克力'");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        for (int i = 0; i < checkedListBox4.Items.Count; i++)
                        {
                            if (dr["TECHNOLOGY"].ToString() == checkedListBox4.Items[i].ToString())
                            {
                                checkedListBox4.SetItemChecked(i, true);
                                break;

                            }
                        }
                    }
                }
                dt = bc.getdt("SELECT * FROM SAMPLE_TECHNOLOGY WHERE SRID='" + IDO + "' AND MATERIAL_TYPE='木'");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        for (int i = 0; i < checkedListBox5.Items.Count; i++)
                        {
                            if (dr["TECHNOLOGY"].ToString() == checkedListBox5.Items[i].ToString())
                            {
                                checkedListBox5.SetItemChecked(i, true);
                                break;

                            }
                        }
                    }
                }
                textBox6.BorderStyle = BorderStyle.FixedSingle;

                /*IDO = cSAMPLE_RELY_LIST.GETID();
                comboBox1.Text = "DBXM1510001";
                textBox3.Text = "1";
                radioButton4.Checked = true;
                checkBox3.Checked = true;*/
                //MessageBox.Show(IDO);
              this.Icon = Resource1.xz_200X200;
                comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
                dateTimePicker1.CustomFormat = "yyyy/MM/dd";
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker2.CustomFormat = "yyyy/MM/dd";
                dateTimePicker2.Format = DateTimePickerFormat.Custom;
                textBox2.BackColor = CCOLOR.qmhs;

                textBox11.BackColor = CCOLOR.lylfnp;
                textBox11.ForeColor = Color.White;
                textBox11.TextAlign = HorizontalAlignment.Right;
                if (PROJECT_ID != null)
                {
                    comboBox1.Text = PROJECT_ID;
                    textBox1.Text = PROJECT_NAME;
                }
                #endregion
                bind();
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
          
        }
        #region clearText
        public void ClearText()
        {
            textBox1.Text = "";
            comboBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            DateTime date1 = Convert.ToDateTime( DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/"));
            dateTimePicker1.Value = date1;
            dateTimePicker2.Value = date1;
            comboBox2.Text = "";
            checkBox1.Checked = false;
           
            comboBox3.Text = "";
      
            textBox4.Text = "";
     
            textBox6.Text = "";
       
            radioButton4.Checked = false;
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton7.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            comboBox10.Text = "";
            label34.Text = "";
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
          
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
          
            label12.Text = "";
            label25.Text = "";
            label26.Text = "";
            label27.Text = "";
            label28.Text = "";
            button4.Text = "待接收确认";
            label28.Text = "";
            button3.Text = "待接收确认";
            label27.Text = "";
            button2.Text = "待接收确认";
            label26.Text = "";
            button1.Text = "待接收确认";
            label25.Text = "";

        }
        #endregion
        #region right
        private void right()
        {
            dtx = cedit_right.RETURN_RIGHT_LIST("打样单新增", LOGIN.USID);
            btnAdd.Visible = false;
            btnSave.Visible = false;
            label15.Visible = false;
            label17.Visible = false;
            pictureBox4.Visible = false;
            label33.Visible = false;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            radioButton4.Enabled = false;
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
            radioButton7.Enabled = false;
            checkBox3.Enabled = false;
            checkBox4.Enabled = false;
            checkBox5.Enabled = false;
            textBox7.Enabled = false;
            textBox8.Enabled = false;
            textBox9.Enabled = false;
            textBox10.Enabled = false;
            btnDel.Visible = false;
            label36.Visible = false;
            btnupload.Visible = false;
            btndelfile.Visible = false;
            label13.Visible = false;
            label14.Visible = false;
        
    
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
                if (dtx.Rows[0]["样板审核"].ToString() == "有权限")
                {
                    pictureBox4.Visible = true;
                    label33.Visible = true;
                }
                if (dtx.Rows[0]["修改权限"].ToString() == "有权限")
                {
                    btnSave.Visible = true;
                    label15.Visible = true;
                    EDIT = "有权限";
                }
                if (dtx.Rows[0]["图片上传"].ToString() == "有权限")
                {
                    btnupload.Visible = true;
                    btndelfile.Visible = true;
                    label13.Visible = true;
                    label14.Visible = true;
                 
                }
                if (dtx.Rows[0]["纸品签核"].ToString() == "有权限")
                {
                    button1.Enabled = true;
                    button5.Visible = true;
                }
                if (dtx.Rows[0]["亚克力签核"].ToString() == "有权限")
                {
                    button2.Enabled = true;
                    button7.Visible = true;
                }
                if (dtx.Rows[0]["木铁签核"].ToString() == "有权限")
                {
                    button3.Enabled = true;
                    button6.Visible = true;
                    button8.Visible = true;
                }
                if (dtx.Rows[0]["采购签核"].ToString() == "有权限")
                {
                    button4.Enabled = true;
                    button8.Visible = true;
                }
            }
            if (cSAMPLE_RELY_LIST.JUAGE_IF_AUDIT_END(IDO))
            {

            }
            else
            {
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
                radioButton6.Enabled = true;
                radioButton7.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox5.Enabled = true;
                textBox7.Enabled = true;
                textBox8.Enabled = true;
                textBox9.Enabled = true;
                textBox10.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = true;
                button8.Enabled = true;

            }

        }
        #endregion
        #region bind
        private void bind()
        {
            label5.Text = "";
            label12.Text = "";
            label34.Text = "";
            label25.Text = "";
            label26.Text = "";
            label27.Text = "";
            label28.Text = "";
            radioButton7.Checked = false;
            comboBox1.Focus();
            textBox3.BackColor = CCOLOR.CUSTOMER_YELLOW;
            comboBox1.BackColor = CCOLOR.CUSTOMER_YELLOW;
            //hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
            hint_bind();
            DataTable dtx = basec.getdts(cSAMPLE_RELY_LIST.sql +" WHERE A.SRID='"+IDO+"'");
            if (dtx.Rows.Count > 0)
            {
                label5.Text = "打样制单：" + dtx.Rows[0]["制单人"].ToString();
                label5.ForeColor = Color.Blue;
                label5.Font = new Font("微软黑体",9,FontStyle.Regular );
                if (dtx.Rows[0]["项目状态"].ToString() == "已完成")
                {
                    pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/audit.png"));
                    label33.Text = "已完成";
                }
                else
                {
                    pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/61.png"));
                    label33.Text = "进行中";
                }
       
                comboBox1.Text = dtx.Rows[0]["项目号"].ToString();
              
                textBox1.Text = dtx.Rows[0]["项目名称"].ToString();
               
                textBox2.Text = dtx.Rows[0]["打样单号"].ToString();
                textBox3.Text = dtx.Rows[0]["需求数量"].ToString();
                dateTimePicker1.Text = dtx.Rows[0]["需求日期"].ToString();
                comboBox2.Text = dtx.Rows[0]["组别"].ToString();
                dateTimePicker2.Text = dtx.Rows[0]["下单日期"].ToString();
                if (dtx.Rows[0]["自购或采购"].ToString() == "自购")
                {
                    checkBox1.Checked = true;
                }
                else if (dtx.Rows[0]["自购或采购"].ToString() == "采购")
                {
                    checkBox2.Checked = true;
                }
                else if (dtx.Rows[0]["自购或采购"].ToString() == "自购与采购")
                {
                    checkBox1.Checked = true;
                    checkBox2.Checked = true;
                }
             
                comboBox3.Text =  dtx.Rows[0]["自购"].ToString();
                comboBox10.Text = dtx.Rows[0]["采购"].ToString();
                label12.Text = dtx.Rows[0]["自购工号"].ToString();
                label34.Text = dtx.Rows[0]["采购工号"].ToString();
                textBox4.Text = dtx.Rows[0]["自购说明"].ToString();
     
                textBox6.Text = dtx.Rows[0]["其他事项"].ToString();
                if (dtx.Rows[0]["陈列类型"].ToString() == "纸品")
                {
                    radioButton4.Checked = true;
                    textBox7.Text = dtx.Rows[0]["陈列数值"].ToString();
                }
                else if (dtx.Rows[0]["陈列类型"].ToString() == "金属")
                {
                    radioButton5.Checked = true;
                    textBox8.Text = dtx.Rows[0]["陈列数值"].ToString();

                }
                else if (dtx.Rows[0]["陈列类型"].ToString() == "木器")
                {
                    textBox9.Text = dtx.Rows[0]["陈列数值"].ToString();
                    radioButton6.Checked = true;
                }
                else if (dtx.Rows[0]["陈列类型"].ToString() == "塑料")
                {
                    textBox10.Text = dtx.Rows[0]["陈列数值"].ToString();
                    radioButton7.Checked = true;
                }
             
                if (dtx.Rows[0]["小POP"].ToString() == "已选")//no
                {
                 
                    checkBox3.Checked = true;
                   
                }
                if (dtx.Rows[0]["陈列架"].ToString() == "已选")
                {
                    checkBox4.Checked = true;
                }
                if (dtx.Rows[0]["堆头"].ToString() == "已选")
                {
                    checkBox5.Checked = true;
                }//have
             
                if (dtx.Rows[0]["是否需纸品签核"].ToString() == "是")
                {
                    checkBox7.Checked = true;
                }
                else
                {
                    checkBox7.Checked = false;
                }
                if (dtx.Rows[0]["是否需亚克力签核"].ToString() == "是")
                {
                    checkBox8.Checked = true;
                }
                else
                {
                    checkBox8.Checked = false;
                }
                if (dtx.Rows[0]["是否需木铁签核"].ToString() == "是")
                {
                    checkBox9.Checked = true;
                }
                else
                {
                    checkBox9.Checked = false;
                }
                if (dtx.Rows[0]["是否需采购签核"].ToString() == "是")
                {
                    checkBox10.Checked = true;
                }
                else
                {
                    checkBox10.Checked = false;
                }
                label25.Text =  dtx.Rows[0]["纸品"].ToString();
                label26.Text =  dtx.Rows[0]["亚克力"].ToString();
                label27.Text = dtx.Rows[0]["木铁"].ToString();
                label28.Text =  dtx.Rows[0]["采购签核"].ToString();

                if (dtx.Rows[0]["纸品签核状态"].ToString() == "已签核")
                {
                    button1.Text = "已接收";
                    label25.Text = dtx.Rows[0]["纸品"].ToString();
                }
       
                if (dtx.Rows[0]["亚克力签核状态"].ToString() == "已签核")
                {
                    button2.Text = "已接收";
                    label26.Text = dtx.Rows[0]["亚克力"].ToString();
                }
         
                if (dtx.Rows[0]["木铁签核状态"].ToString() == "已签核")
                {
                    button3.Text = "已接收";
                    label27.Text = dtx.Rows[0]["木铁"].ToString();
                }
          
                if (dtx.Rows[0]["采购签核状态"].ToString() == "已签核")
                {
                    button4.Text = "已接收";
                    label28.Text = dtx.Rows[0]["采购签核"].ToString();
                
                }
             
           
            }
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            dtx = bc.getdt("SELECT * FROM PROJECT_INFO");
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx.Rows)
                {

                    string suggestWord = dr["PROJECT_ID"].ToString();
                    inputInfoSource.Add(suggestWord);
                }
            }
            this.comboBox1 .AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox1 .AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.comboBox1 .AutoCompleteCustomSource = inputInfoSource;
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
                if (bc.exists(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'") && EDIT != "有权限")
                {
                    hint.Text = "本账号无修改权限！";
                }
                else if (juage())
                {

                    IFExecution_SUCCESS = false;
                }

                else if ( checkBox7.Checked == false && checkBox8.Checked == false && checkBox9.Checked == false &&
                        checkBox10.Checked == false)
                {
                    IFExecution_SUCCESS = false;
                    hint.Text = string.Format("至少要选择一种签核人");
                }
                else
                {
                   
                    save();
                    if (IFExecution_SUCCESS == true )
                    {
                        bind();
                        F1.bind();
                    }
                    else if (IFExecution_SUCCESS == false)
                    {
                        hint.Text = cSAMPLE_RELY_LIST.ErrowInfo;
                    }
                  
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);


            }
        }
        private void add()
        {
            ClearText();
            IDO = cSAMPLE_RELY_LIST.GETID();
            IFExecution_SUCCESS = false;
            bind();
            ADD_OR_UPDATE = "ADD";
            clear_checkboxlist();
         
          
        }
        private void clear_checkboxlist()
        {
            for (i = 0; i < checkedListBox1.Items  .Count ; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
            for (i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, false);
            }
            for (i = 0; i < checkedListBox3.Items.Count; i++)
            {
                checkedListBox3.SetItemChecked(i, false);
            }
            for (i = 0; i < checkedListBox4.Items.Count; i++)
            {
                checkedListBox4.SetItemChecked(i, false);
            }
            for (i = 0; i < checkedListBox5.Items.Count; i++)
            {
                checkedListBox5.SetItemChecked(i, false);
            }
        }
        private void save()
        {
            btnSave.Focus();
            cSAMPLE_RELY_LIST.EMID = LOGIN.EMID;
            dtx = basec.getdts(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (dtx.Rows.Count > 0)
            {
                if (dtx.Rows[0]["项目状态"].ToString() == "已完成")
                {
                  
                    cSAMPLE_RELY_LIST.PAPER_MAKERID = "";
                    cSAMPLE_RELY_LIST.ACRYLIC_MAKERID = "";
                    cSAMPLE_RELY_LIST.WOOD_IRON_MAKERID = "";
                    cSAMPLE_RELY_LIST.PURCHASE_AUDIT_MAKERID = "";
                    dtx = bc.getdt("SELECT * FROM SAMPLE_RELY_LIST WHERE SRID='" + IDO + "'");
                    if (dtx.Rows.Count > 0)
                    {
                     
                        cSAMPLE_RELY_LIST.PAPER_AUDIT_STATUS = "";
                        cSAMPLE_RELY_LIST.ACRYLIC_AUDIT_STATUS = "";
                        cSAMPLE_RELY_LIST.WOOD_IRON_AUDIT_STATUS = "";
                        cSAMPLE_RELY_LIST.PURCHASE_AUDIT_STATUS = "";
                    }
                    cSAMPLE_RELY_LIST.CHARGE_AUDIT_STATUS = "N";
                    IDO = cSAMPLE_RELY_LIST.GETID();
                    cSAMPLE_RELY_LIST.SAMPLE_ID = cSAMPLE_RELY_LIST.GETID_SAMPLE_ID(comboBox1.Text);
                    checkBox7.Checked = false;
                    checkBox8.Checked = false;
                    checkBox9.Checked = false;
                    checkBox10.Checked = false;
                    button1.Text = "待接收确认";
                    button2.Text = "待接收确认";
                    button3.Text = "待接收确认";
                    button4.Text = "待接收确认";
                    radioButton4.Checked = false;
                    radioButton5.Checked = false;
                    radioButton6.Checked = false;
                    radioButton7.Checked = false;
                    checkBox3.Checked = false;
                    checkBox4.Checked = false;
                    checkBox5.Checked = false;
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                }
                else
                {
                    cSAMPLE_RELY_LIST.SAMPLE_ID = textBox2.Text;
                    cSAMPLE_RELY_LIST.PAPER_MAKERID = bc.RETURN_EMID(dtx.Rows[0]["纸品工号"].ToString());
                    cSAMPLE_RELY_LIST.ACRYLIC_MAKERID = bc.RETURN_EMID(dtx.Rows[0]["亚克力工号"].ToString());
                    cSAMPLE_RELY_LIST.WOOD_IRON_MAKERID = bc.RETURN_EMID(dtx.Rows[0]["木铁工号"].ToString());
                    cSAMPLE_RELY_LIST.PURCHASE_AUDIT_MAKERID = bc.RETURN_EMID(dtx.Rows[0]["采购签核工号"].ToString());
                    dtx = bc.getdt("SELECT * FROM SAMPLE_RELY_LIST WHERE SRID='" + IDO + "'");
    
                    if (dtx.Rows.Count > 0)
                    {
                    
                        cSAMPLE_RELY_LIST.PAPER_AUDIT_STATUS = dtx.Rows[0]["PAPER_AUDIT_STATUS"].ToString();
                        cSAMPLE_RELY_LIST.ACRYLIC_AUDIT_STATUS = dtx.Rows[0]["ACRYLIC_AUDIT_STATUS"].ToString();
                        cSAMPLE_RELY_LIST.WOOD_IRON_AUDIT_STATUS = dtx.Rows[0]["WOOD_IRON_AUDIT_STATUS"].ToString();
                        cSAMPLE_RELY_LIST.PURCHASE_AUDIT_STATUS = dtx.Rows[0]["PURCHASE_AUDIT_STATUS"].ToString();
                    }
                    cSAMPLE_RELY_LIST.CHARGE_AUDIT_STATUS = dtx.Rows[0]["CHARGE_AUDIT_STATUS"].ToString();
                }
                //MessageBox.Show(bc.RETURN_EMID(dtx.Rows[0]["项目工号"].ToString()) + "," + dtx.Rows[0]["项目签核状态"].ToString()+","+IDO );
          
            }
            else
            {
       
                cSAMPLE_RELY_LIST.PAPER_MAKERID = "";
                cSAMPLE_RELY_LIST.ACRYLIC_MAKERID = "";
                cSAMPLE_RELY_LIST.WOOD_IRON_MAKERID = "";
                cSAMPLE_RELY_LIST.PURCHASE_AUDIT_MAKERID = "";
                cSAMPLE_RELY_LIST.PAPER_AUDIT_STATUS = "N";
                cSAMPLE_RELY_LIST.ACRYLIC_AUDIT_STATUS = "N";
                cSAMPLE_RELY_LIST.WOOD_IRON_AUDIT_STATUS = "N";
                cSAMPLE_RELY_LIST.PURCHASE_AUDIT_STATUS = "N";
                cSAMPLE_RELY_LIST.SAMPLE_ID = cSAMPLE_RELY_LIST.GETID_SAMPLE_ID(comboBox1.Text);
                cSAMPLE_RELY_LIST.CHARGE_AUDIT_STATUS = "N";
            }
        
            cSAMPLE_RELY_LIST.SRID = IDO;
            cSAMPLE_RELY_LIST.NEED_COUNT = textBox3.Text;
            cSAMPLE_RELY_LIST.NEED_DATE = dateTimePicker1.Text;
            cSAMPLE_RELY_LIST.GROUP_TYPE = comboBox2.Text;
            cSAMPLE_RELY_LIST.ORDER_DATE = dateTimePicker2.Text;
        
                cSAMPLE_RELY_LIST.QUALITY_LEVAL = "";
            

            if (checkBox1.Checked && checkBox2.Checked == false)
            {
                cSAMPLE_RELY_LIST.OWN_OR_PURCHASE = "OWN";
                cSAMPLE_RELY_LIST.OWN_MAKERID = bc.RETURN_EMID(label12.Text);
                cSAMPLE_RELY_LIST.PURCHASE_MAKERID = "";

            }
            else if (checkBox1.Checked == false && checkBox2.Checked)
            {
                cSAMPLE_RELY_LIST.OWN_OR_PURCHASE = "PURCHASE";
                cSAMPLE_RELY_LIST.OWN_MAKERID = "";
                cSAMPLE_RELY_LIST.PURCHASE_MAKERID = bc.RETURN_EMID(label34.Text);

            }
            else if (checkBox1.Checked && checkBox2.Checked)
            {
                cSAMPLE_RELY_LIST.OWN_OR_PURCHASE = "OWN_AND_PURCHASE";
                cSAMPLE_RELY_LIST.OWN_MAKERID = bc.RETURN_EMID(label12.Text);
                cSAMPLE_RELY_LIST.PURCHASE_MAKERID = bc.RETURN_EMID(label34.Text);

            }
            else
            {
                cSAMPLE_RELY_LIST.OWN_OR_PURCHASE = "NOT ALL";
                cSAMPLE_RELY_LIST.OWN_MAKERID = "";
                cSAMPLE_RELY_LIST.PURCHASE_MAKERID = "";
            }
            cSAMPLE_RELY_LIST.OWN_REMARK = textBox4.Text;
            cSAMPLE_RELY_LIST.OTHER = textBox6.Text;
            if (radioButton4.Checked)
            {
                cSAMPLE_RELY_LIST.DISPLAY_TYPE = "PAPER";
            }
            else if (radioButton5.Checked)
            {
                cSAMPLE_RELY_LIST.DISPLAY_TYPE = "METAL";
            }
            else if (radioButton6.Checked)
            {
                cSAMPLE_RELY_LIST.DISPLAY_TYPE = "WOOD";
            }
            else if (radioButton7.Checked)
            {
               
                cSAMPLE_RELY_LIST.DISPLAY_TYPE = "PLASTIC";
            }
            else
            {
                cSAMPLE_RELY_LIST.DISPLAY_TYPE = "";
            }
            if (radioButton4.Checked)
            {
                if (checkBox5.Checked)
                {
                    cSAMPLE_RELY_LIST.DISPLAY_VALUE = textBox7.Text;
                }
                else
                {
                    cSAMPLE_RELY_LIST.DISPLAY_VALUE = "";
                }
            }
            else if (radioButton5.Checked)
            {
                cSAMPLE_RELY_LIST.DISPLAY_VALUE = textBox8.Text;
            }
            else if (radioButton6.Checked)
            {
                cSAMPLE_RELY_LIST.DISPLAY_VALUE = textBox9.Text;

            }
            else if (radioButton7.Checked)
            {
                cSAMPLE_RELY_LIST.DISPLAY_VALUE = textBox10.Text;
            }
            else
            {
                cSAMPLE_RELY_LIST.DISPLAY_VALUE = "";
            }
            if (radioButton4.Checked)
            {
                if (checkBox3.Checked)
                {
                    cSAMPLE_RELY_LIST.SMAL_POP = "Y";
                }
                else
                {
                    cSAMPLE_RELY_LIST.SMAL_POP = "N";
                }
                if (checkBox4.Checked)
                {
                    cSAMPLE_RELY_LIST.DISPLAY_FRAME = "Y";
                }
                else
                {
                    cSAMPLE_RELY_LIST.DISPLAY_FRAME = "N";
                }
                if (checkBox5.Checked)
                {
                    cSAMPLE_RELY_LIST.ALONE_DEPOSIT = "Y";
                }
                else
                {
                    cSAMPLE_RELY_LIST.ALONE_DEPOSIT = "N";
                }
            }
            else
            {
                cSAMPLE_RELY_LIST.SMAL_POP = "";
                cSAMPLE_RELY_LIST.DISPLAY_FRAME = "";
                cSAMPLE_RELY_LIST.ALONE_DEPOSIT = "";
            }
            cSAMPLE_RELY_LIST.IF_PROJECT_AUDIT = "Y";
            if (checkBox7.Checked)
            {
                cSAMPLE_RELY_LIST.IF_PAPER_AUDIT = "Y";
         
            }
            else
            {
                cSAMPLE_RELY_LIST.IF_PAPER_AUDIT = "N";
           
            }
            if (checkBox8.Checked)
            {
                cSAMPLE_RELY_LIST.IF_ACRYLIC_AUDIT = "Y";
            
            }
            else
            {
                cSAMPLE_RELY_LIST.IF_ACRYLIC_AUDIT = "N";
              
            }
            if (checkBox9.Checked)
            {
                cSAMPLE_RELY_LIST.IF_WOOD_IRON_AUDIT = "Y";
           
            }
            else
            {
                cSAMPLE_RELY_LIST.IF_WOOD_IRON_AUDIT = "N";
             
            }
            if (checkBox10.Checked)
            {
                cSAMPLE_RELY_LIST.IF_PURCHASE_AUDIT = "Y";
         
            }
            else
            {
                cSAMPLE_RELY_LIST.IF_PURCHASE_AUDIT = "N";
          
            }
         
            cSAMPLE_RELY_LIST.save(checkedListBox1 ,checkedListBox2 ,checkedListBox3 ,checkedListBox4 ,checkedListBox5 );
            IFExecution_SUCCESS = cSAMPLE_RELY_LIST.IFExecution_SUCCESS;
            hint.Text = cSAMPLE_RELY_LIST.ErrowInfo;
            try
            {
      
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region juage
        private bool juage()
        {
         
           bool b = false;
           if (IDO == null)
           {
               hint.Text = "编号不能为空";
               b = true;
           }
           else if (comboBox1.Text == "")
           {
               hint.Text = "项目号不能为空";
               b = true;
           }
           else if (!bc.exists(string.Format("SELECT * FROM PROJECT_INFO WHERE PROJECT_ID='{0}'",comboBox1.Text )))
           {
               b = true;
               hint.Text = string.Format("项目号不存在系统中");

           }
           else if (textBox3.Text== "")
           {
               hint.Text = "数量不能为空";
               b = true;
           }
           else if (bc.yesno (textBox3 .Text )==0)
           {
               hint.Text = "数量只能输入数字";
               b = true;
           }
           else if (juage2())
           {

               b = true;
           }
          else if (juage5())
          {

           b = true;
           }
         
            /*else if (bc.exists (string.Format ("SELECT * FROM WORKORDER_MST WHERE SRID='{0}'",bc.RETURN_SRID(textBox2 .Text ))))
            {
                hint.Text = string.Format("尺寸 {0} 已经在工单中使用不允许修改", textBox2 .Text );
                b = true;
            }*/
            return b;
        }
        #endregion
        #region juage2()

        private bool juage2()
        {
            bool b = false;
            DataTable dtx = dt;
            if (dtx.Rows.Count > 0)
            {
                if (checkBox1.Checked && label12 .Text  == "")
                {
                    b = true;
                    hint.Text = string.Format("选择自购时需输入自购工号");

                }
                else  if (label12 .Text  != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'",label12 .Text  )))
                {

                    b = true;
                    hint.Text = string.Format("自采工号不存在系统中");

                }
                else if (checkBox2.Checked && label34.Text  == "")
                {
                    b = true;
                    hint.Text = string.Format("选择采购时需输入采购工号");

                }
                else if (label34.Text != "" && !bc.exists(string.Format("SELECT * FROM EMPLOYEEINFO WHERE EMPLOYEE_ID='{0}'", label34.Text)))
                {

                    b = true;
                    hint.Text = string.Format("采购工号不存在系统中");

                }
                else if (textBox4.Text.Length > 300)
                {
                    b = true;
                    hint.Text = string.Format("只能输入最多300个汉字");
                }
       
         
                else if (textBox6.Text.Length > 400)
                {
                    b = true;
                    hint.Text = string.Format("只能输入最多400个汉字");
                }
           
             
            }

            return b;
        }
        #endregion
        #region juage2()

        private bool juage5()
        {
            bool b = false;
            if (cSAMPLE_RELY_LIST.JUAGE_IF_AUDIT_END(IDO))
            {

            }
            else  if (radioButton4.Checked == false && radioButton5.Checked == false && radioButton6.Checked == false && radioButton7.Checked == false)
                {

                    b = true;
                    hint.Text = string.Format("至少要选择一种陈列类型");
                }
                else if (radioButton4.Checked && checkBox3.Checked == false && checkBox4.Checked == false && checkBox5.Checked == false)
                {

                    b = true;
                    hint.Text = string.Format("POP 陈列架 堆头至少要选择一种");
                }
                else if (radioButton4.Checked && checkBox5.Checked && textBox7 .Text =="")
                {

                    b = true;
                    hint.Text = string.Format("堆头值不能为空");
                }
                else if (textBox7.Text != "" && bc.yesno(textBox7.Text) == 0)
                {
                    b = true;
                    hint.Text = string.Format("纸品平方数只能输入数字");
                }
                else if (radioButton5.Checked && textBox8.Text == "")
                {

                    b = true;
                    hint.Text = string.Format("金属值不能为空");
                }
           
                else if (textBox8.Text != "" && bc.yesno(textBox8.Text) == 0)
                {

                    b = true;
                    hint.Text = string.Format("金属值只能输入数字");
                }
                else if (radioButton6.Checked && textBox9.Text == "")
                {
                 
                    b = true;
                    hint.Text = string.Format("木器值不能为空");
                }
                else if (textBox9.Text != "" && bc.yesno(textBox9.Text) == 0)
                {
                    b = true;
                    hint.Text = string.Format("木器值只能输入数字");
                }
        
                else if (radioButton7.Checked && textBox10.Text == "")
                {

                    b = true;
                    hint.Text = string.Format("塑料值不能为空");
                }
             
                else if (textBox10.Text != "" && bc.yesno(textBox10.Text) == 0)
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
            dtx = basec.getdts(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            string v1 = "";
            if (dtx.Rows.Count > 0)
            {
                v1 = dtx.Rows[0]["项目状态"].ToString();
            }
            if (v1 == "已完成")
            {
                hint.Text = "此样板单号已审核通过不允许修改";
                b = true;
            }
    
            return b;
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (this.ActiveControl.TabIndex == 19 || this.ActiveControl.TabIndex == 20)
            {

            }
            else
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
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
   
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                add();
            }
            catch (Exception)
            {
                MessageBox.Show("网络连接中断");
            }
        }
  
        private void btnSearch_Click(object sender, EventArgs e)
        {
            bind();
            
        }
  
        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            DataTable dtx = new DataTable();
            dtx = bc.getdt("SELECT * FROM PROJECT_INFO");
            if (dtx.Rows.Count > 0)
            {
                comboBox1.Items.Clear();
                comboBox1.Items.Add("");
                foreach (DataRow dr in dtx.Rows)
                {
                    comboBox1.Items.Add(dr["PROJECT_ID"].ToString());
                }

            }
           
        }

        private void comboBox2_DropDown(object sender, EventArgs e)
        {

            dtx = bc.getdt("SELECT * FROM DEPART WHERE DEPART LIKE '%项目%'");
            if (dtx.Rows.Count > 0)
            {
                comboBox2.Items.Clear();
                foreach (DataRow dr in dtx.Rows)
                {
                    
                    comboBox2.Items.Add(dr["DEPART"].ToString());
                }

            }
        }

        private void comboBox3_DropDown(object sender, EventArgs e)
        {
            try
            {
                IF_DOUBLE_CLICK = false;
                BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
                FRM.GROUP = comboBox2.Text;
                FRM.SAMPLE_REAL_LIST_1920();
                FRM.ShowDialog();
                this.comboBox3.IntegralHeight = false;//使组合框不调整大小以显示其所有项
                this.comboBox3.DroppedDown = false;//使组合框不显示其下拉部分
                this.comboBox3.IntegralHeight = true;//恢复默认值
                if (IF_DOUBLE_CLICK)
                {

                    comboBox3.Text = ENAME;
                    label12.Text = EMPLOYEE_ID;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            
                textBox7.ReadOnly = false;
                textBox8.ReadOnly = true;
                textBox9.ReadOnly = true;
                textBox10.ReadOnly = true;
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox7.Focus();
            
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
           
                textBox8.ReadOnly = false;
                textBox7.ReadOnly = true;
                textBox9.ReadOnly = true;
                textBox10.ReadOnly = true;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                textBox7.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox8.Focus();
                textBox11.Text = RETURN_SAMPLE_PRICE();
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
                textBox9.ReadOnly = false;
                textBox7.ReadOnly = true;
                textBox8.ReadOnly = true;
                textBox10.ReadOnly = true;
                textBox11.Text = RETURN_SAMPLE_PRICE();
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                textBox9.Focus();
                textBox7.Text = "";
                textBox8.Text = "";
                textBox10.Text = "";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
                textBox10.ReadOnly = false;
                textBox7.ReadOnly = true;
                textBox8.ReadOnly = true;
                textBox9.ReadOnly = true;
                textBox11.Text = RETURN_SAMPLE_PRICE();
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Focus();
        }

        private void btnupload_Click(object sender, EventArgs e)
        {
           
     
            try
            {

                DataTable dty = bc.getdt("SELECT * FROM WAREFILE WHERE WAREID='" + IDO  + "'");
                if (juage())
                {

                }
                else if (dty.Rows.Count == 6)
                {

                    hint.Text = "最多只能上传三张图片";
                }
                else
                {
                    uploadfile();
                }
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
            if (bc.RETURN_SERVER_IP_OR_DOMAIN() == "")
            {
                hint.Text = "未设置服务器IP或域名";
            }

            else
            {
                OpenFileDialog openf = new OpenFileDialog();
                if (openf.ShowDialog() == DialogResult.OK)
                {
                    Random ro = new Random();
                    string stro = ro.Next(80, 10000000).ToString() + "-";
                    string NewName = DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString() + stro;

                    cfileinfo.SERVER_IP_OR_DOMAIN = bc.RETURN_SERVER_IP_OR_DOMAIN();
                    WATER_MARK_CONTENT = "";//水印内容
                    //cfileinfo.UploadImage(openf.FileName, Path.GetFileName(openf.FileName), textBox1 .Text );
                    //this.UploadFile(openf.FileName, System.IO.Path.GetFileName(openf.FileName), "File/", textBox1.Text);

                    string v21 = bc.FROM_RIGHT_UNTIL_CHAR(Path.GetFileName(openf.FileName), 46);
                    OLD_FILE_NAME = Path.GetFileName(openf.FileName);
                    NEW_FILE_NAME = NewName + Path.GetFileName(openf.FileName);
                    //如果上传的是图片文件
                    if (v21 == "jpeg" || v21 == "jpg" || v21 == "JPG" || v21 == "png" || v21 == "bmp" || v21 == "gif")
                    {
                        //裁切小图
                        cfileinfo.MakeThumbnail(openf.FileName, "d:\\" + Path.GetFileName(openf.FileName), 80, 80, "Cut");
                        //小图加水印
                        cfileinfo.ADD_WATER_MARK("d:\\" + Path.GetFileName(openf.FileName), "d:\\80X80" + NewName + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        //原图加水印
                        cfileinfo.ADD_WATER_MARK(openf.FileName, "d:\\INITIAL" + NewName + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        INITIAL_OR_OTHER = "INITIAL";
                        label5.Text = "";
                        //上传原图
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", "D:\\INITIAL" + NewName + System.IO.Path.GetFileName(openf.FileName),
                                "INITIAL" + NewName + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);

                        //上传80X80的缩略图
                        INITIAL_OR_OTHER = "80X80";
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", "D:\\80X80" + NewName + System.IO.Path.GetFileName(openf.FileName),
                                "80X80" + NewName + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);


                        //删除本地临时水印图及剪切图
                        if (File.Exists("d:\\80X80" + NewName + Path.GetFileName(openf.FileName)))
                        {
                            File.Delete("d:\\80X80" + NewName + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\" + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\INITIAL" + NewName + Path.GetFileName(openf.FileName));
                        }
                    }
                    else
                    {
                        label53.Visible = true;
                        label55.Visible = true;
                        label56.Visible = true;
                        label57.Visible = true;
                        progressBar1.Visible = true;
                        //上传的是非图片文件
                        INITIAL_OR_OTHER = "INITIAL";
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", openf.FileName,
                                                      "INITIAL" + NewName + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);
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
            }

        }
        #endregion
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
                uriString = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/uploadfile/" + newFileName;


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
            return returnValue;
        }
        #endregion
        #region bind2
        private void bind2()
        {

            dt3 = bc.getdt(@"
SELECT cast(0   as   bit)   as   复选框,
OLD_FILE_NAME AS 文件名,NEW_FILE_NAME AS 新文件名,FLKEY AS 索引,
PATH FROM WAREFILE WHERE WAREID='" + IDO + "'  AND INITIAL_OR_OTHER='80X80'");


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
                                bc.getcom("DELETE WAREFILE WHERE NEW_FILE_NAME='" + v4 + "'");

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
        #region dataGridView2_Click
        private void dataGridView2_Click(object sender, EventArgs e)
        {
          
          
            try
            {
                int i = dataGridView1.CurrentCell.ColumnIndex;

                if (dataGridView1.CurrentCell.ColumnIndex == 1)
                {
                    SaveFileDialog sfl = new SaveFileDialog();
                    sfl.FileName = dt3.Rows[dataGridView1.CurrentCell.RowIndex]["文件名"].ToString();
                    //sfl.Filter = "*.xls|*.doc|*.xlsx|*.docx";
                    if (sfl.ShowDialog() == DialogResult.OK)
                    {

                        WebClient wclient = new WebClient();
                        string v1 = bc.getOnlyString("SELECT PATH FROM WAREFILE WHERE FLKEY='" + dt3.Rows[dataGridView1.CurrentCell.RowIndex]["索引"].ToString() + "'");
                        wclient.DownloadFile(v1, sfl.FileName);

                        /*DataTable dt3x = bc.getdt("SELECT * FROM WAREFILE WHERE FLKEY='" + dt3.Rows[dataGridView2.CurrentCell.RowIndex]["索引"].ToString() + "'");
                        Byte[] byte2 = (byte[])dt3x.Rows[0]["IMAGE_DATA"];
                        System.IO.File.WriteAllBytes(sfl.FileName, byte2);*/
                        hint.Text = "已下载";
                    }
                }
                else if (i == 2)
                {
                    string path = bc.getOnlyString("SELECT PATH FROM WAREFILE WHERE FLKEY='" + dt3.Rows[dataGridView1.CurrentCell.RowIndex]["索引"].ToString() + "'");

                    string v21 = bc.FROM_RIGHT_UNTIL_CHAR(Path.GetFileName(path), 46);
                    if (v21 == "jpeg" || v21 == "jpg" || v21 == "png" || v21 == "bmp" || v21 == "gif")
                    {
                        //pictureBox1.Image = Image.FromStream(System.Net.WebRequest.Create(path).GetResponse().GetResponseStream());
                        //pictureBox1.Visible = true;
                      
                        //show_image.IMAGE_PATH = path;
                        //show_image.Show();

                    }
                    else
                    {
                        

                    }

                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = RETURN_SAMPLE_PRICE();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = RETURN_SAMPLE_PRICE();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
           
            textBox11.Text = RETURN_SAMPLE_PRICE();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = RETURN_SAMPLE_PRICE();
        }
        #region RETURN_SAMPLE_PRICE()
        private string  RETURN_SAMPLE_PRICE()
        {
            hint.Text = "";
            string v1 = "";
            if (juage())
            {

            }
            else if (cSAMPLE_RELY_LIST.JUAGE_IF_AUDIT_END(IDO))
            {

            }
            else if (radioButton4.Checked)
            {
                dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "纸"));
                decimal d1 = 0, d2 = 0, d3 = 0;
                if (dtx.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                    {
                        d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                    {
                        d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                    {
                        d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                    }

                }
                if (checkBox3.Checked && checkBox4.Checked == false && checkBox5.Checked == false)
                {

                    if (d1 != 0)
                    {
                        v1 = d1.ToString("0.00");
                    }
                }
                else if (checkBox3.Checked == false && checkBox4.Checked == true && checkBox5.Checked == false)
                {
                    if (d2 != 0)
                    {
                        v1 = d2.ToString("0.00");
                    }
                }
                else if (checkBox3.Checked == false && checkBox4.Checked == false && checkBox5.Checked == true)
                {
                    d3 = d3 * decimal.Parse(textBox7.Text);
                    if (d3 != 0)
                    {
                        v1 = d3.ToString("0.00");
                    }
                }
                else if (checkBox3.Checked && checkBox4.Checked && checkBox5.Checked == false)
                {
                    d3 = d1 + d2;
                    if (d3 != 0)
                    {
                        v1 = d3.ToString("0.00");
                    }
                }
                else if (checkBox3.Checked && checkBox4.Checked == false && checkBox5.Checked)
                {
                    d3 = d1 + d3 * decimal.Parse(textBox7.Text);
                    if (d3 != 0)
                    {
                        v1 = d3.ToString("0.00");
                    }
                }
                else if (checkBox3.Checked == false && checkBox4.Checked && checkBox5.Checked)
                {
                    d3 = d2 + d3 * decimal.Parse(textBox7.Text);
                    if (d3 != 0)
                    {
                        v1 = d3.ToString("0.00");
                    }
                }
                else if (checkBox3.Checked && checkBox4.Checked && checkBox5.Checked)
                {
                    d3 = d1 + d2 + d3 * decimal.Parse(textBox7.Text);
                    if (d3 != 0)
                    {
                        v1 = d3.ToString("0.00");
                    }
                }
            }
            else if (radioButton5.Checked)
            {
                dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "金属"));
                decimal d1 = 0, d2 = 0, d3 = 0;
                if (dtx.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                    {
                        d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                    {
                        d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                    {
                        d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                    }

                }

                d2 = d2 * decimal.Parse(textBox8.Text);
                if (d2 <= d1)
                {
                    d2 = d1;
                }
                else if (d2 >= d3)
                {
                    d2 = d3;
                }
                if (d2 != 0)
                {
                    v1 = d2.ToString("0.00");
                }
            }
            else if (radioButton6.Checked)
            {
                dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "木"));
                decimal d1 = 0, d2 = 0, d3 = 0;
                if (dtx.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                    {
                        d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                    {
                        d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                    {
                        d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                    }

                }

                d2 = d2 * decimal.Parse(textBox9.Text);
                if (d2 <= d1)
                {
                    d2 = d1;
                }
                else if (d2 >= d3)
                {
                    d2 = d3;
                }
                if (d2 != 0)
                {
                    v1 = d2.ToString("0.00");
                }
            }
            else if (radioButton7.Checked)
            {
                dtx = bc.getdt(cmaterial_price.sql + string.Format(" WHERE A.MATERIAL_TYPE='{0}'", "亚克力"));
                decimal d1 = 0, d2 = 0, d3 = 0;
                if (dtx.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["起步价"].ToString()))
                    {
                        d1 = decimal.Parse(dtx.Rows[0]["起步价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["单位计价"].ToString()))
                    {
                        d2 = decimal.Parse(dtx.Rows[0]["单位计价"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[0]["封顶金额"].ToString()))
                    {
                        d3 = decimal.Parse(dtx.Rows[0]["封顶金额"].ToString());
                    }

                }

                d2 = d2 * decimal.Parse(textBox10.Text);
                if (d2 <= d1)
                {
                    d2 = d1;
                }
                else if (d2 >= d3)
                {
                    d2 = d3;
                }
                if (d2 != 0)
                {
                    v1 = d2.ToString("0.00");
                }
            }
            return v1;

        }
        #endregion
        #region  ACTIVE_DISPALY_INF
        private void  ACTIVE_DISPALY_INFO()
        {
            right();
            if (cSAMPLE_RELY_LIST.JUAGE_IF_AUDIT_END(IDO))
            {

            }
            else
            {

                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
                radioButton6.Enabled = true;
                radioButton7.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox5.Enabled = true;
                textBox7.Enabled = true;
                textBox8.Enabled = true;
                textBox9.Enabled = true;
                textBox10.Enabled = true;
                string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
                if (bc.exists("SELECT * FROM REMIND WHERE RIID='" + IDO + "'"))
                {

                }
                else
                {
                    AE_MAKERID_ONE = bc.getOnlyString("SELECT AE_MAKERID_1 FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'");
                    INITIAL_MAKERID = bc.getOnlyString("SELECT MAKERID FROM SAMPLE_RELY_LIST WHERE SRID='" + IDO + "'");
                    if (cSAMPLE_RELY_LIST.JUAGE_IF_AUDIT_END(IDO))
                    {

                    }
                    else
                    {
                        basec.getcoms(@"
INSERT INTO REMIND
(
RIID,
NOTICE_MAKERID,
RECEIVE_STATUS,
DATE
) 
VALUES
('" + IDO + "','" + AE_MAKERID_ONE + "','N','" + varDate + "')");
                        basec.getcoms(@"
INSERT INTO REMIND
(
RIID,
NOTICE_MAKERID,
RECEIVE_STATUS,
DATE
) 
VALUES
('" + IDO + "','" + INITIAL_MAKERID + "','N','" + varDate + "')");
                    }


                }
            }
        }
        #endregion

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            textBox11.Text = RETURN_SAMPLE_PRICE();

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            textBox11.Text = RETURN_SAMPLE_PRICE();
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            textBox11.Text = RETURN_SAMPLE_PRICE();
            textBox7.Focus();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {

            }
            else
            {
                comboBox3.Text = "";
                label12 .Text ="";
            }
        }

        private void comboBox10_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
            FRM.GROUP = "采购";
            FRM.SAMPLE_REAL_LIST_1920();
            FRM.ShowDialog();
            this.comboBox10.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox10.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox10.IntegralHeight = true;//恢复默认值
            if (IF_DOUBLE_CLICK)
            {

                comboBox10.Text = ENAME;
                label34.Text = EMPLOYEE_ID;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_DropDown(object sender, EventArgs e)
        {
    
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
         
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST .sql  + " WHERE A.SRID='" + IDO  + "'");
            if (dtx.Rows.Count > 0)
            {
               
                if (label33.Text == "进行中")
                {
                  
                    basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET CHARGE_AUDIT_STATUS='Y' ,CHARGE_MAKERID='"+LOGIN.EMID +"' WHERE SRID='" + IDO + "'");
                    pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/audit.png"));
                    label33.Text = "已完成";
                    F1.bind();
                }
                else
                {
                    basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET CHARGE_AUDIT_STATUS='N',CHARGE_MAKERID='' WHERE SRID='" + IDO + "'");
                    pictureBox4.Image = Image.FromFile(System.IO.Path.GetFullPath("Image/61.png"));
                    label33.Text = "进行中";
                    F1.bind();
                }
                //bind();
            }
            else
            {
                hint.Text = "先保存单据才能做审核";
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
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            dtx = cSAMPLE_RELY_LIST.RETURN_DT(dtx);
            if (dtx.Rows.Count > 0)
            {

                cSAMPLE_RELY_LIST.ExcelPrint(dtx, "xxx样板依赖单", System.IO.Path.GetFullPath("xxx样板依赖单.xls"));
            }
            else
            {
                hint.Text = "先保存单据才能导出";
            }
         
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    basec.getcoms("DELETE SAMPLE_RELY_LIST WHERE SRID='" + IDO + "'");
                    add();
                    F1.load();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string v1 = bc.getOnlyString("SELECT PROJECT_NAME FROM PROJECT_INFO WHERE PROJECT_ID='" + comboBox1.Text + "'");
                if (v1 != "")
                {
                    textBox1.Text = v1;
                    textBox3.Focus();
                    bind2();
                }
            }
            catch (Exception)
            {

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage6())
            {

            }
            else if (checkBox7.Checked == false)
            {
                hint.Text = "需先勾选按扭前的纸品复选框";
            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button1.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET PAPER_AUDIT_STATUS='Y',PAPER_MAKERID='" + LOGIN.EMID + "' WHERE SRID='" + IDO + "'");
                        label25.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button1.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET PAPER_AUDIT_STATUS='N',PAPER_MAKERID='' WHERE SRID='" + IDO + "'");
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
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage6())
            {

            }
            else if (checkBox8.Checked == false)
            {
                hint.Text = "需先勾选按扭前的亚克力复选框";
            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button2.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET ACRYLIC_AUDIT_STATUS='Y',ACRYLIC_MAKERID='" + LOGIN.EMID + "' WHERE SRID='" + IDO + "'");
                        label26.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button2.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET ACRYLIC_AUDIT_STATUS='N',ACRYLIC_MAKERID='' WHERE SRID='" + IDO + "'");
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
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage6())
            {

            }
            else if (checkBox9.Checked == false)
            {
                hint.Text = "需先勾选按扭前的木铁复选框";
            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button3.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET WOOD_IRON_AUDIT_STATUS='Y',WOOD_IRON_MAKERID='" + LOGIN.EMID + "' WHERE SRID='" + IDO + "'");
                        label27.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button3.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET WOOD_IRON_AUDIT_STATUS='N',WOOD_IRON_MAKERID='' WHERE SRID='" + IDO + "'");
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
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage6())
            {

            }
            else if (checkBox10.Checked == false)
            {
                hint.Text = "需先勾选按扭前的采购复选框";
            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    hint.Text = "";
                    if (button4.Text == "待接收确认")
                    {

                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET PURCHASE_AUDIT_STATUS='Y',PURCHASE_AUDIT_MAKERID='" + LOGIN.EMID + "' WHERE SRID='" + IDO + "'");
                        label28.Text = bc.RETURN_ENMAE_USE_EMID(LOGIN.EMID);
                        button4.Text = "已接收";
                        F1.bind();
                    }
                    else
                    {
                        basec.getcoms(@"UPDATE SAMPLE_RELY_LIST SET PURCHASE_AUDIT_STATUS='N',PURCHASE_AUDIT_MAKERID='' WHERE SRID='" + IDO + "'");
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

        private void button5_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage5())
            {

            }
            else  if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat("UPDATE SAMPLE_RELY_LIST SET ");
                    if (radioButton4.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PAPER',");
                    }
                    else if (radioButton5.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='METAL',");
         
                    }
                    else if (radioButton6.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='WOOD',");
                    }
                    else if (radioButton7.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PLASTIC',");
                    }
                    else
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='',");
                    }
         
                    if (checkBox3.Checked)
                    {
                        sqb.AppendFormat("SMAL_POP='{0}',", "Y");
                    }
                    else
                    {
                        sqb.AppendFormat("SMAL_POP='{0}',", "N");
                    }
                    if (checkBox4.Checked)
                    {
                        sqb.AppendFormat("DISPLAY_FRAME='{0}',", "Y");
                    }
                    else
                    {
                        sqb.AppendFormat("DISPLAY_FRAME='{0}',", "N");
                    }
                    if (checkBox5.Checked)
                    {
                        sqb.AppendFormat("ALONE_DEPOSIT='{0}',DISPLAY_VALUE='{1}'", "Y", textBox7.Text);
                    }
                    else
                    {
                        sqb.AppendFormat("ALONE_DEPOSIT='{0}',DISPLAY_VALUE='{1}'", "N", "");
                    }
                    sqb.AppendFormat(" WHERE SRID='{0}'",IDO );
                    basec.getcoms(sqb.ToString ());
                    IFExecution_SUCCESS = true;
                    hint_bind();
                    F1.bind();

                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage5())
            {

            }
            else if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat("UPDATE SAMPLE_RELY_LIST SET ");
                    if (radioButton4.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PAPER',");
                    }
                    else if (radioButton5.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='METAL',");

                    }
                    else if (radioButton6.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='WOOD',");
                    }
                    else if (radioButton7.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PLASTIC',");
                    }
                    else
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='',");
                    }
                    sqb.AppendFormat(" SMAL_POP='{0}',", "N");
                    sqb.AppendFormat(" DISPLAY_FRAME='{0}',", "N");
                    sqb.AppendFormat(" ALONE_DEPOSIT='{0}',", "N");
                    sqb.AppendFormat(" DISPLAY_VALUE='{0}'", textBox8.Text);
                    sqb.AppendFormat(" WHERE SRID='{0}'", IDO);
                    basec.getcoms(sqb.ToString());
                    IFExecution_SUCCESS = true;
                    hint_bind();
                    F1.bind();

                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage5())
            {

            }
            else if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat("UPDATE SAMPLE_RELY_LIST SET ");
                    if (radioButton4.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PAPER',");
                    }
                    else if (radioButton5.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='METAL',");

                    }
                    else if (radioButton6.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='WOOD',");
                    }
                    else if (radioButton7.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PLASTIC',");
                    }
                    else
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='',");
                    }
                    sqb.AppendFormat(" SMAL_POP='{0}',", "N");
                    sqb.AppendFormat(" DISPLAY_FRAME='{0}',", "N");
                    sqb.AppendFormat(" ALONE_DEPOSIT='{0}',", "N");
                    sqb.AppendFormat(" DISPLAY_VALUE='{0}'", textBox10.Text);
                    sqb.AppendFormat(" WHERE SRID='{0}'", IDO);
                    basec.getcoms(sqb.ToString());
                    IFExecution_SUCCESS = true;
                    hint_bind();
                    F1.bind();

                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt(cSAMPLE_RELY_LIST.sql + " WHERE A.SRID='" + IDO + "'");
            if (juage5())
            {

            }
            else if (juage6())
            {

            }
            else
            {
                if (dtx.Rows.Count > 0)
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat("UPDATE SAMPLE_RELY_LIST SET ");
                    if (radioButton4.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PAPER',");
                    }
                    else if (radioButton5.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='METAL',");

                    }
                    else if (radioButton6.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='WOOD',");
                    }
                    else if (radioButton7.Checked)
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='PLASTIC',");
                    }
                    else
                    {
                        sqb.AppendFormat(" DISPLAY_TYPE='',");
                    }
                    sqb.AppendFormat(" SMAL_POP='{0}',", "N");
                    sqb.AppendFormat(" DISPLAY_FRAME='{0}',", "N");
                    sqb.AppendFormat(" ALONE_DEPOSIT='{0}',", "N");
                    sqb.AppendFormat(" DISPLAY_VALUE='{0}'", textBox9.Text);
                    sqb.AppendFormat(" WHERE SRID='{0}'", IDO);
                    basec.getcoms(sqb.ToString());
                    IFExecution_SUCCESS = true;
                    hint_bind();
                    F1.bind();

                }
                else
                {
                    hint.Text = "先保存单据才能做审核";
                }
            }

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
                        string v1 = bc.getOnlyString(sqb.ToString());
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

    }
}
