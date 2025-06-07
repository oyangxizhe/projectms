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
    public partial class SUN_SCREEN : Form
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
        public SUN_SCREEN()
        {
            InitializeComponent();
        }
        public SUN_SCREEN(OFFER_MANAGE .PROJECT_INFOT FRM)
        {
            F1 = FRM;
            InitializeComponent();
        }
        private void SUN_SCREEN_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            label3.Text = "项目号："+PROJECT_ID ;
            label4.Text = "项目名称：" + PROJECT_NAME ;
            label3.Font = new Font("", 20);
            label4.Font = new Font("", 20);
            label3.BackColor = CCOLOR.lylfnp;
            label3.ForeColor = Color.White;
            label4.BackColor = CCOLOR.lylfnp;
            label4.ForeColor = Color.White;
         
        }
  
        #region bind()
        private void Bind()
        {
    
           
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

   
        private void btnExit_Click(object sender, EventArgs e)
        {
            
            this.Close();
            F1.Close();
        }

        #region
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            OFFER_MANAGE.SAMPLE_RELY_LISTT FRM = new CSPSS.OFFER_MANAGE.SAMPLE_RELY_LISTT();
            FRM.IDO = csample_rely_list.GETID();
            FRM.PROJECT_ID = PROJECT_ID;
            FRM.PROJECT_NAME = PROJECT_NAME;
            FRM.Show();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            OFFER_MANAGE.PRINTING_OFFERT FRM = new CSPSS.OFFER_MANAGE.PRINTING_OFFERT();
            FRM.IDO = cprinting_offer.GETID();
            FRM.PROJECT_ID = PROJECT_ID;
            FRM.PROJECT_NAME = PROJECT_NAME;
            FRM.ADD_OR_UPDATE = "ADD";
            FRM.Show();
        }
    }
}
