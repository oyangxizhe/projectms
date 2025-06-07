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
    public partial class SHOW_IMAGE : Form
    {
        DataTable dt = new DataTable();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _PATH;
        public string PATH
        {
            set { _PATH = value; }
            get { return _PATH; }

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
        basec bc = new basec();
        protected int M_int_judge, i;
        protected int select;
        public SHOW_IMAGE()
        {
            InitializeComponent();
        }
        #region double_click
        private void dgvEmployeeInfo_DoubleClick(object sender, EventArgs e)
        {
            
        }
        #endregion
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SHOW_IMAGE_Load(object sender, EventArgs e)
        {
          this.Icon = Resource1.xz_200X200;
            pictureBox1.Image = Image.FromStream(System.Net.WebRequest.Create(PATH).GetResponse().GetResponseStream());
            pictureBox1.Visible = true;
        }

     
    
   
    }
}
