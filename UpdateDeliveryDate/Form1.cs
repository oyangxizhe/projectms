using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using XizheC;
using System.IO;

namespace UpdateDeliveryDate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {

            timer1.Enabled = true;
          

        }
        private void bind()
        {
            string v1=@"
SELECT 
A.FLKEY AS FLKEY,
A.NEW_FILE_NAME AS NEW_FILE_NAME
FROM SERVER_DELETE_FILE A
";
            dt = bc.getdt(v1);
       
            try
            {
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string v2 = "d:/uploadfile/" + "INITIAL"+dr["NEW_FILE_NAME"].ToString();
                        if (File.Exists(v2))
                        {
                            File.Delete(v2);
                        }
                        string v3 = "d:/uploadfile/" + "80X80" + dr["NEW_FILE_NAME"].ToString();
                        if (File.Exists(v3))
                        {
                            File.Delete(v3);
                            bc.getcom("DELETE SERVER_DELETE_FILE WHERE NEW_FILE_NAME='" + dr["NEW_FILE_NAME"].ToString() + "'");
                        }

                    }

                }
            }
            catch (Exception)
            {

            }
            //dataGridView1.DataSource = bc.getdt(sqlo);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
           
            timer1.Interval = 1000;
            bind();
        }
  

   
    }
}
