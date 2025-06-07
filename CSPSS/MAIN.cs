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
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace CSPSS
{
    public partial class MAIN : Form
    {
         DataTable dt = new DataTable();
         DataTable dt2 = new DataTable();
         basec bc = new basec();
         CUSER cuser = new CUSER();
         CEMPLOYEE_INFO cemplyee_info = new CEMPLOYEE_INFO();
         Color c2 = System.Drawing.ColorTranslator.FromHtml("#4a7bb8");
         Color c3 = System.Drawing.ColorTranslator.FromHtml("#24ade5");
         CDEPART cdepart = new CDEPART();
         CPOSITION cposition = new CPOSITION();
         CPRINT_OPTION cprint_option = new CPRINT_OPTION();
         CTISSUE_SPEC ctissue_sepc = new CTISSUE_SPEC();
         CPAPER_CORE cpaper_core = new CPAPER_CORE();
         CPRIMARY_COLORS cprimary_colors = new CPRIMARY_COLORS();
         CCOLOR_PARAMETERS ccolor_parameters = new CCOLOR_PARAMETERS();
         CSURFACE_PROCESSING csurface_processing = new CSURFACE_PROCESSING();
         CLAMINATING_PROCESS claminating_process = new CLAMINATING_PROCESS();
         CPRINTING_OFFER cprinting_offer = new CPRINTING_OFFER();
         CPRINTING_TYPE cprinting_type = new CPRINTING_TYPE();
         CMACHINING cmachining = new CMACHINING();
         CDOOR_PARAMETERS cdoor_parameters = new CDOOR_PARAMETERS();
         CCUSTOMER_INFO ccustomer_info = new CCUSTOMER_INFO();
         CPROJECT_INFO cproject_info = new CPROJECT_INFO();
         CPROCESSING_TECHNOLOGY cprocessing_technology = new CPROCESSING_TECHNOLOGY();
         CSAMPLE_RELY_LIST csample_rely_list = new CSAMPLE_RELY_LIST();
         CUNIT cunit = new CUNIT();
         CMATERIAL_PRICE cmaterial_price = new CMATERIAL_PRICE();
         CDIE_CUTTING_COST cdie_cutting_cost = new CDIE_CUTTING_COST();
         CPARTS_AUXILIARY cparts_auxiliary = new CPARTS_AUXILIARY();
         CPORTRAY cportray = new CPORTRAY();
         CPACK_MATERIAL cpack_material = new CPACK_MATERIAL();
         CARTIFICIAL cartificial = new CARTIFICIAL();
         CPURCHASE cpurchase = new CPURCHASE();
         CTRANSPORT ctransport = new CTRANSPORT();
         COTHER_COST cother_cost = new COTHER_COST();
         CPRINTING_MACHINE_SIZE cprinting_machine_size = new CPRINTING_MACHINE_SIZE();
         CUSER_GROUP cuser_group = new CUSER_GROUP();
         StringBuilder sqb = new StringBuilder();
         CNO_PAPER_OFFER cno_paper_offer = new CNO_PAPER_OFFER();
         CORDER_TYPE corder_type = new CORDER_TYPE();
         CPN_PRODUCTION_INSTRUCTIONS cpn_production_instruction = new CPN_PRODUCTION_INSTRUCTIONS();
         CNOTICE_LIST cnotice_list = new CNOTICE_LIST();
         CINVENTORY cinventory = new CINVENTORY();
         CINVOICE cinvoice = new CINVOICE();
         CRECEIVABLE creceivable = new CRECEIVABLE();
         BASE_INFO.AUDIT_LIST audit_list = new BASE_INFO.AUDIT_LIST();
         bool b = false;
         private string _SAMPLE_ID;
         public string SAMPLE_ID
         {
             set { _SAMPLE_ID = value; }
             get { return _SAMPLE_ID; }

         }
         private string _ID;
         public string ID
         {
             set { _ID = value; }
             get { return _ID; }
         }
         private string _PNID;
         public string PNID
         {
             set { _PNID = value; }
             get { return _PNID; }
         }
         private string _PROJECT_NAME;
         public string PROJECT_NAME
         {
             set { _PROJECT_NAME = value; }
             get { return _PROJECT_NAME; }

         }
         LOGIN F1 = new LOGIN();
         CFileInfo cfileinfo = new CFileInfo();
        public MAIN()
        {
            InitializeComponent();
        }
        public MAIN(LOGIN FRM)
        {
            F1 = FRM;
            InitializeComponent();
        }
        #region bind1
        private void bind1()
        {
            this.Icon = Resource1.xz_200X200;
            timer1.Enabled = true;
            timer1.Interval = 1000;
            pictureBox1.BackColor = c2;
            notifyIcon1.Icon = Resource1.xz_200X200;
            notifyIcon1.Text = "xxx项目管理系统";
            pictureBox1.Image = Resource1.project_ms;
            sqb = new StringBuilder();
            sqb.AppendFormat("xxx股份有限公司项目管理系统 ");
            sqb.AppendFormat("Version 1.0.0 ");
            sqb.AppendFormat("当前版本更新日期：{0}", F1.CURRENT_EDITION);//将LOGIN里的版本信息传到主窗口
            this.Text = sqb.ToString();
            dt = bc.getdt("SELECT * from RightList where USID = '" + LOGIN.USID + "'" );
            SHOW_TREEVIEW(dt);
            menuStrip1.Font = new Font("宋体", 9);
            this.WindowState = FormWindowState.Maximized;
            toolStripStatusLabel1.Text = "||当前用户：" + LOGIN.UNAME;
            toolStripStatusLabel2.Text = "||所属部门：" + LOGIN.DEPART;
            toolStripStatusLabel3.Text = "||登录时间：" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + " || 技术支持：苏州好用软件有限公司";
            listView1.BackColor = c2;
            listView1.ForeColor = Color.White;
            listView1.Font = new Font("新宋体", 11);
            listView2.BorderStyle = BorderStyle.None;
            imageList1.Images.Add(CSPSS.Resource1._1);
            imageList1.Images.Add(CSPSS.Resource1._2);
            imageList1.Images.Add(CSPSS.Resource1._3);
            imageList1.Images.Add(CSPSS.Resource1._4);
            imageList1.Images.Add(CSPSS.Resource1._5);
            imageList1.Images.Add(CSPSS.Resource1._6);
            imageList1.Images.Add(CSPSS.Resource1._7);
            imageList1.Images.Add(CSPSS.Resource1._8);
            imageList1.Images.Add(CSPSS.Resource1._9);
            imageList1.Images.Add(CSPSS.Resource1._10);
            imageList1.Images.Add(CSPSS.Resource1._11);
            imageList1.Images.Add(CSPSS.Resource1._12);
            imageList1.Images.Add(CSPSS.Resource1._13);
            imageList1.Images.Add(CSPSS.Resource1._14);
            imageList1.Images.Add(CSPSS.Resource1._15);
            imageList1.Images.Add(CSPSS.Resource1._16);
            imageList1.Images.Add(CSPSS.Resource1._17);
            imageList1.Images.Add(CSPSS.Resource1._18);
            imageList1.Images.Add(CSPSS.Resource1._19);
            imageList1.Images.Add(CSPSS.Resource1._20);
            imageList1.Images.Add(CSPSS.Resource1._21);
            imageList1.Images.Add(CSPSS.Resource1._22);
            imageList1.Images.Add(CSPSS.Resource1._23);
            imageList1.Images.Add(CSPSS.Resource1._24);
            imageList1.Images.Add(CSPSS.Resource1._25);
            imageList1.Images.Add(CSPSS.Resource1._26);
            imageList1.Images.Add(CSPSS.Resource1._27);
            imageList1.Images.Add(CSPSS.Resource1._28);
            imageList1.Images.Add(CSPSS.Resource1._29);
            imageList1.Images.Add(CSPSS.Resource1._30);
            imageList1.Images.Add(CSPSS.Resource1._31);
            imageList1.Images.Add(CSPSS.Resource1._32);
            imageList1.Images.Add(CSPSS.Resource1._33);
            imageList1.Images.Add(CSPSS.Resource1._34);
            imageList1.Images.Add(CSPSS.Resource1._35);
            imageList1.Images.Add(CSPSS.Resource1._36);
            imageList1.Images.Add(CSPSS.Resource1._37);
            imageList1.Images.Add(CSPSS.Resource1._38);
            imageList1.Images.Add(CSPSS.Resource1._39);
            imageList1.Images.Add(CSPSS.Resource1._40);
            imageList1.Images.Add(CSPSS.Resource1._41);
            imageList1.Images.Add(CSPSS.Resource1._42);
            imageList1.Images.Add(CSPSS.Resource1._43);
            imageList1.Images.Add(CSPSS.Resource1._44);
            imageList1.Images.Add(CSPSS.Resource1._45);
            imageList1.Images.Add(CSPSS.Resource1._46);
            imageList1.Images.Add(CSPSS.Resource1._47);
            imageList1.Images.Add(CSPSS.Resource1._48);
            imageList1.Images.Add(CSPSS.Resource1._49);
            imageList1.Images.Add(CSPSS.Resource1._50);
            imageList1.Images.Add(CSPSS.Resource1._51);
            imageList1.Images.Add(CSPSS.Resource1._52);
            imageList1.Images.Add(CSPSS.Resource1._53);
            imageList1.Images.Add(CSPSS.Resource1._54);
            imageList1.Images.Add(CSPSS.Resource1._55);
            imageList1.Images.Add(CSPSS.Resource1._56);
            imageList1.Images.Add(CSPSS.Resource1._57);

          

            imageList1.ColorDepth = ColorDepth.Depth32Bit;/*防止图片失真*/
            listView1.View = View.SmallIcon;
            listView2.View = View.LargeIcon;
            imageList1.ImageSize = new Size(48, 48);/*set imglist size*/
            listView1.SmallImageList = imageList1;
            listView2.LargeImageList = imageList1;
        }
        #endregion

        #region load
        private void MAIN_Load(object sender, EventArgs e)
        {
            try
            {
                bind1();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "NODEID ASC", "PARENT_NODEID=0");
            if (dt.Rows.Count > 0)
            {
                for(int i=0;i<dt.Rows.Count ;i++)
                {
                    ListViewItem lvi = listView1.Items.Add(dt.Rows[i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
                DataTable  dtx = bc.GET_DT_TO_DV_TO_DT(dt, "", "NODE_NAME='项目管理'");
                if (dtx.Rows.Count > 0)
                {
                    click(dtx.Rows[0]["NODE_NAME"].ToString());
                    if(listView1.Items.Count ==1)
                    {
                        listView1.Items[0].BackColor = c3;
                    }
                    else
                    {
                        listView1.Items[1].BackColor = c3;
                    }
                }
                else
                {
                    click(dt.Rows[0]["NODE_NAME"].ToString());
                    listView1.Items[0].BackColor = c3;
                }
            }
        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string NODEID)
        {
          
            dt2 = bc.getdt("SELECT * FROM RIGHTLIST WHERE PARENT_NODEID='" + NODEID  + "'AND  USID = '" + LOGIN.USID + "' ORDER BY NODEID ASC" );
            if (dt2.Rows.Count > 0)
            {
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    ListViewItem lvi = listView2.Items.Add(dt2.Rows[i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt2.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
            }
            else
            {
                

            }
        }
        #endregion

         private void 退出系统ToolStripMenuItem1_Click(object sender, EventArgs e)
         {
             if (MessageBox.Show("确定要退出本系统吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
             {
                 EXIT();
             }
             else
             {
                 
             }
         }
         private void listView1_Click(object sender, EventArgs e)
         {
             try
             {
                 string v1 = listView1.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/
                 click(v1);
             }
             catch (Exception)
             {


             }
            
         }
         private void click(string NODE_NAME)
         {
            
             listView2.Items.Clear();
             string id = bc.getOnlyString("SELECT NODEID FROM RIGHTLIST WHERE NODE_NAME='" + NODE_NAME + "'");
             SHOW_TREEVIEW_O(id);

             foreach (ListViewItem lvi in listView1.Items)
             {
                 if (lvi.Selected)
                 {
                     lvi.BackColor = c3;
                     pictureBox1.Focus();/*SELECTED AFTER MOVE FOCUS*/
                 }
                 else
                 {
                     lvi.BackColor = c2;
                 }

             }

         }
         #region listview2
         private void listView2_Click(object sender, EventArgs e)
         {
             string v1 = listView2.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/
             #region v1
            if (v1 == "员工信息维护")
             {
                 CSPSS.BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
                 FRM.IDO = cemplyee_info.GETID();
                 FRM.Show();
             }
            else if (v1 == "签核名单")
            {
                CSPSS.BASE_INFO.AUDIT_LIST FRM = new BASE_INFO.AUDIT_LIST();
                FRM.Show();
            }
            else if (v1 == "开票及收款汇总")
            {
                CSPSS.ORDER_MANAGE.INVOICE_AND_RECEIVABLE FRM = new ORDER_MANAGE.INVOICE_AND_RECEIVABLE();
                FRM.Show();
            }
             else if (v1 == "登录信息")
            {
                CSPSS.USER_MANAGE.LOGIN_INFO FRM = new USER_MANAGE.LOGIN_INFO();
                FRM.Show();
            }
            else if (v1 == "出货及开票汇总")
            {
                CSPSS.ORDER_MANAGE.SALES_AND_INVOICE FRM = new ORDER_MANAGE.SALES_AND_INVOICE();
                FRM.Show();
            }
            else if (v1 == "收款维护")
            {
                CSPSS.ORDER_MANAGE.RECEIVABLET FRM = new ORDER_MANAGE.RECEIVABLET();
                FRM.IDO = creceivable.GETID();
                FRM.Show();
            }
            else if (v1 == "收款信息汇总表")
            {
                CSPSS.ORDER_MANAGE.RECEIVABLE FRM = new ORDER_MANAGE.RECEIVABLE();;
                FRM.Show();
            }
            else if (v1 == "开票维护")
            {
                CSPSS.ORDER_MANAGE.INVOICET FRM = new ORDER_MANAGE.INVOICET();
                FRM.IDO = cinvoice.GETID();
                FRM.Show();
            }
            else if (v1 == "开票信息汇总表")
            {
                CSPSS.ORDER_MANAGE.INVOICE FRM = new ORDER_MANAGE.INVOICE();
                FRM.Show();
            }
            else if (v1 == "库存维护")
            {
                CSPSS.ORDER_MANAGE.INVENTORYT FRM = new ORDER_MANAGE.INVENTORYT();
                FRM.IDO = cinventory.GETID();
                FRM.Show();
            }
            else if (v1 == "库存信息汇总表")
            {
                CSPSS.ORDER_MANAGE.INVENTORY FRM = new ORDER_MANAGE.INVENTORY();
                FRM.Show();
            }
            else if (v1 == "通知名单")
            {
                CSPSS.BASE_INFO.NOTICE_LIST FRM = new BASE_INFO.NOTICE_LIST();
                FRM.IDO = cnotice_list.GETID();
                FRM.Show();
            }
            else if (v1 == "生产指示书")
            {
                CSPSS.ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST FRM = new ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST();
                FRM.IDO = cpn_production_instruction.GETID();
                FRM.Show();
            }
            else if (v1 == "订单信息查询")
            {
                CSPSS.ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONS FRM = new ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONS();
                FRM.Show();
            }
            else if (v1 == "订单类型")
            {
                BASE_INFO.ORDER_TYPE FRM = new BASE_INFO.ORDER_TYPE();
                FRM.IDO = corder_type.GETID();
                FRM.Show();
            }
            else if (v1 == "非纸品报价新增")
            {
                CSPSS.ORDER_MANAGE.NO_PAPER_OFFERT FRM = new ORDER_MANAGE.NO_PAPER_OFFERT();
                FRM.IDO = cno_paper_offer.GETID();
                FRM.Show();
            }
            else if (v1 == "非纸品报价查询")
            {
                CSPSS.ORDER_MANAGE.NO_PAPER_OFFER FRM = new ORDER_MANAGE.NO_PAPER_OFFER();
                FRM.Show();
            }
            else if (v1 == "其它费用")
            {
                CSPSS.BASE_INFO.OTHER_COST FRM = new CSPSS.BASE_INFO.OTHER_COST();
                FRM.IDO = cother_cost.GETID();
                FRM.Show();
            }
            else if (v1 == "运输")
            {
                CSPSS.BASE_INFO.TRANSPORT FRM = new CSPSS.BASE_INFO.TRANSPORT();
                FRM.IDO = ctransport.GETID();
                FRM.Show();
            }
            else if (v1 == "外购件")
            {
                CSPSS.BASE_INFO.PURCHASE FRM = new CSPSS.BASE_INFO.PURCHASE();
                FRM.IDO = cpurchase.GETID();
                FRM.Show();
            }
            else if (v1 == "人工费")
            {
                CSPSS.BASE_INFO.ARTIFICIAL FRM = new CSPSS.BASE_INFO.ARTIFICIAL();
                FRM.IDO = cartificial.GETID();
                FRM.Show();
            }
            else if (v1 == "包装材料")
            {
                CSPSS.BASE_INFO.PACK_MATERIAL FRM = new CSPSS.BASE_INFO.PACK_MATERIAL();
                FRM.IDO = cpack_material.GETID();
                FRM.Show();
            }
            else if (v1 == "画面写真")
            {
                CSPSS.BASE_INFO.PORTRAY FRM = new CSPSS.BASE_INFO.PORTRAY();
                FRM.IDO = cportray.GETID();
                FRM.Show();
            }
            else if (v1 == "配件辅材")
            {
                CSPSS.BASE_INFO.PARTS_AUXILIARY FRM = new CSPSS.BASE_INFO.PARTS_AUXILIARY();
                FRM.IDO = cparts_auxiliary.GETID();
                FRM.Show();
            }
            else if (v1 == "刀模费")
            {
                CSPSS.BASE_INFO.DIE_CUTTING_COST FRM = new CSPSS.BASE_INFO.DIE_CUTTING_COST();
                FRM.IDO = cdie_cutting_cost.GETID();
                FRM.Show();
            }
            else if (v1 == "单位")
            {
                CSPSS.BASE_INFO.UNIT FRM = new CSPSS.BASE_INFO.UNIT();
                FRM.IDO = cunit.GETID();
                FRM.Show();

            }
            else if (v1 == "类型计价")
            {
                CSPSS.BASE_INFO.MATERIAL_PRICE FRM = new CSPSS.BASE_INFO.MATERIAL_PRICE();
                FRM.IDO = cmaterial_price.GETID();
                FRM.Show();

            }
            else if (v1 == "文件存储IP")
            {
                CSPSS.BASE_INFO.UPLOADFILE_DOMAIN FRM = new CSPSS.BASE_INFO.UPLOADFILE_DOMAIN();
                FRM.Show();

            }
            else if (v1 == "加工工艺")
            {
                CSPSS.BASE_INFO.PROCESSING_TECHNOLOGY FRM = new CSPSS.BASE_INFO.PROCESSING_TECHNOLOGY();
                FRM.IDO = cprocessing_technology.GETID();
                FRM.Show();

            }
            else if (v1 == "打样单新增")
            {
                try
                {
                    OFFER_MANAGE.SAMPLE_RELY_LISTT FRM = new CSPSS.OFFER_MANAGE.SAMPLE_RELY_LISTT();
                    FRM.ADD_OR_UPDATE = "ADD";
                    FRM.IDO = csample_rely_list.GETID();
                    FRM.Show();
                }
                catch (Exception)
                {
                    MessageBox.Show("网络连接中断");
                }
            }
            else if (v1 == "打样单查询")
            {
                OFFER_MANAGE.SAMPLE_RELY_LIST FRM = new CSPSS.OFFER_MANAGE.SAMPLE_RELY_LIST();
                FRM.Show();

            }
            else if (v1 == "项目新增")
            {
                try
                {
                    OFFER_MANAGE.PROJECT_INFOT FRM = new CSPSS.OFFER_MANAGE.PROJECT_INFOT();
                    FRM.ADD_OR_UPDATE = "ADD";
                    FRM.IDO = cproject_info.GETID();
                    FRM.Show();
                }
                catch (Exception)
                {
                    MessageBox.Show("网络连接中断");
                }

            }
            else if (v1 == "项目查询")
            {
                try
                {
                    OFFER_MANAGE.PROJECT_INFO FRM = new CSPSS.OFFER_MANAGE.PROJECT_INFO();
                    FRM.Show();
                }
                catch (Exception)
                {
                    MessageBox.Show("网络连接中断");
                }

            }
            else if (v1 == "客户信息")
            {
                CSPSS.BASE_INFO.CUSTOMER_INFO FRM = new CSPSS.BASE_INFO.CUSTOMER_INFO();
                FRM.Show();

            }

            else if (v1 == "材料门幅参数")
            {
                CSPSS.BASE_INFO.DOOR_PARAMETERS FRM = new CSPSS.BASE_INFO.DOOR_PARAMETERS();
                FRM.IDO = cdoor_parameters.GETID();
                FRM.Show();

            }
            else if (v1 == "机加工")
            {
                CSPSS.BASE_INFO.MACHINING FRM = new CSPSS.BASE_INFO.MACHINING();
                FRM.IDO = cmachining.GETID();
                FRM.Show();

            }
            else if (v1 == "印刷类")
            {
                CSPSS.BASE_INFO.PRINTING_TYPE FRM = new CSPSS.BASE_INFO.PRINTING_TYPE();

                FRM.Show();

            }
            else if (v1 == "纸品报价新增")
            {
                try
                {
                    CSPSS.OFFER_MANAGE.PRINTING_OFFERT FRM = new CSPSS.OFFER_MANAGE.PRINTING_OFFERT();
                    FRM.ADD_OR_UPDATE = "ADD";
                    FRM.IDO = cprinting_offer.GETID();
                    FRM.Show();
                }
                catch (Exception)
                {
                    MessageBox.Show("网络连接中断");
                }
            }
            else if (v1 == "纸品报价查询")
            {
                try
                {
                    CSPSS.OFFER_MANAGE.PRINTING_OFFER FRM = new CSPSS.OFFER_MANAGE.PRINTING_OFFER();
                    FRM.Show();
                }
                catch (Exception)
                {
                    MessageBox.Show("网络连接中断");
                }

            }
            else if (v1 == "专色参数")
            {
                CSPSS.BASE_INFO.COLOR_PARAMETERS FRM = new CSPSS.BASE_INFO.COLOR_PARAMETERS();
                FRM.IDO = ccolor_parameters.GETID();
                FRM.Show();

            }
            else if (v1 == "表面处理")
            {
                CSPSS.BASE_INFO.SURFACE_PROCESSING FRM = new CSPSS.BASE_INFO.SURFACE_PROCESSING();
                FRM.IDO = csurface_processing.GETID();
                FRM.Show();

            }
            else if (v1 == "裱纸工艺")
            {
                CSPSS.BASE_INFO.LAMINATING_PROCESS FRM = new CSPSS.BASE_INFO.LAMINATING_PROCESS();
                FRM.IDO = claminating_process.GETID();
                FRM.Show();

            }

            else if (v1 == "部门信息维护")
            {
                CSPSS.BASE_INFO.DEPART FRM = new CSPSS.BASE_INFO.DEPART();
                FRM.IDO = cdepart.GETID();
                FRM.Show();

            }

            else if (v1 == "印刷选项")
            {
                CSPSS.BASE_INFO.PRINT_OPTION FRM = new CSPSS.BASE_INFO.PRINT_OPTION();
                FRM.IDO = cprint_option.GETID();
                FRM.PMID = cprinting_machine_size.GETID();

                FRM.Show();

            }
            else if (v1 == "印刷用面纸")
            {
                CSPSS.BASE_INFO.TISSUE_SPEC FRM = new CSPSS.BASE_INFO.TISSUE_SPEC();
                FRM.IDO = ctissue_sepc.GETID();
                FRM.Show();

            }
            else if (v1 == "复合用芯纸")
            {
                CSPSS.BASE_INFO.PAPER_CORE FRM = new CSPSS.BASE_INFO.PAPER_CORE();
                FRM.IDO = cpaper_core.GETID();
                FRM.Show();

            }
            else if (v1 == "原色")
            {
                CSPSS.BASE_INFO.PRIMARY_COLORS FRM = new CSPSS.BASE_INFO.PRIMARY_COLORS();
                FRM.IDO = cprimary_colors.GETID();
                FRM.Show();

            }

            else if (v1 == "职务信息维护")
            {
                CSPSS.BASE_INFO.POSITION FRM = new CSPSS.BASE_INFO.POSITION();
                FRM.IDO = cposition.GETID();
                FRM.Show();

            }
            else if (v1 == "服务器IP")
            {
                CSPSS.BASE_INFO.UPLOADFILE_DOMAIN FRM = new CSPSS.BASE_INFO.UPLOADFILE_DOMAIN();

                FRM.Show();

            }

            else if (v1 == "用户帐户")
            {
                CSPSS.USER_MANAGE.USER_INFO FRM = new CSPSS.USER_MANAGE.USER_INFO();
                FRM.IDO = cuser.GETID();
                FRM.ADD_OR_UPDATE = "ADD";
                FRM.Show();

            }
            else if (v1 == "更改密码")
            {
                CSPSS.USER_MANAGE.EDIT_PWD FRM = new CSPSS.USER_MANAGE.EDIT_PWD();
                FRM.Show();
            }
            else if (v1 == "权限管理")
            {
                CSPSS.USER_MANAGE.EDIT_RIGHT FRM = new CSPSS.USER_MANAGE.EDIT_RIGHT();
                FRM.IDO = cuser_group.GETID();
                FRM.Show();

            }
     
             #endregion
         }
         #endregion
         private void notifyIcon1_Click(object sender, EventArgs e)
         {
             click();//托盘单击事件
  
         }
         private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
         {
             click();//气泡单击事件
             showform();
         }
         private void notifyIcon1_BalloonTipClosed(object sender, EventArgs e)
         {
             click();//气泡关闭单击事件
             //MessageBox.Show("ok");
         }
         private void showform()
         {
           
             if (ID.Substring(0, 2) == "SR")//此为打样单
             {
                 OFFER_MANAGE.SAMPLE_RELY_LISTT FRM = new OFFER_MANAGE.SAMPLE_RELY_LISTT();
                 FRM.IDO = ID;
                 FRM.ShowDialog();
             }
             else if (ID.Substring(0, 2) == "PN")//此为生产指示书
             {
                 ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST FRM = new ORDER_MANAGE.PN_PRODUCTION_INSTRUCTIONST();
                 FRM.IDO = ID;
                 FRM.ShowDialog();
             }

         }
         private void click()
         {
             
             bc.getcom("UPDATE REMIND SET RECEIVE_STATUS='Y' WHERE RIID='" + ID + "' AND NOTICE_MAKERID='" + LOGIN.EMID + "'");
             timer2.Enabled = false;
      
             this.WindowState = FormWindowState.Maximized;
             ContextMenu c = new ContextMenu();
             MenuItem s = new MenuItem("退出");
             c.MenuItems.Add(s);
             notifyIcon1.ContextMenu = c;
             notifyIcon1.Icon = this.Icon = Resource1.xz_200X200;
             s.Click += new EventHandler(notify_Click);
             this.Show();

         }
         private void notify_Click(object sender, EventArgs e)
         {
             EXIT();
         }
         private void EXIT()
         {
             this.Dispose();
             notifyIcon1.Dispose();
             Application.Exit();
             string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
             bc.getcom(@"UPDATE AUTHORIZATION_USER SET STATUS='N' ,LEAVE_DATE='" + varDate + "'WHERE AUID='" + LOGIN.AUID + "'");
         }
         private void timer1_Tick(object sender, EventArgs e)
         {
             try
             {
                 bind();
                 bind_PN_PRODUCTION_INSTRUCTIONS();//生产指示书通知
             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }
         }
         private void timer2_Tick(object sender, EventArgs e)
         {
             try
             {
                 bind();
                 bind_PN_PRODUCTION_INSTRUCTIONS();//生产指示书通知
             }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }
         }
         #region bind
         private void bind()
         {
          
             dt = bc.getdt("SELECT * FROM REMIND WHERE RECEIVE_STATUS='N' AND SUBSTRING(RIID,1,2)='SR'");
             if (dt.Rows.Count > 0)
             {
                 foreach (DataRow dr in dt.Rows)
                 {
                     if (LOGIN.EMID == dr["NOTICE_MAKERID"].ToString())
                     {
                         timer2.Enabled = true;
                         DataTable dtx = bc.getdt(csample_rely_list.sql + " WHERE A.SRID='" + dr["RIID"].ToString() + "'" );
                         if (dtx.Rows.Count > 0)
                         {
                             SAMPLE_ID = dtx.Rows[0]["打样单号"].ToString();
                             ID = dtx.Rows[0]["打样编号"].ToString();
                         }
                         if (b == false)
                         {
                             notifyIcon1.Icon = this.Icon = Resource1.xz_200X200;
                             notifyIcon1.ShowBalloonTip(30, string.Format("{0} 打样单号 {1} 已签核完毕", LOGIN.ENAME,SAMPLE_ID),
                                 "提醒", ToolTipIcon.Info );
                             b = true;
                         }
                         else
                         {
                             notifyIcon1.Icon = Resource1.twinkle;
                             //notifyIcon1.ShowBalloonTip(30, "你好", "ok", ToolTipIcon.Info);
                             b = false;
                         }
                     }

                 }
             }
             try
             {

             }
             catch (Exception)
             {

             }
             //dataGridView1.DataSource = bc.getdt(sqlo);
         }
         #endregion
         #region bind_PN_PRODUCTION_INSTRUCTIONS
         private void bind_PN_PRODUCTION_INSTRUCTIONS()
         {
           
             dt = bc.getdt("SELECT * FROM REMIND WHERE RECEIVE_STATUS='N' AND SUBSTRING(RIID,1,2)='PN'");
             if (dt.Rows.Count > 0)
             {
              
                 foreach (DataRow dr in dt.Rows)
                 {
                     if (LOGIN.EMID == dr["NOTICE_MAKERID"].ToString())
                     {
                         timer2.Enabled = true;
                         DataTable dtx = bc.getdt(cpn_production_instruction .sql + " WHERE A.PNID='" + dr["RIID"].ToString() + "'" );
                         if (dtx.Rows.Count > 0)
                         {
                             SAMPLE_ID = dtx.Rows[0]["订单编号"].ToString();
                            
                         }
                         ID = dr["RIID"].ToString();
                         if (b == false)
                         {
                             notifyIcon1.Icon = this.Icon = Resource1.xz_200X200;
                             if (dr["NOTICE_OR_AUDIT"].ToString () == "AUDIT")
                             {
                                 notifyIcon1.ShowBalloonTip(30, string.Format("{0} 订单编号 {1} 已新增完毕 需要你签核", LOGIN.ENAME, SAMPLE_ID),
                                 "提醒", ToolTipIcon.Info);
                             }
                             else
                             {
                                 notifyIcon1.ShowBalloonTip(30, string.Format("{0} 订单编号 {1} 已签核完毕", LOGIN.ENAME, SAMPLE_ID),
                              "提醒", ToolTipIcon.Info);
                             }
                          
                             b = true;
                         }
                         else
                         {
                             notifyIcon1.Icon = new Icon(System.IO.Path.GetFullPath("Image/twinkle.ico"));
                             //notifyIcon1.ShowBalloonTip(30, "你好", "ok", ToolTipIcon.Info);
                             b = false;
                         }
                     }

                 }
             }
             try
             {

             }
             catch (Exception)
             {

             }
             //dataGridView1.DataSource = bc.getdt(sqlo);
         }
         #endregion
         private void MAIN_FormClosing(object sender, FormClosingEventArgs e)
         {
             e.Cancel = true;
             this.Hide();
         }

         private void MAIN_FormClosed(object sender, FormClosedEventArgs e)
         {
          

         }
   
         private void groupBox1_Paint(object sender, PaintEventArgs e)
         {
             e.Graphics.Clear(this.c2);
         }
    }
}
