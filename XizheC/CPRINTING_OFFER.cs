using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.SqlClient;
using XizheC;
using System.Windows.Forms;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;
namespace XizheC
{
    public class CPRINTING_OFFER
    {
        basec bc = new basec();
        CDOOR_PARAMETERS cdoor_parameters = new CDOOR_PARAMETERS();
        CPAPER_CORE cpaper_core = new CPAPER_CORE();
        CEDIT_RIGHT cedit_right = new CEDIT_RIGHT();
        CSAMPLE_RELY_LIST csample_rely_list = new CSAMPLE_RELY_LIST();
        COTHER_COST cother_cost = new COTHER_COST();
        CPRINT_DIE_CUTTING cprint_die_cutting = new CPRINT_DIE_CUTTING();
        CDIE_CUTTING_COST cdie_cutting_cost = new CDIE_CUTTING_COST();
        CPORTRAY cportray = new CPORTRAY();
        CPRINT_PORTRAY cprint_portray = new CPRINT_PORTRAY();
        CPARTS_AUXILIARY cparts_auxiliary = new CPARTS_AUXILIARY();
        CPRINT_PARTS_AUXILIARY cprint_parts_auxiliary = new CPRINT_PARTS_AUXILIARY();
        CPACK_MATERIAL cpack_material = new CPACK_MATERIAL();
        CPRINT_PACK_MATERIAL cprint_pack_material = new CPRINT_PACK_MATERIAL();
        CTRANSPORT ctransport = new CTRANSPORT();
        CPRINT_TRANSPORT cprint_transport = new CPRINT_TRANSPORT();
        CPRINT_ARTIFICIALL cprint_artificial = new CPRINT_ARTIFICIALL();
        CARTIFICIAL cartificial = new CARTIFICIAL();
        CPRINT_PURCHASE cprint_purchase = new CPRINT_PURCHASE();
        CPRINT_COST_TOTAL cprint_cost_total = new CPRINT_COST_TOTAL();
        CPURCHASE cpurchase = new CPURCHASE();
        CPRINT_OPTION cprint_option = new CPRINT_OPTION();
        CPRINTING_MACHINE_SIZE cprint_machine_size = new CPRINTING_MACHINE_SIZE();
        CPAPER_CORE_OPTION cpaper_core_option = new CPAPER_CORE_OPTION();
        StringBuilder sqb=new StringBuilder ();
        #region nature

        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private decimal _RETURN_BATCH_SUBTOTAL_ARTIFICIAL_SET;
        public decimal RETURN_BATCH_SUBTOTAL_ARTIFICIAL_SET
        {
            set { _RETURN_BATCH_SUBTOTAL_ARTIFICIAL_SET = value; }
            get { return _RETURN_BATCH_SUBTOTAL_ARTIFICIAL_SET; }
        }
        private decimal _RETURN_MAIN_DOSAGE_MANAGE;
        public decimal RETURN_MAIN_DOSAGE_MANAGE
        {
            set { _RETURN_MAIN_DOSAGE_MANAGE = value; }
            get { return _RETURN_MAIN_DOSAGE_MANAGE; }
        }
        private decimal _RETURN_BATCH_SUBTOTAL_MANAGE;
        public decimal RETURN_BATCH_SUBTOTAL_MANAGE
        {
            set { _RETURN_BATCH_SUBTOTAL_MANAGE = value; }
            get { return _RETURN_BATCH_SUBTOTAL_MANAGE; }
        }
        private decimal _RETURN_BATCH_SUBTOTAL_COST_SET;
        public decimal RETURN_BATCH_SUBTOTAL_COST_SET
        {
            set { _RETURN_BATCH_SUBTOTAL_COST_SET = value; }
            get { return _RETURN_BATCH_SUBTOTAL_COST_SET; }
        }

        private decimal _RETURN_MAIN_DOSAGE_PROFIT;
        public decimal RETURN_MAIN_DOSAGE_PROFIT
        {
            set { _RETURN_MAIN_DOSAGE_PROFIT = value; }
            get { return _RETURN_MAIN_DOSAGE_PROFIT; }
        }
        private decimal _RETURN_MAIN_DOSAGE_PURCHASE_COST;
        public decimal RETURN_MAIN_DOSAGE_PURCHASE_COST
        {
            set { _RETURN_MAIN_DOSAGE_PURCHASE_COST = value; }
            get { return _RETURN_MAIN_DOSAGE_PURCHASE_COST; }
        }
        private decimal _RETURN_YUAN_SET_PURCHASE_COST;
        public decimal RETURN_YUAN_SET_PURCHASE_COST
        {
            set { _RETURN_YUAN_SET_PURCHASE_COST = value; }
            get { return _RETURN_YUAN_SET_PURCHASE_COST; }
        }
        private decimal _MIN_LENGTH;
        public decimal MIN_LENGTH
        {
            set { _MIN_LENGTH = value; }
            get { return _MIN_LENGTH; }
        }
        private string _sqlse;
        public string sqlse
        {
            set { _sqlse = value; }
            get { return _sqlse; }

        }
        private string _CUSTOMER_TYPE;
        public string CUSTOMER_TYPE
        {
            set { _CUSTOMER_TYPE = value; }
            get { return _CUSTOMER_TYPE; }

        }
        private string _SAMPLE_CODE;
        public string SAMPLE_CODE
        {
            set { _SAMPLE_CODE = value; }
            get { return _SAMPLE_CODE; }

        }
        private string _CNAME;
        public string CNAME
        {
            set { _CNAME = value; }
            get { return _CNAME; }

        }

        private string _BRAND;
        public string BRAND
        {
            set { _BRAND = value; }
            get { return _BRAND; }

        }
        private bool _RECEPTION_USE;
        public bool RECEPTION_USE
        {
            set { _RECEPTION_USE = value; }
            get { return _RECEPTION_USE; }

        }
        private decimal _BATCH_TOTAL_NO_TAX;
        public decimal BATCH_TOTAL_NO_TAX
        {
            set { _BATCH_TOTAL_NO_TAX = value; }
            get { return _BATCH_TOTAL_NO_TAX; }
        }
        private decimal _BATCH_TOTAL_HAVE_TAX;
        public decimal BATCH_TOTAL_HAVE_TAX
        {
            set { _BATCH_TOTAL_HAVE_TAX = value; }
            get { return _BATCH_TOTAL_HAVE_TAX; }
        }
    
        private string _PIID;
        public string PIID
        {
            set { _PIID = value; }
            get { return _PIID; }

        }
        private string _OFFER_ID_SENVEN;
        public string OFFER_ID_SENVEN
        {
            set { _OFFER_ID_SENVEN = value; }
            get { return _OFFER_ID_SENVEN; }

        }
        private string _CHARGE_AUDIT_STATUS;
        public string CHARGE_AUDIT_STATUS
        {
            set { _CHARGE_AUDIT_STATUS = value; }
            get { return _CHARGE_AUDIT_STATUS; }

        }

        private decimal _PAPER_CORE_DOOR;
        public decimal PAPER_CORE_DOOR
        {
            set { _PAPER_CORE_DOOR = value; }
            get { return _PAPER_CORE_DOOR; }

        }
        private decimal _OPPOSITE_COLOR_TOTAL;
        public decimal OPPOSITE_COLOR_TOTAL
        {
            set { _OPPOSITE_COLOR_TOTAL = value; }
            get { return _OPPOSITE_COLOR_TOTAL; }

        }

        private decimal _PAPER_CORE_AVAILABLE;
        public decimal PAPER_CORE_AVAILABLE
        {
            set { _PAPER_CORE_AVAILABLE = value; }
            get { return _PAPER_CORE_AVAILABLE; }

        }
        private decimal _YUAN_SET_MANAGE;
        public decimal YUAN_SET_MANAGE
        {
            set { _YUAN_SET_MANAGE = value; }
            get { return _YUAN_SET_MANAGE; }

        }
        private decimal _YUAN_SET_PROFIT;
        public decimal YUAN_SET_PROFIT
        {
            set { _YUAN_SET_PROFIT = value; }
            get { return _YUAN_SET_PROFIT; }
        }

        private decimal _YUAN_SET_PURCHASE;
        public decimal YUAN_SET_PURCHASE
        {
            set { _YUAN_SET_PURCHASE = value; }
            get { return _YUAN_SET_PURCHASE; }

        }
        private decimal _YUAN_SET_NO_TAX;
        public decimal YUAN_SET_NO_TAX
        {
            set { _YUAN_SET_NO_TAX = value; }
            get { return _YUAN_SET_NO_TAX; }
        }

        private decimal _YUAN_SET_HAVE_TAX;
        public decimal YUAN_SET_HAVE_TAX
        {
            set { _YUAN_SET_HAVE_TAX = value; }
            get { return _YUAN_SET_HAVE_TAX; }

        }
        private decimal _MAIN_MANAGE;
        public decimal MAIN_MANAGE
        {
            set { _MAIN_MANAGE = value; }
            get { return _MAIN_MANAGE; }

        }
        private decimal _MAIN_PROFIT;
        public decimal MAIN_PROFIT
        {
            set { _MAIN_PROFIT = value; }
            get { return _MAIN_PROFIT; }

        }
        private decimal _MAIN_PURCHASE;
        public decimal MAIN_PURCHASE
        {
            set { _MAIN_PURCHASE = value; }
            get { return _MAIN_PURCHASE; }

        }
        private decimal _TISSUE_DOSAGE;
        public decimal TISSUE_DOSAGE
        {
            set { _TISSUE_DOSAGE = value; }
            get { return _TISSUE_DOSAGE; }

        }
        private decimal _PAPER_CORE_DOSAGE;
        public decimal PAPER_CORE_DOSAGE
        {
            set { _PAPER_CORE_DOSAGE = value; }
            get { return _PAPER_CORE_DOSAGE; }

        }
 
        private decimal _TISSUE_OUTSIDE_LOSE;
        public decimal TISSUE_OUTSIDE_LOSE
        {
            set { _TISSUE_OUTSIDE_LOSE = value; }
            get { return _TISSUE_OUTSIDE_LOSE; }

        }
        private decimal _TISSUE_INSIDE_LOSE;
        public decimal TISSUE_INSIDE_LOSE
        {
            set { _TISSUE_INSIDE_LOSE = value; }
            get { return _TISSUE_INSIDE_LOSE; }

        }
        private decimal _YUAN_SET_PURCHASE_PERCENT;
        public decimal YUAN_SET_PURCHASE_PERCENT
        {
            set { _YUAN_SET_PURCHASE_PERCENT = value; }
            get { return _YUAN_SET_PURCHASE_PERCENT; }
        }
        private decimal _MAIN_DOSAGE_PURCHASE_PERCENT;
        public decimal MAIN_DOSAGE_PURCHASE_PERCENT
        {
            set { _MAIN_DOSAGE_PURCHASE_PERCENT = value; }
            get { return _MAIN_DOSAGE_PURCHASE_PERCENT; }
        }
        private decimal _TOTAL_TISSUE;
        public decimal TOTAL_TISSUE
        {
            set { _TOTAL_TISSUE = value; }
            get { return _TOTAL_TISSUE; }

        }
        private decimal _TOTAL_PAPAER_CORE;
        public decimal TOTAL_PAPAER_CORE
        {
            set { _TOTAL_PAPAER_CORE = value; }
            get { return _TOTAL_PAPAER_CORE; }

        }
        private decimal _TOTAL_BODY_PAPER;
        public decimal TOTAL_BODY_PAPER
        {
            set { _TOTAL_BODY_PAPER = value; }
            get { return _TOTAL_BODY_PAPER; }

        }
     
        private decimal _TOTAL_COST_PORTRAY;
        public decimal TOTAL_COST_PORTRAY
        {
            set { _TOTAL_COST_PORTRAY = value; }
            get { return _TOTAL_COST_PORTRAY; }
        }
        private decimal _TOTAL_COST_ARTIFICIAL;
        public decimal TOTAL_COST_ARTIFICIAL
        {
            set { _TOTAL_COST_ARTIFICIAL = value; }
            get { return _TOTAL_COST_ARTIFICIAL; }
        }
        private decimal _TOTAL_COST_PURCHASE;
        public decimal TOTAL_COST_PURCHASE
        {
            set { _TOTAL_COST_PURCHASE = value; }
            get { return _TOTAL_COST_PURCHASE; }
        }
        private decimal _TOTAL_COST_PURCHASE_TWO;
        public decimal TOTAL_COST_PURCHASE_TWO
        {
            set { _TOTAL_COST_PURCHASE_TWO = value; }
            get { return _TOTAL_COST_PURCHASE_TWO; }
        }
        private decimal _TOTAL_COST_PARTS_AUXILIARY;
        public decimal TOTAL_COST_PARTS_AUXILIARY
        {
            set { _TOTAL_COST_PARTS_AUXILIARY = value; }
            get { return _TOTAL_COST_PARTS_AUXILIARY; }
        }
        private decimal _TOTAL_POSITIVE_AND_OPPOSITE_PRINTING;
        public decimal TOTAL_POSITIVE_AND_OPPOSITE_PRINTING
        {
            set { _TOTAL_POSITIVE_AND_OPPOSITE_PRINTING = value; }
            get { return _TOTAL_POSITIVE_AND_OPPOSITE_PRINTING; }

        }
        private decimal _TOTAL_COST_PACK_MATERIAL;
        public decimal TOTAL_COST_PACK_MATERIAL
        {
            set { _TOTAL_COST_PACK_MATERIAL = value; }
            get { return _TOTAL_COST_PACK_MATERIAL; }
        }
        private decimal _TOTAL_LAMINATING_PROCESS;
        public decimal TOTAL_LAMINATING_PROCESS
        {
            set { _TOTAL_LAMINATING_PROCESS = value; }
            get { return _TOTAL_LAMINATING_PROCESS; }

        }
        private decimal _TOTAL_DIE_CUTTING;
        public decimal TOTAL_DIE_CUTTING
        {
            set { _TOTAL_DIE_CUTTING = value; }
            get { return _TOTAL_DIE_CUTTING; }

        }
        private decimal _TOTAL_SURFACE_PROCESSING;
        public decimal TOTAL_SURFACE_PROCESSING
        {
            set { _TOTAL_SURFACE_PROCESSING = value; }
            get { return _TOTAL_SURFACE_PROCESSING; }

        }
        private decimal _TOTAL_COST_TISSUE;
        public decimal TOTAL_COST_TISSUE
        {
            set { _TOTAL_COST_TISSUE = value; }
            get { return _TOTAL_COST_TISSUE; }

        }
        private decimal _TOTAL_COST_PAPAER_CORE;
        public decimal TOTAL_COST_PAPAER_CORE
        {
            set { _TOTAL_COST_PAPAER_CORE = value; }
            get { return _TOTAL_COST_PAPAER_CORE; }

        }
        private decimal _TOTAL_COST_BODY_PAPER;
        public decimal TOTAL_COST_BODY_PAPER
        {
            set { _TOTAL_COST_BODY_PAPER = value; }
            get { return _TOTAL_COST_BODY_PAPER; }

        }
        private decimal _TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP;
        public decimal TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP
        {
            set { _TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP = value; }
            get { return _TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP; }

        }
        private decimal _TOTAL_COST_LAMINATING_PROCESS;
        public decimal TOTAL_COST_LAMINATING_PROCESS
        {
            set { _TOTAL_COST_LAMINATING_PROCESS = value; }
            get { return _TOTAL_COST_LAMINATING_PROCESS; }

        }
        private decimal _TOTAL_COST_CUTTING;
        public decimal TOTAL_COST_CUTTING
        {
            set { _TOTAL_COST_CUTTING = value; }
            get { return _TOTAL_COST_CUTTING; }
        }
  
        private decimal _TOTAL_COST_DIE_CUTTING;
        public decimal TOTAL_COST_DIE_CUTTING
        {
            set { _TOTAL_COST_DIE_CUTTING = value; }
            get { return _TOTAL_COST_DIE_CUTTING; }
        }
        private decimal _TOTAL_COST_SURFACE_PROCESSING;
        public decimal TOTAL_COST_SURFACE_PROCESSING
        {
            set { _TOTAL_COST_SURFACE_PROCESSING = value; }
            get { return _TOTAL_COST_SURFACE_PROCESSING; }

        }
        private decimal _TISSUE_DOOR;
        public decimal TISSUE_DOOR
        {
            set { _TISSUE_DOOR = value; }
            get { return _TISSUE_DOOR; }

        }
        private decimal _TISSUE_LENGTH;
        public decimal TISSUE_LENGTH
        {
            set { _TISSUE_LENGTH = value; }
            get { return _TISSUE_LENGTH; }

        }
        private decimal _PROCESSING_DOOR;
        public decimal PROCESSING_DOOR
        {
            set { _PROCESSING_DOOR = value; }
            get { return _PROCESSING_DOOR; }

        }
 
        private decimal _BODY_PAPER_INSIDE_LOSE;
        public decimal BODY_PAPER_INSIDE_LOSE
        {
            set { _BODY_PAPER_INSIDE_LOSE = value; }
            get { return _BODY_PAPER_INSIDE_LOSE; }

        }
        private decimal _BODY_PAPER_OUTSIDE_LOSE;
        public decimal BODY_PAPER_OUTSIDE_LOSE
        {
            set { _BODY_PAPER_OUTSIDE_LOSE = value; }
            get { return _BODY_PAPER_OUTSIDE_LOSE; }

        }

        private decimal _PAPER_CORE_LENGTH;
        public decimal PAPER_CORE_LENGTH
        {
            set { _PAPER_CORE_LENGTH = value; }
            get { return _PAPER_CORE_LENGTH; }

        }
        private decimal _PACK_LENGTH;
        public decimal PACK_LENGTH
        {
            set { _PACK_LENGTH = value; }
            get { return _PACK_LENGTH; }

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
        private decimal _MACHINING_MACHINE_FREE;
        public decimal MACHINING_MACHINE_FREE
        {
            set { _MACHINING_MACHINE_FREE = value; }
            get { return _MACHINING_MACHINE_FREE; }

        }
        private decimal _TOTAL_PRODUCT_NUMBER;
        public decimal TOTAL_PRODUCT_NUMBER
        {
            set { _TOTAL_PRODUCT_NUMBER = value; }
            get { return _TOTAL_PRODUCT_NUMBER; }

        }
        private decimal _LAMINATING_PROCESS_COUNT;
        public decimal LAMINATING_PROCESS_COUNT
        {
            set { _LAMINATING_PROCESS_COUNT = value; }
            get { return _LAMINATING_PROCESS_COUNT; }

        }
        private decimal _PROCESSING_LENGTH;
        public decimal PROCESSING_LENGTH
        {
            set { _PROCESSING_LENGTH = value; }
            get { return _PROCESSING_LENGTH; }

        }
        private decimal _POSITIVE_CTP_COUNT;
        public decimal POSITIVE_CTP_COUNT
        {
            set { _POSITIVE_CTP_COUNT = value; }
            get { return _POSITIVE_CTP_COUNT; }

        }

        private decimal _POSITIVE_PRINTING_TOTAL;
        public decimal POSITIVE_PRINTING_TOTAL
        {
            set { _POSITIVE_PRINTING_TOTAL = value; }
            get { return _POSITIVE_PRINTING_TOTAL; }

        }
        private decimal _DIE_CUTTING_PRICE;
        public decimal  DIE_CUTTING_PRICE
        {
            set { _DIE_CUTTING_PRICE = value; }
            get { return _DIE_CUTTING_PRICE; }

        }
        private decimal _OPPOSITE_PRINTING_TOTAL;
        public decimal OPPOSITE_PRINTING_TOTAL
        {
            set { _OPPOSITE_PRINTING_TOTAL = value; }
            get { return _OPPOSITE_PRINTING_TOTAL; }

        }
        private decimal _SUN_SCREEN_INK;
        public decimal SUN_SCREEN_INK
        {
            set { _SUN_SCREEN_INK = value; }
            get { return _SUN_SCREEN_INK; }

        }
        private decimal _PASS_COLOR_UNIT_PRICE;
        public decimal PASS_COLOR_UNIT_PRICE
        {
            set { _PASS_COLOR_UNIT_PRICE = value; }
            get { return _PASS_COLOR_UNIT_PRICE; }

        }
        private decimal _SQUARE_OR_METRE_MIN;
        public decimal SQUARE_OR_METRE_MIN
        {
            set { _SQUARE_OR_METRE_MIN = value; }
            get { return _SQUARE_OR_METRE_MIN; }

        }
        private decimal _TISSUE_ORDER;
        public decimal TISSUE_ORDER
        {
            set { _TISSUE_ORDER = value; }
            get { return _TISSUE_ORDER; }

        }
        private decimal _PRINTING_UNIT_PRICE;
        public decimal PRINTING_UNIT_PRICE
        {
            set { _PRINTING_UNIT_PRICE = value; }
            get { return _PRINTING_UNIT_PRICE; }

        }
        private decimal _LAMINATING_PROCESS_DOSAGE;
        public decimal LAMINATING_PROCESS_DOSAGE
        {
            set { _LAMINATING_PROCESS_DOSAGE = value; }
            get { return _LAMINATING_PROCESS_DOSAGE; }

        }
        private decimal _NEED_COUNT;
        public decimal NEED_COUNT
        {
            set { _NEED_COUNT = value; }
            get { return _NEED_COUNT; }

        }

        private string _BODY_WEIGHT;
        public string BODY_WEIGHT
        {
            set { _BODY_WEIGHT = value; }
            get { return _BODY_WEIGHT; }

        }
        private decimal  _MIN_PRINTING;
        public decimal  MIN_PRINTING
        {
            set { _MIN_PRINTING = value; }
            get { return _MIN_PRINTING; }

        }
        private decimal _POSITIVE_SUN_SCREEN_TOTAL;
        public decimal POSITIVE_SUN_SCREEN_TOTAL
        {
            set { _POSITIVE_SUN_SCREEN_TOTAL = value; }
            get { return _POSITIVE_SUN_SCREEN_TOTAL; }

        }
        private decimal _OPPOSITE_SUN_SCREEN_TOTAL;
        public decimal OPPOSITE_SUN_SCREEN_TOTAL
        {
            set { _OPPOSITE_SUN_SCREEN_TOTAL = value; }
            get { return _OPPOSITE_SUN_SCREEN_TOTAL; }

        }
        private decimal _POSITIVE_COLOR_COUNT_TOTAL;
        public decimal POSITIVE_COLOR_COUNT_TOTAL
        {
            set { _POSITIVE_COLOR_COUNT_TOTAL = value; }
            get { return _POSITIVE_COLOR_COUNT_TOTAL; }

        }

        private decimal _OUT_OF_PRINT;
        public decimal OUT_OF_PRINT
        {
            set { _OUT_OF_PRINT = value; }
            get { return _OUT_OF_PRINT; }

        }
        private decimal _SURFACE_NUMBER;
        public decimal SURFACE_NUMBER
        {
            set { _SURFACE_NUMBER = value; }
            get { return _SURFACE_NUMBER; }

        }
        private decimal _COLOR_COUNT;
        public decimal COLOR_COUNT
        {
            set { _COLOR_COUNT = value; }
            get { return _COLOR_COUNT; }

        }
        private decimal _CTP_UNIT_PRICE;
        public decimal CTP_UNIT_PRICE
        {
            set { _CTP_UNIT_PRICE = value; }
            get { return _CTP_UNIT_PRICE; }

        }
        private string _DET_WNAME;
        public string DET_WNAME
        {
            set { _DET_WNAME = value; }
            get { return _DET_WNAME; }

        }
        private string _UNIT_DOSAGE;
        public string UNIT_DOSAGE
        {
            set { _UNIT_DOSAGE = value; }
            get { return _UNIT_DOSAGE; }

        }
        private decimal _POSITIVE_CTP_PRICE_TOTAL;
        public decimal POSITIVE_CTP_PRICE_TOTAL
        {
            set { _POSITIVE_CTP_PRICE_TOTAL = value; }
            get { return _POSITIVE_CTP_PRICE_TOTAL; }

        }
        private decimal _MACHINING_MIN_PRINTING;
        public decimal MACHINING_MIN_PRINTING
        {
            set { _MACHINING_MIN_PRINTING = value; }
            get { return _MACHINING_MIN_PRINTING; }

        }
        private decimal _OPPOSITE_CTP_PRICE_TOTAL;
        public decimal OPPOSITE_CTP_PRICE_TOTAL
        {
            set { _OPPOSITE_CTP_PRICE_TOTAL = value; }
            get { return _OPPOSITE_CTP_PRICE_TOTAL; }

        }
        private string _PAPER_LENGTH;
        public string PAPER_LENGTH
        {
            set { _PAPER_LENGTH = value; }
            get { return _PAPER_LENGTH; }

        }
        private string _PRINT_OPTION;
        public string PRINT_OPTION
        {
            set { _PRINT_OPTION = value; }
            get { return _PRINT_OPTION; }

        }
        private string _TISSUE_SPEC;
        public string TISSUE_SPEC
        {
            set { _TISSUE_SPEC = value; }
            get { return _TISSUE_SPEC; }

        }
        private string _WEIGHT;
        public string WEIGHT
        {
            set { _WEIGHT = value; }
            get { return _WEIGHT; }

        }

        private string _DRAWING_DOOR;
        public string DRAWING_DOOR
        {
            set { _DRAWING_DOOR = value; }
            get { return _DRAWING_DOOR; }

        }
        private string _PAPER_CORE;
        public string PAPER_CORE
        {
            set { _PAPER_CORE = value; }
            get { return _PAPER_CORE; }

        }
        private string _SPEC;
        public string SPEC
        {
            set { _SPEC = value; }
            get { return _SPEC; }

        }
   
       private string _BODY_PAPER;
        public string BODY_PAPER
        {
            set { _BODY_PAPER = value; }
            get { return _BODY_PAPER; }

        }
        private string _BODY_POSITIVE_COLOR;
        public string BODY_POSITIVE_COLOR
        {
            set { _BODY_POSITIVE_COLOR = value; }
            get { return _BODY_POSITIVE_COLOR; }

        }
        private decimal _POSITIVE_4C;
        public decimal  POSITIVE_4C
        {
            set { _POSITIVE_4C = value; }
            get { return _POSITIVE_4C; }

        }
        private decimal _OPPOSITE_THE_PAPER_LOSE;
        public decimal OPPOSITE_THE_PAPER_LOSE
        {
            set { _OPPOSITE_THE_PAPER_LOSE = value; }
            get { return _OPPOSITE_THE_PAPER_LOSE; }

        }
        private string _POSITIVE_COLOR;
        public string POSITIVE_COLOR
        {
            set { _POSITIVE_COLOR = value; }
            get { return _POSITIVE_COLOR; }

        }

        private string _POSITIVE_SUN_SCREEN;
        public string POSITIVE_SUN_SCREEN
        {
            set { _POSITIVE_SUN_SCREEN = value; }
            get { return _POSITIVE_SUN_SCREEN; }

        }
      
        private string _PFID;
        public string PFID
        {
            set { _PFID = value; }
            get { return _PFID; }

        }
        private decimal  _COUNT;
        public decimal  COUNT
        {
            set { _COUNT = value; }
            get { return _COUNT; }

        }
        private decimal _LAMINATING_NUMBER;
        public decimal LAMINATING_NUMBER
        {
            set { _LAMINATING_NUMBER = value; }
            get { return _LAMINATING_NUMBER; }

        }
        private string _OFFER_ID;
        public string OFFER_ID
        {
            set { _OFFER_ID = value; }
            get { return _OFFER_ID; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        private decimal _POSITIVE_THE_PAPER_LOSE;
        public decimal POSITIVE_THE_PAPER_LOSE
        {
            set { _POSITIVE_THE_PAPER_LOSE = value; }
            get { return _POSITIVE_THE_PAPER_LOSE; }

        }
        private string _sql;
        public string sql
        {
            set { _sql = value; }
            get { return _sql; }

        }
        private string _sqlo;
        public string sqlo
        {
            set { _sqlo = value; }
            get { return _sqlo; }

        }
        private string _sqlt;
        public string sqlt
        {
            set { _sqlt = value; }
            get { return _sqlt; }

        }
        private string _sqlth;
        public string sqlth
        {
            set { _sqlth = value; }
            get { return _sqlth; }

        }
        private string _sqlf;
        public string sqlf
        {
            set { _sqlf = value; }
            get { return _sqlf; }

        }
        private string _sqlfi;
        public string sqlfi
        {
            set { _sqlfi = value; }
            get { return _sqlfi; }

        }
       
        private string _sqlsi;
        public string sqlsi
        {
            set { _sqlsi = value; }
            get { return _sqlsi; }

        }
        private string _sqlei;
        public string sqlei
        {
            set { _sqlei = value; }
            get { return _sqlei; }

        }
        private string _sqlni;
        public string sqlni
        {
            set { _sqlni = value; }
            get { return _sqlni; }
        }
        private string _sqlte;
        public string sqlte
        {
            set { _sqlte = value; }
            get { return _sqlte; }
        }
        private string _MAKERID;
        public string MAKERID
        {
            set { _MAKERID = value; }
            get { return _MAKERID; }

        }
        private string _PFKEY;
        public string PFKEY
        {
            set { _PFKEY = value; }
            get { return _PFKEY; }

        }
        private  bool _IFExecutionSUCCESS;
        public  bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }
        }
      
        private string _SN;
        public string SN
        {
            set { _SN = value; }
            get { return _SN; }

        }
        private string _DOUBLE_PRINTING;
        public string DOUBLE_PRINTING
        {
            set { _DOUBLE_PRINTING = value; }
            get { return _DOUBLE_PRINTING; }

        }
        private decimal  _OPPOSITE_4C;
        public decimal  OPPOSITE_4C
        {
            set { _OPPOSITE_4C = value; }
            get { return _OPPOSITE_4C; }

        }
        private decimal _POSITIVE_CTP_EDITION_FOR_PARAMETERS;
        public decimal POSITIVE_CTP_EDITION_FOR_PARAMETERS
        {
            set { _POSITIVE_CTP_EDITION_FOR_PARAMETERS = value; }
            get { return _POSITIVE_CTP_EDITION_FOR_PARAMETERS; }

        }
        private string _OPPOSITE_COLOR;
        public string OPPOSITE_COLOR
        {
            set { _OPPOSITE_COLOR = value; }
            get { return _OPPOSITE_COLOR; }

        }

        private string _LAMINATING_PROCESS;
        public string LAMINATING_PROCESS
        {
            set { _LAMINATING_PROCESS = value; }
            get { return _LAMINATING_PROCESS; }

        }
        private decimal _LAMINATING_PROCESS_PRICE;
        public decimal LAMINATING_PROCESS_PRICE
        {
            set { _LAMINATING_PROCESS_PRICE = value; }
            get { return _LAMINATING_PROCESS_PRICE; }

        }
     
        private decimal  _DIE_CUTTING;
        public decimal  DIE_CUTTING
        {
            set { _DIE_CUTTING = value; }
            get { return _DIE_CUTTING; }

        }
        private string _OFFER_MAKERID;
        public string OFFER_MAKERID
        {
            set { _OFFER_MAKERID = value; }
            get { return _OFFER_MAKERID; }

        }
        private string _OPPOSITE_SUN_SCREEN;
        public string OPPOSITE_SUN_SCREEN
        {
            set { _OPPOSITE_SUN_SCREEN = value; }
            get { return _OPPOSITE_SUN_SCREEN; }

        }

        private string _SURFACE_PROCESSING;
        public string SURFACE_PROCESSING
        {
            set { _SURFACE_PROCESSING = value; }
            get { return _SURFACE_PROCESSING; }

        }

 
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }

        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

        }
        #endregion
        decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0, d5 = 0,d6=0,d7=0,d8=0,d9=0,d10=0;
        DataTable dt = new DataTable();
        DataTable dtx = new DataTable();
        DataTable dtx2 = new DataTable();//印刷选项相关数据
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable dt4 = new DataTable();
        DataTable dt5 = new DataTable();
        DataTable dt6 = new DataTable();
        DataTable dt7 = new DataTable();
        DataTable dt8 = new DataTable();
        DataTable dtt = new DataTable();
        int i;
        #region sql
        #region sql
        string setsql = @"
SELECT 
C.PROJECT_NAME AS 项目名称,
D.CNAME AS 客户,
C.BRAND AS 品牌,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=C.AE_MAKERID_1) AS AE,
B.COUNT AS 数量,
C.PROJECT_ID  AS 项目号,
B.OFFER_ID AS 报价编号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 报价,
SUBSTRING(B.DATE,1,10) AS 日期,
A.DET_WNAME AS 部品名,
A.UNIT_DOSAGE_M AS 部品个数,
A.UNIT_DOSAGE_D AS 拼模数,
RTRIM(CONVERT(DECIMAL(18,2),A.UNIT_DOSAGE_M/A.UNIT_DOSAGE_D)) AS 部品数,
A.PRINT_OPTION  AS 印刷选项,
A.SN AS 项次,
A.DRAWING_DOOR  AS 图纸门幅,
A.PAPER_LENGTH  AS 图纸纸长,
A.PRINT_OPTION  AS 印刷选项,
A.TISSUE_SPEC  AS 面纸,
A.WEIGHT  AS 面纸克重,
A.PAPER_CORE  AS 芯纸,
A.SPEC  AS 芯纸规格,
A.BODY_PAPER  AS 底纸,
A.BODY_WEIGHT  AS 底纸克重,
A.POSITIVE_4C  AS 正面4C,
A.POSITIVE_COLOR  AS 正面专色,
A.POSITIVE_SUN_SCREEN  AS 正面防晒,
A.DOUBLE_PRINTING  AS 双面印刷,
A.OPPOSITE_4C  AS 反面4C,
A.OPPOSITE_COLOR  AS 反面专色,
A.OPPOSITE_SUN_SCREEN  AS 反面防晒,
A.SURFACE_PROCESSING  AS 表面加工,
A.SURFACE_COUNT  AS 表面次数,
A.LAMINATING_PROCESS  AS 裱纸工艺,
A.LAMINATING_COUNT  AS 裱纸次数,
A.DIE_CUTTING  AS 模切,
CASE WHEN B.CHARGE_AUDIT_STATUS='Y' THEN '已审核'
ELSE '未审核'
END AS 审核状态
FROM PRINTING_OFFER_DET A 
LEFT JOIN PRINTING_OFFER_MST B ON A.PFID=B.PFID
LEFT JOIN PROJECT_INFO C ON C.PIID=B.PIID
LEFT JOIN CUSTOMERINFO_MST D ON C.CUID=D.CUID

";
        #endregion
        #region sqlo
        string setsqlo = @"
INSERT INTO PRINTING_OFFER_DET
(
PFKEY,
PFID,
SN,
DET_WNAME,
DRAWING_DOOR ,
PAPER_LENGTH ,
UNIT_DOSAGE_M,
UNIT_DOSAGE_D,
PRINT_OPTION ,
TISSUE_SPEC ,
WEIGHT ,
PAPER_CORE ,
SPEC ,
BODY_PAPER ,
BODY_WEIGHT ,
POSITIVE_4C ,
POSITIVE_COLOR ,
POSITIVE_SUN_SCREEN ,
DOUBLE_PRINTING ,
OPPOSITE_4C ,
OPPOSITE_COLOR ,
OPPOSITE_SUN_SCREEN ,
SURFACE_PROCESSING ,
SURFACE_COUNT ,
LAMINATING_PROCESS ,
LAMINATING_COUNT ,
DIE_CUTTING ,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
)
VALUES
(
@PFKEY,
@PFID,
@SN,
@DET_WNAME,
@DRAWING_DOOR ,
@PAPER_LENGTH ,
@UNIT_DOSAGE_M,
@UNIT_DOSAGE_D,
@PRINT_OPTION ,
@TISSUE_SPEC ,
@WEIGHT ,
@PAPER_CORE ,
@SPEC ,
@BODY_PAPER ,
@BODY_WEIGHT ,
@POSITIVE_4C ,
@POSITIVE_COLOR ,
@POSITIVE_SUN_SCREEN ,
@DOUBLE_PRINTING ,
@OPPOSITE_4C ,
@OPPOSITE_COLOR ,
@OPPOSITE_SUN_SCREEN ,
@SURFACE_PROCESSING ,
@SURFACE_COUNT ,
@LAMINATING_PROCESS ,
@LAMINATING_COUNT ,
@DIE_CUTTING ,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY
)


";
        #endregion
        string setsqlt = @"

INSERT INTO PRINTING_OFFER_MST
(
PFID,
OFFER_ID,
PIID,
COUNT,
CHARGE_AUDIT_STATUS,
DATE,
EDIT_TIME,
MAKERID,
YEAR,
MONTH,
DAY
)
VALUES
(
@PFID,
@OFFER_ID,
@PIID,
@COUNT,
@CHARGE_AUDIT_STATUS,
@DATE,
@EDIT_TIME,
@MAKERID,
@YEAR,
@MONTH,
@DAY
)
";
        string setsqlth = @"
UPDATE PRINTING_OFFER_MST SET
COUNT=@COUNT,
EDIT_TIME=@EDIT_TIME
";

        string setsqlf = @"
SELECT 
B.PFID AS 编号,
B.OFFER_ID AS 报价编号,
C.PROJECT_ID AS 项目号,
C.PROJECT_NAME AS 项目名称,
B.COUNT AS 数量,
A.SN AS 项次,
A.DRAWING_DOOR  AS 图纸门幅,
A.PAPER_LENGTH  AS 图纸纸长,
A.PRINT_OPTION  AS 印刷选项,
A.TISSUE_SPEC  AS 面纸,
A.WEIGHT  AS 面纸克重,
A.PAPER_CORE  AS 芯纸,
A.SPEC  AS 芯纸规格,
A.BODY_PAPER  AS 底纸,
A.BODY_WEIGHT  AS 底纸克重,
A.POSITIVE_4C  AS 正面4C,
A.POSITIVE_COLOR  AS 正面专色,
A.POSITIVE_SUN_SCREEN  AS 正面防晒,
A.DOUBLE_PRINTING  AS 双面印刷,
A.OPPOSITE_4C  AS 反面4C,
A.OPPOSITE_COLOR  AS 反面专色,
A.OPPOSITE_SUN_SCREEN  AS 反面防晒,
A.SURFACE_PROCESSING  AS 表面加工,
A.SURFACE_COUNT  AS 表面次数,
A.LAMINATING_PROCESS  AS 裱纸工艺,
A.LAMINATING_COUNT  AS 裱纸次数,
A.DIE_CUTTING  AS 模切,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 制单人,
B.DATE AS 制单日期,
CASE WHEN B.CHARGE_AUDIT_STATUS='Y' THEN '已审核'
ELSE '未审核'
END AS 审核状态,
C.BRAND AS 品牌,
D.CNAME AS 客户名称
FROM PRINTING_OFFER_DET A 
LEFT JOIN PRINTING_OFFER_MST B ON A.PFID=B.PFID
LEFT JOIN PROJECT_INFO C ON C.PIID=B.PIID
LEFT JOIN CustomerInfo_MST D ON C.CUID=D.CUID 
";
        string setsqlfi = @"
SELECT 
B.PROJECT_NAME AS 项目名称,
B.PROJECT_ID AS 项目号,
B.BRAND AS 品牌,
A.SAMPLE_ID AS 打样单号,
C.CName AS 客户,
D.EName  AS AE
FROM  SAMPLE_RELY_LIST  A
LEFT JOIN PROJECT_INFO B ON B.PROJECT_ID =SUBSTRING (A.SAMPLE_ID,1,LEN(A.SAMPLE_ID)-3)
LEFT JOIN CustomerInfo_MST C ON B.CUID=C.CUID 
LEFT JOIN EmployeeInfo D ON B.AE_MAKERID_1 =D.EMID 

";
        string setsqlsi = @"
SELECT 
A.OFFER_ID AS 报价编号,
A.COUNT AS 报价数量,
A.AUDIT_OPINION AS 审核批注,
B.PROJECT_ID AS 项目号,
B.PROJECT_NAME AS 项目名称,
B.BRAND AS 品牌,
C.CName AS 客户,
D.EName  AS AE,
E.ENAME AS 制单人,
A.DATE AS 制单日期,
A.EDIT_TIME AS 修改日期
FROM PRINTING_OFFER_MST  A
LEFT JOIN PROJECT_INFO B ON A.PIID =B.PIID 
LEFT JOIN CustomerInfo_MST C ON B.CUID=C.CUID 
LEFT JOIN EmployeeInfo D ON B.AE_MAKERID_1 =D.EMID 
LEFT JOIN EMPLOYEEINFO E ON A.MAKERID=E.EMID
ORDER BY A.OFFER_ID ASC

";
        string setsqlse = @"
SELECT 
C.PROJECT_NAME AS 项目名称,
D.CNAME AS 客户,
C.BRAND AS 品牌,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=C.AE_MAKERID_1) AS AE,
B.COUNT AS 数量,
C.PROJECT_ID  AS 项目号,
B.OFFER_ID AS 报价编号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 报价,
SUBSTRING(B.DATE,1,10) AS 日期
FROM PRINTING_OFFER_MST B
LEFT JOIN PROJECT_INFO C ON B.PIID=C.PIID
LEFT JOIN CUSTOMERINFO_MST D ON C.CUID=D.CUID


";
        #region sqlei
        string setsqlei = @"

INSERT INTO PRINT_TOTAL
(
PTID,
PFID,
C1,
C9,
C10,
C11,
C12,
C13,
C14,
C15,
C16,
C17,
C18,
C19,
C20,
C21,
C22,
C23,
C24,
C25,
C26,
C27,
C28,
C29,
C30,
C31,
C32,
C33,
C34,
C35,
C36,
C37,
C38,
C39,
C40,
C41,
C42,
C43,
C44,
C45,
C46,
C47,
C48,
C49,
C50,
C51,
C52,
C53,
C54,
C55,
C56,
C57,
C58,
C59,
C60,
C61,
C62,
C63,
C64,
C65,
MakerID,
Date,
Year,
Month,
DAY
)
VALUES
(
@PTID,
@PFID,
@C1,
@C9,
@C10,
@C11,
@C12,
@C13,
@C14,
@C15,
@C16,
@C17,
@C18,
@C19,
@C20,
@C21,
@C22,
@C23,
@C24,
@C25,
@C26,
@C27,
@C28,
@C29,
@C30,
@C31,
@C32,
@C33,
@C34,
@C35,
@C36,
@C37,
@C38,
@C39,
@C40,
@C41,
@C42,
@C43,
@C44,
@C45,
@C46,
@C47,
@C48,
@C49,
@C50,
@C51,
@C52,
@C53,
@C54,
@C55,
@C56,
@C57,
@C58,
@C59,
@C60,
@C61,
@C62,
@C63,
@C64,
@C65,
@MakerID,
@Date,
@Year,
@Month,
@DAY
)


";
        #endregion
        string setsqlni = @"
SELECT 
A.PFID AS 报价ID,
D.CNAME AS 客户,
C.BRAND AS 品牌,
A.C1 AS 序号,
C.PROJECT_NAME AS 项目名称,
B.COUNT AS 数量,
C.PROJECT_ID AS 项目号,
B.OFFER_ID AS 报价编号,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=C.AE_MAKERID_1) AS AE,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=B.MAKERID) AS 报价,
SUBSTRING(B.DATE,1,10) AS 日期,
E.DET_WNAME AS 部品名,
A.C9 AS 加工门幅,
A.C10 AS 加工长度,
A.C11 AS 部品总数,
A.C12 AS 机器型号,
A.C13 AS 部品单价,
A.C14 AS 部品总价,
A.C15 AS 面纸单价,
A.C16 AS 面纸用量,
A.C17 AS 面纸内耗,
A.C18 AS 面纸下单,
A.C19 AS 面纸外耗,
A.C20 AS 面纸门幅,
A.C21 AS 面纸纸长,
A.C22 AS 面纸可用,
A.C23 AS 面纸单个用量,
A.C24 AS 面纸小计,
A.C25 AS 芯纸单价,
A.C26 AS 芯纸内耗,
A.C27 AS 芯纸用量,
A.C28 AS 芯纸门幅,
A.C29 AS 芯纸纸长,
A.C30 AS 芯纸可用,
A.C31 AS 芯纸单个用量,
A.C32 AS 芯纸小计,
A.C33 AS 底纸单价,
A.C34 AS 底纸用量,
A.C35 AS 底纸内耗,
A.C36 AS 底纸下单,
A.C37 AS 底纸外耗,
A.C38 AS 底纸单个用量,
A.C39 AS 底纸小计,
A.C40 AS 印工单色单价,
A.C41 AS 超出单色单张价,
A.C42 AS CTP单张价,
A.C43 AS 正面色数共计,
A.C44 AS 正面CTP张数,
A.C45 AS 正面纸张损耗,
A.C46 AS 正面防晒合计,
A.C47 AS 正面CTP价计,
A.C48 AS 正面印工合计,
A.C49 AS 反面色数共计,
A.C50 AS 反面CTP张数,
A.C51 AS 反面纸张损耗,
A.C52 AS 反面防晒合计,
A.C53 AS 反面CTP价计,
A.C54 AS 反面印工合计,
A.C55 AS 正反CTP合计,
A.C56 AS 正反印工合计,
A.C57 AS 表面处理单价,
A.C58 AS 无印刷表面处理损耗,
A.C59 AS 表面处理用量,
A.C60 AS 表面加工小计,
A.C61 AS 裱工单价,
A.C62 AS 裱工用量,
A.C63 AS 裱工小计,
A.C64 AS 刀模小计,
A.C65 AS 模切小计,
E.TISSUE_SPEC  AS 面纸,
E.WEIGHT  AS 面纸克重,
E.PAPER_CORE  AS 芯纸,
E.SPEC  AS 芯纸规格,
E.BODY_PAPER  AS 底纸,
E.BODY_WEIGHT  AS 底纸克重,
E.SURFACE_PROCESSING  AS 表面加工,
E.LAMINATING_PROCESS  AS 裱纸工艺,
E.PRINT_OPTION AS 印刷选项,
E.DIE_CUTTING AS 模切
FROM
PRINT_TOTAL A
LEFT JOIN PRINTING_OFFER_MST B ON A.PFID=B.PFID
LEFT JOIN PROJECT_INFO C ON C.PIID=B.PIID
LEFT JOIN CUSTOMERINFO_MST D ON C.CUID=D.CUID
LEFT JOIN PRINTING_OFFER_DET E ON A.PFID=E.PFID AND A.C1=E.SN 
";
        string setsqlte = @"
SELECT 
B.PFID AS 编号,
B.OFFER_ID AS 报价编号,
C.PROJECT_ID AS 项目号,
C.PROJECT_NAME AS 项目名称,
CASE WHEN B.CHARGE_AUDIT_STATUS='Y' THEN '已审核'
ELSE '未审核'
END AS 审核状态,
C.BRAND AS 品牌,
D.CNAME AS 客户名称
FROM PRINTING_OFFER_MST B
LEFT JOIN PROJECT_INFO C ON C.PIID=B.PIID
LEFT JOIN CustomerInfo_MST D ON C.CUID=D.CUID 
";
        #endregion
        #region CPRINTING_OFFER()
        public CPRINTING_OFFER()
        {
            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
            sqlfi = setsqlfi;
            sqlsi = setsqlsi;
            sqlse = setsqlse;
            sqlei = setsqlei;
            sqlni = setsqlni;
            sqlte = setsqlte;
        }
        #endregion
        #region GetTableInfo
        public DataTable GetTableInfo()
        {
            DataTable dt = new DataTable();
            //dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("部品名", typeof(string));
            dt.Columns.Add("图纸门幅", typeof(string));
            dt.Columns.Add("图纸纸长", typeof(string ));
            dt.Columns.Add("部品个数", typeof(decimal));
            dt.Columns.Add("拼模数", typeof(decimal));
            dt.Columns.Add("部品数", typeof(decimal));
            dt.Columns.Add("印刷选项", typeof(string));
            dt.Columns.Add("面纸", typeof(string));
            dt.Columns.Add("面纸克重", typeof(string));
            dt.Columns.Add("芯纸", typeof(string));
            dt.Columns.Add("芯纸规格", typeof(string));
            dt.Columns.Add("底纸", typeof(string));
            dt.Columns.Add("底纸克重", typeof(string));
            dt.Columns.Add("正面4C", typeof(string));
            dt.Columns.Add("正面专色", typeof(string));
            dt.Columns.Add("正面防晒", typeof(string));
            dt.Columns.Add("双面印刷", typeof(string));
            dt.Columns.Add("反面4C", typeof(string));
            dt.Columns.Add("反面专色", typeof(string));
            dt.Columns.Add("反面防晒", typeof(string));
            dt.Columns.Add("表面加工", typeof(string));
            dt.Columns.Add("表面次数", typeof(string));
            dt.Columns.Add("裱纸工艺", typeof(string));
            dt.Columns.Add("裱纸次数", typeof(string));
            dt.Columns.Add("模切", typeof(string));
            return dt;
        }

        #endregion
        #region GetTableInfo_show_all
        public DataTable GetTableInfo_show_all()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("客户", typeof(string));
            dt.Columns.Add("品牌", typeof(string));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("打样单号", typeof(string));
            dt.Columns.Add("打样金额", typeof(string));
            dt.Columns.Add("报价编号", typeof(string));
            dt.Columns.Add("报价数量", typeof(string));
            dt.Columns.Add("报出价", typeof(string));
            dt.Columns.Add("审核批注", typeof(string));
            dt.Columns.Add("报价", typeof(string));
            dt.Columns.Add("项目号", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("审核状态", typeof(string));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("部品名", typeof(string));
            dt.Columns.Add("图纸门幅", typeof(string));
            dt.Columns.Add("图纸纸长", typeof(string));
            dt.Columns.Add("部品数", typeof(decimal));
            dt.Columns.Add("印刷选项", typeof(string));
            dt.Columns.Add("面纸", typeof(string));
            dt.Columns.Add("面纸克重", typeof(decimal));
            dt.Columns.Add("芯纸", typeof(string));
            dt.Columns.Add("芯纸规格", typeof(string));
            dt.Columns.Add("底纸", typeof(string));
            dt.Columns.Add("底纸克重", typeof(decimal));
            dt.Columns.Add("正面4C", typeof(string));
            dt.Columns.Add("正面专色", typeof(string));
            dt.Columns.Add("正面防晒", typeof(string));
            dt.Columns.Add("双面印刷", typeof(string));
            dt.Columns.Add("反面4C", typeof(string));
            dt.Columns.Add("反面专色", typeof(string));
            dt.Columns.Add("反面防晒", typeof(string));
            dt.Columns.Add("表面加工", typeof(string));
            dt.Columns.Add("表面次数", typeof(string));
            dt.Columns.Add("裱纸工艺", typeof(string));
            dt.Columns.Add("裱纸次数", typeof(string));
            dt.Columns.Add("模切", typeof(string));
            dt.Columns.Add("加工门幅", typeof(string));
            dt.Columns.Add("加工长度", typeof(string));
            dt.Columns.Add("部品总数", typeof(decimal));
            dt.Columns.Add("部品个数", typeof(decimal));
            dt.Columns.Add("拼模数", typeof(decimal));
            dt.Columns.Add("机器型号", typeof(string));
            dt.Columns.Add("部品单价", typeof(decimal));
            dt.Columns.Add("部品总价", typeof(string));
            dt.Columns.Add("面纸单价", typeof(decimal));
            dt.Columns.Add("面纸用量", typeof(decimal));
            dt.Columns.Add("面纸内耗", typeof(string));
            dt.Columns.Add("面纸下单", typeof(string));
            dt.Columns.Add("面纸外耗", typeof(string));
            dt.Columns.Add("面纸门幅", typeof(string));
            dt.Columns.Add("面纸纸长", typeof(string));
            dt.Columns.Add("面纸可用", typeof(string));
            dt.Columns.Add("面纸单个用量", typeof(decimal));
            dt.Columns.Add("面纸小计", typeof(decimal));
            dt.Columns.Add("芯纸单价", typeof(string));
            dt.Columns.Add("芯纸内耗", typeof(string));
            dt.Columns.Add("芯纸用量", typeof(string));
            dt.Columns.Add("芯纸门幅", typeof(string));
            dt.Columns.Add("芯纸纸长", typeof(string));
            dt.Columns.Add("芯纸可用", typeof(string));
            dt.Columns.Add("芯纸单个用量", typeof(decimal));
            dt.Columns.Add("芯纸小计", typeof(decimal));
            dt.Columns.Add("底纸单价", typeof(decimal));
            dt.Columns.Add("底纸用量", typeof(string));
            dt.Columns.Add("底纸内耗", typeof(string));
            dt.Columns.Add("底纸下单", typeof(string));
            dt.Columns.Add("底纸外耗", typeof(string));
            dt.Columns.Add("底纸单个用量", typeof(decimal));
            dt.Columns.Add("底纸小计", typeof(decimal));
            dt.Columns.Add("印工单色单价", typeof(string));
            dt.Columns.Add("超出单色单张价", typeof(string));
            dt.Columns.Add("CTP单张价", typeof(string));
            dt.Columns.Add("正面色数共计", typeof(string));
            dt.Columns.Add("正面CTP张数", typeof(string));
            dt.Columns.Add("正面纸张损耗", typeof(string));
            dt.Columns.Add("正面防晒合计", typeof(string));
            dt.Columns.Add("正面CTP价计", typeof(string));
            dt.Columns.Add("正面印工合计", typeof(string));
            dt.Columns.Add("反面色数共计", typeof(string));
            dt.Columns.Add("反面CTP张数", typeof(string));
            dt.Columns.Add("反面纸张损耗", typeof(string));
            dt.Columns.Add("反面防晒合计", typeof(string));
            dt.Columns.Add("反面CTP价计", typeof(string));
            dt.Columns.Add("反面印工合计", typeof(string));
            dt.Columns.Add("正反CTP合计", typeof(decimal));
            dt.Columns.Add("正反印工合计", typeof(decimal));
            dt.Columns.Add("表面处理单价", typeof(string));
            dt.Columns.Add("无印刷表面处理损耗", typeof(string));
            dt.Columns.Add("表面处理用量", typeof(decimal));
            dt.Columns.Add("表面加工小计", typeof(decimal));
            dt.Columns.Add("裱工单价", typeof(string));
            dt.Columns.Add("裱工用量", typeof(string));
            dt.Columns.Add("裱工小计", typeof(decimal));
            dt.Columns.Add("刀模小计", typeof(string));
            dt.Columns.Add("模切小计", typeof(decimal));
            return dt;
        }

        #endregion
        #region GetTableInfo_search
        public DataTable GetTableInfo_search()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("项目号", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("客户", typeof(string));
            dt.Columns.Add("品牌", typeof(string));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("打样单号", typeof(string));
            dt.Columns.Add("打样金额", typeof(string));
            dt.Columns.Add("报价编号", typeof(string));
            dt.Columns.Add("报价数量", typeof(string));
            dt.Columns.Add("报出价", typeof(string));
            dt.Columns.Add("审核批注", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            dt.Columns.Add("制单日期", typeof(string));
            dt.Columns.Add("修改日期", typeof(string));
            return dt;
        }

        #endregion
        #region GetTableInfo_show_hide
        public DataTable GetTableInfo_show_hide()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("项目号", typeof(string));
            dt.Columns.Add("报价ID", typeof(string));
            dt.Columns.Add("客户", typeof(string));
            dt.Columns.Add("品牌", typeof(string));
            dt.Columns.Add("AE", typeof(string));
            dt.Columns.Add("打样单号", typeof(string));
            dt.Columns.Add("打样金额", typeof(string));
            dt.Columns.Add("报价编号", typeof(string));
            dt.Columns.Add("报价数量", typeof(string));
            dt.Columns.Add("报出价", typeof(string));
            dt.Columns.Add("审核批注", typeof(string));
            dt.Columns.Add("报价", typeof(string));

            dt.Columns.Add("面纸", typeof(string));
            dt.Columns.Add("面纸克重", typeof(string));
            dt.Columns.Add("芯纸", typeof(string));
            dt.Columns.Add("芯纸规格", typeof(string));
            dt.Columns.Add("底纸", typeof(string));
            dt.Columns.Add("底纸克重", typeof(string));
            dt.Columns.Add("表面加工", typeof(string));
            dt.Columns.Add("裱纸工艺", typeof(string));

            dt.Columns.Add("印刷选项", typeof(string));
            dt.Columns.Add("模切", typeof(string));
            dt.Columns.Add("日期", typeof(string));
            dt.Columns.Add("部品名", typeof(string));
            dt.Columns.Add("加工门幅", typeof(decimal));
            dt.Columns.Add("加工长度", typeof(decimal));
            dt.Columns.Add("部品总数", typeof(decimal));
            dt.Columns.Add("机器型号", typeof(string));
            dt.Columns.Add("部品单价", typeof(decimal));
            dt.Columns.Add("部品总价", typeof(string));
            dt.Columns.Add("面纸单价", typeof(decimal));
            dt.Columns.Add("面纸用量", typeof(decimal));
            dt.Columns.Add("面纸内耗", typeof(string));
            dt.Columns.Add("面纸下单", typeof(string));
            dt.Columns.Add("面纸外耗", typeof(string));
            dt.Columns.Add("面纸门幅", typeof(string));
            dt.Columns.Add("面纸纸长", typeof(string));
            dt.Columns.Add("面纸可用", typeof(string));
            dt.Columns.Add("面纸单个用量", typeof(decimal));
            dt.Columns.Add("面纸小计", typeof(decimal));
            dt.Columns.Add("芯纸单价", typeof(decimal));
            dt.Columns.Add("芯纸内耗", typeof(decimal));
            dt.Columns.Add("芯纸用量", typeof(decimal));
            dt.Columns.Add("芯纸门幅", typeof(string));
            dt.Columns.Add("芯纸纸长", typeof(string));
            dt.Columns.Add("芯纸可用", typeof(string));
            dt.Columns.Add("芯纸单个用量", typeof(decimal));
            dt.Columns.Add("芯纸小计", typeof(decimal));
            dt.Columns.Add("底纸单价", typeof(string));
            dt.Columns.Add("底纸用量", typeof(string));
            dt.Columns.Add("底纸内耗", typeof(decimal));
            dt.Columns.Add("底纸下单", typeof(decimal));
            dt.Columns.Add("底纸外耗", typeof(string));
            dt.Columns.Add("底纸单个用量", typeof(decimal ));
            dt.Columns.Add("底纸小计", typeof(decimal));
            dt.Columns.Add("印工单色单价", typeof(string));
            dt.Columns.Add("超出单色单张价", typeof(string));
            dt.Columns.Add("CTP单张价", typeof(decimal));
            dt.Columns.Add("正面色数共计", typeof(string));
            dt.Columns.Add("正面CTP张数", typeof(string));
            dt.Columns.Add("正面纸张损耗", typeof(string));
            dt.Columns.Add("正面防晒合计", typeof(decimal));
            dt.Columns.Add("正面CTP价计", typeof(string));
            dt.Columns.Add("正面印工合计", typeof(decimal));
            dt.Columns.Add("反面色数共计", typeof(string));
            dt.Columns.Add("反面CTP张数", typeof(string));
            dt.Columns.Add("反面纸张损耗", typeof(string));
            dt.Columns.Add("反面防晒合计", typeof(string));
            dt.Columns.Add("反面CTP价计", typeof(string));
            dt.Columns.Add("反面印工合计", typeof(string));
            dt.Columns.Add("正反CTP合计", typeof(decimal));
            dt.Columns.Add("正反印工合计", typeof(decimal));
            dt.Columns.Add("表面处理单价", typeof(string));
            dt.Columns.Add("无印刷表面处理损耗", typeof(string));
            dt.Columns.Add("表面处理用量", typeof(decimal));
            dt.Columns.Add("表面加工小计", typeof(decimal));
            dt.Columns.Add("裱工单价", typeof(string));
            dt.Columns.Add("裱工用量", typeof(decimal));
            dt.Columns.Add("裱工小计", typeof(decimal));
            dt.Columns.Add("刀模小计", typeof(decimal));
            dt.Columns.Add("模切小计", typeof(decimal));
            return dt;
        }

        #endregion
        #region GETID
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from PRINTING_OFFER_NO", "PFID", "PF");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
                bc.getcom("INSERT INTO PRINTING_OFFER_NO(PFID,DATE,YEAR,MONTH) VALUES ('" + v1 + "','" + varDate + "','" + year + 
                    "','" + month + "')");
            }
            return GETID;
        }
        #endregion
        #region GETID_OFFER_ID
        public void  GETID_OFFER_ID(string SAMPLE_CODE_FIRST,string OFFER_TYPE_CODE)
        { //1610Z002-03-ADM
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            SAMPLE_CODE = bc.getOnlyString("SELECT SAMPLE_CODE FROM EMPLOYEEINFO WHERE EMID='" +MAKERID + "'");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            OFFER_ID_SENVEN = "";
            OFFER_ID = "";
            if (!bc.exists("SELECT * FROM PRINTING_OFFER_ID_NO WHERE PIID='" + PIID + "' AND YEAR='"+year +"' AND MONTH='"+month +"'"))//且年且月 161010
            {
                
                OFFER_ID_SENVEN = numYY(8, 3, "001", "select * from PRINTING_OFFER_ID_NO", "OFFER_ID_SENVEN", OFFER_TYPE_CODE);
                OFFER_ID = bc.numNOYMD(11, 2, "01", "select * from PRINTING_OFFER_MST WHERE SUBSTRING(OFFER_ID,1,8)='" + OFFER_ID_SENVEN + "'", "OFFER_ID",
                OFFER_ID_SENVEN + "-");
                OFFER_ID = OFFER_ID + "-" + SAMPLE_CODE + SAMPLE_CODE_FIRST;
                //编码时没有按OFFER_ID_SENVEN 排序导致写入重复编号
                //出现写入重复的OFFER_ID_SENVEN使得有重复的报价编号，
               
                if (!bc.exists("SELECT * FROM PRINTING_OFFER_ID_NO WHERE OFFER_ID_SENVEN='" + OFFER_ID_SENVEN + 
                    "' AND YEAR='"+year +"' AND MONTH='"+month +"'"))
                {
                    
                    basec.getcoms(@"
INSERT INTO PRINTING_OFFER_ID_NO
(PIID,
OFFER_ID_SENVEN,
YEAR,
MONTH,
DAY,
MAKERID,
DATE
) VALUES ('" + PIID + "','" + OFFER_ID_SENVEN +
                        "','" + year + 
                        "','" + month +
                        "','"+day +
                        "','"+MAKERID +
                        "','" + varDate + "')");//加此判断条件，写入时判断这7码是吗存在系统160523
                }
            }
            else
            {
                OFFER_ID_SENVEN = bc.getOnlyString(string.Format("SELECT OFFER_ID_SENVEN FROM PRINTING_OFFER_ID_NO WHERE PIID='{0}'  AND YEAR='" + year + 
                    "' AND MONTH='" + month + "'", PIID));//且年且月 161010
                OFFER_ID = bc.numNOYMD(11, 2, "01", "select * from PRINTING_OFFER_MST WHERE SUBSTRING(OFFER_ID,1,8)='" + OFFER_ID_SENVEN + "'", "OFFER_ID",
                OFFER_ID_SENVEN + "-");
                OFFER_ID = OFFER_ID + "-" + SAMPLE_CODE + SAMPLE_CODE_FIRST;
            }
            string GETID = "";
            if (OFFER_ID  != "Exceed Limited")
            {
                GETID = OFFER_ID;

            }

        }
        #endregion
        #region 编号 YY
        public string numYY(int digit, int wcodedigit, string wcode, string sql, string tbColumns, string prifix)
        {
            string year, month, day;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            day = DateTime.Now.ToString("dd");
            string P_str_Code, t, r, sql1, q = "";
            int P_int_Code, w, w1;//由于PIID为主键，如果不指定OFFER_ID_SENVEN为排序方式的话，默认是按PIID排序，实际用的是OFFER_ID_SENVEN的最后一行
            sql1 = sql + string.Format (" WHERE  YEAR='{0}' AND LEN({1})=8 AND MONTH={2} ORDER BY OFFER_ID_SENVEN ASC",year,tbColumns,month);
            DataTable dt = bc.getdt(sql1);

            if (dt.Rows.Count > 0)
            {
                P_str_Code = Convert.ToString(dt.Rows[(dt.Rows.Count - 1)][tbColumns]);
                w1 = digit - wcodedigit;
                P_int_Code = Convert.ToInt32(P_str_Code.Substring(w1, wcodedigit)) + 1;
                t = Convert.ToString(P_int_Code);
                w = wcodedigit - t.Length;
                if (w >= 0)
                {
                    while (w >= 1)
                    {
                        q = q + "0";
                        w = w - 1;

                    }
                    r = year + month + prifix + q + P_int_Code;
                }
                else
                {
                    r = "Exceed Limited";

                }

            }
            else
            {
                r = year + month + prifix + wcode;
            }
            return r;
        }
        #endregion
        #region RETURN_MONTH
        public string RETURN_MONTH(string MONTH_VALUE)
        {
            string GETID = "";
            if (MONTH_VALUE == "01")
            {
                GETID = "A";
            }
            else if (MONTH_VALUE == "02")
            {
                GETID = "B";
            }
            else if (MONTH_VALUE == "03")
            {
                GETID = "C";
            }
            else if (MONTH_VALUE == "04")
            {
                GETID = "D";
            }
            else if (MONTH_VALUE == "05")
            {
                GETID = "E";
            }
            else if (MONTH_VALUE == "06")
            {
                GETID = "F";
            }
            else if (MONTH_VALUE == "07")
            {
                GETID = "G";
            }
            else if (MONTH_VALUE == "08")
            {
                GETID = "H";
            }
            else if (MONTH_VALUE == "09")
            {
                GETID = "I";
            }
            else if (MONTH_VALUE == "10")
            {
                GETID = "J";
            }
            else if (MONTH_VALUE == "11")
            {
                GETID = "K";
            }
            else if (MONTH_VALUE == "12")
            {
                GETID = "L";
            }
        
            return GETID;
        }
        #endregion
        #region save
        public void save(DataGridView dgv,string OFFER_TYPE_CODE)
        {
          
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string GET_OFFER_ID = bc.getOnlyString("SELECT OFFER_ID FROM PRINTING_OFFER_MST WHERE  PFID='" +PFID  + "'");
            string v1 = bc.getOnlyString("SELECT CHARGE_AUDIT_STATUS FROM PRINTING_OFFER_MST WHERE PFID='" + PFID + "'");
            if (!bc.exists("SELECT PFID FROM PRINTING_OFFER_DET WHERE PFID='" + PFID  + "'"))
            {
                SQlcommandE_DET(dgv, sqlo);
                SQlcommandE_MST(sqlt);
                IFExecution_SUCCESS = true;
                OFFER_ID_SENVEN = numYY(8, 3, "001", "select * from PRINTING_OFFER_ID_NO", "OFFER_ID_SENVEN", OFFER_TYPE_CODE);
         
            }
            else if (bc.exists("SELECT PFID FROM PRINTING_OFFER_DET WHERE PFID='" + PFID + "'") && v1!="Y")
            {
                SQlcommandE_DET(dgv, sqlo);
                SQlcommandE_MST(sqlth + " WHERE PFID='" + PFID  + "'");
                IFExecution_SUCCESS = true;
            }
            else
            {
                SQlcommandE_DET(dgv, sqlo);
                SQlcommandE_MST(sqlt);
                IFExecution_SUCCESS = true;
            
            }
        }
        #endregion
        #region SQlcommandE_DET
        protected void SQlcommandE_DET(DataGridView dgv, string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
            basec.getcoms("DELETE PRINTING_OFFER_DET WHERE PFID='" + PFID + "'");
            for (i = 0; i < dgv.Rows.Count; i++)
            {
               
                if (dgv["部品名", i].FormattedValue.ToString() == "")
                {

                }
                else
                {
                    SqlConnection sqlcon = bc.getcon();
                    sqlcon.Open();
                    SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                    PFKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM PRINTING_OFFER_DET", "PFKEY", "PF");
                    sqlcom.Parameters.Add("@PFKEY", SqlDbType.VarChar, 20).Value = PFKEY;
                    sqlcom.Parameters.Add("@PFID", SqlDbType.VarChar, 20).Value = PFID;
                    sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = dgv["项次", i].Value.ToString();
                    sqlcom.Parameters.Add("@DET_WNAME", SqlDbType.VarChar, 50).Value = dgv["部品名", i].Value.ToString();
                    sqlcom.Parameters.Add("@DRAWING_DOOR ", SqlDbType.VarChar, 20).Value = dgv["图纸门幅", i].Value.ToString();
                    sqlcom.Parameters.Add("@PAPER_LENGTH ", SqlDbType.VarChar, 20).Value = dgv["图纸纸长", i].Value.ToString();
                    sqlcom.Parameters.Add("@UNIT_DOSAGE_M ", SqlDbType.VarChar, 20).Value = dgv["部品个数", i].Value.ToString();
                    sqlcom.Parameters.Add("@UNIT_DOSAGE_D ", SqlDbType.VarChar, 20).Value = dgv["拼模数", i].Value.ToString();
                    sqlcom.Parameters.Add("@PRINT_OPTION ", SqlDbType.VarChar, 20).Value = dgv["印刷选项", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@TISSUE_SPEC ", SqlDbType.VarChar, 20).Value = dgv["面纸", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@WEIGHT ", SqlDbType.VarChar, 20).Value = dgv["面纸克重", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@PAPER_CORE ", SqlDbType.VarChar, 20).Value = dgv["芯纸", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@SPEC ", SqlDbType.VarChar, 20).Value = dgv["芯纸规格", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@BODY_PAPER ", SqlDbType.VarChar, 20).Value = dgv["底纸", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@BODY_WEIGHT ", SqlDbType.VarChar, 20).Value = dgv["底纸克重", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@POSITIVE_4C ", SqlDbType.VarChar, 20).Value = dgv["正面4C", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@POSITIVE_COLOR ", SqlDbType.VarChar, 20).Value = dgv["正面专色", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@POSITIVE_SUN_SCREEN ", SqlDbType.VarChar, 20).Value = dgv["正面防晒", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@DOUBLE_PRINTING ", SqlDbType.VarChar, 20).Value = dgv["双面印刷", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@OPPOSITE_4C ", SqlDbType.VarChar, 20).Value = dgv["反面4C", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@OPPOSITE_COLOR ", SqlDbType.VarChar, 20).Value = dgv["反面专色", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@OPPOSITE_SUN_SCREEN ", SqlDbType.VarChar, 20).Value = dgv["反面防晒", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@SURFACE_PROCESSING ", SqlDbType.VarChar, 20).Value = dgv["表面加工", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@SURFACE_COUNT ", SqlDbType.VarChar, 20).Value = dgv["表面次数", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@LAMINATING_PROCESS ", SqlDbType.VarChar, 20).Value = dgv["裱纸工艺", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@LAMINATING_COUNT ", SqlDbType.VarChar, 20).Value = dgv["裱纸次数", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@DIE_CUTTING ", SqlDbType.VarChar, 22).Value = dgv["模切", i].FormattedValue.ToString();
                    sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = MAKERID;
                    sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
                    sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                    sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                    sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                    sqlcom.ExecuteNonQuery();
                    sqlcon.Close();
                }
               
            }
          
        }
        #endregion
        #region SQlcommandE_MST
        protected void SQlcommandE_MST(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcon.Open();
            sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
            sqlcom.Parameters.Add("PIID", SqlDbType.VarChar, 20).Value = PIID;
            sqlcom.Parameters.Add("OFFER_ID", SqlDbType.VarChar, 20).Value = OFFER_ID;
            sqlcom.Parameters.Add("CHARGE_AUDIT_STATUS", SqlDbType.VarChar, 20).Value = CHARGE_AUDIT_STATUS;
            sqlcom.Parameters.Add("COUNT", SqlDbType.VarChar, 20).Value = COUNT;
            sqlcom.Parameters.Add("OFFER_MAKERID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("OFFER_DATE", SqlDbType.VarChar, 20).Value = varDate.Substring(0, 10);
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@EDIT_TIME", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region RETURN_DT
        public DataTable RETURN_DT(DataTable dtt)
        {
            DataTable dt = GetTableInfo_show_all();
            foreach (DataRow dr1 in dtt.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["客户"] = dr1["客户"].ToString();
                dr["品牌"] = dr1["品牌"].ToString();
                dr["AE"] = dr1["AE"].ToString();
                dr["数量"] = dr1["数量"].ToString();
                dr["项目号"] = dr1["项目号"].ToString();
                dr["报价编号"] = dr1["报价编号"].ToString();
                dr["报价"] = dr1["报价"].ToString();
                dr["日期"] = dr1["日期"].ToString();
                dr["审核状态"] = dr1["审核状态"].ToString();
                dr["项次"] = dr1["项次"].ToString();
                dr["部品名"] = dr1["部品名"].ToString();
                dr["图纸门幅"] = dr1["图纸门幅"].ToString();
                dr["图纸纸长"] = dr1["图纸纸长"].ToString();
                if (!string.IsNullOrEmpty(dr1["部品个数"].ToString()))
                {
                    dr["部品个数"] = dr1["部品个数"].ToString();
                }
                else
                {
                    dr["部品个数"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["拼模数"].ToString()))
                {
                    dr["拼模数"] = dr1["拼模数"].ToString();
                }
                else
                {
                    dr["拼模数"] = DBNull.Value;
                }
                dr["印刷选项"] = dr1["印刷选项"].ToString();
                dr["面纸"] = dr1["面纸"].ToString();
                if (!string.IsNullOrEmpty(dr1["面纸克重"].ToString()))
                {
                    dr["面纸克重"] = dr1["面纸克重"].ToString();
                }
                else
                {
                    dr["面纸克重"] = DBNull.Value;
                }
                dr["芯纸"] = dr1["芯纸"].ToString();
                dr["芯纸规格"] = dr1["芯纸规格"].ToString();
                dr["底纸"] = dr1["底纸"].ToString();
                if (!string.IsNullOrEmpty(dr1["底纸克重"].ToString()))
                {
                    dr["底纸克重"] = dr1["底纸克重"].ToString();
                }
                else
                {
                    dr["底纸克重"] = DBNull.Value;
                }
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
            return dt;
        }
        #endregion
        #region bind2()
        public DataTable bind2(DataTable dt ,int d,string count)
        {
            int i1=0;
            string v1 = "", v2 = "", v3 = "", v5 = "", v8 = "";
            decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0, d5 = 0, d6 = 0, d7 = 0,d8=0,d9=0,d10=0,d11=0;
            i = 0;
            
            StringBuilder sqb = new StringBuilder();
           DataTable  dtx1 = bc.getdt(string.Format (@"
SELECT
B.CName AS 客户名称,
c.BRAND AS 品牌,
SUBSTRING(C.CUSTOMER_TYPE,1,1) 客户类别 
FROM  CUSTOMERINFO_MST B 
LEFT JOIN CustomerInfo_DET C ON B.CUID =C.CUID 
WHERE B.CNAME='{0}' AND C.BRAND='{1}'", dt.Rows[0]["客户"].ToString(), dt.Rows[0]["品牌"].ToString()));
           if (dtx1.Rows.Count > 0)
           {

               CUSTOMER_TYPE = dtx1.Rows[0]["客户类别"].ToString();

           }
         
            foreach (DataRow dr in dt.Rows)
            {
                
                if (CUSTOMER_TYPE == null)
                {
                    ErrowInfo = "本客户未设置客户类别 或 项目号没有设置品牌";
                    break;
                }
                //MessageBox.Show(dr["图纸门幅"].ToString());
           
                dtx2 = bc.getdt(cprint_option.sql + string.Format(" WHERE A.PRINT_OPTION='{0}'", dr["印刷选项"].ToString()));
                ErrowInfo = null;
                if(!string .IsNullOrEmpty(dr["表面次数"].ToString ()))
                {
                    SURFACE_NUMBER = decimal.Parse(dr["表面次数"].ToString());
                }
                else 
                {
                    SURFACE_NUMBER = 0;
                }
                if (d == 0)
                {
                    COUNT = decimal.Parse(dr["数量"].ToString());
                }
                else
                {
                    if (!string.IsNullOrEmpty(count) && bc.yesno(count)!=0)
                    {
                        COUNT = decimal.Parse(count);
                    }
                 
                }
                
                if (dr["图纸门幅"].ToString() == "" || dr["图纸纸长"].ToString() == "" || dr["印刷选项"].ToString() == "")
                {


                }
                else
                {
                    if (dtx2.Rows.Count > 0)
                    {
                        v1 = dtx2.Rows[0]["修边"].ToString();//修边放数
                        if (!string.IsNullOrEmpty(v1))
                        {
                            i1 = Convert.ToInt32(v1);
                        }

                    }
          
                
                }
             
                if (!string.IsNullOrEmpty(dr["裱纸次数"].ToString()))
                {
                    LAMINATING_NUMBER = decimal.Parse(dr["裱纸次数"].ToString());
                }
                else
                {
                    LAMINATING_NUMBER = 0;
                }
                if (!string.IsNullOrEmpty(dr["图纸门幅"].ToString()))
                {
                    d1 = decimal.Parse(dr["图纸门幅"].ToString());
                    dr["加工门幅"] = d1 + i1;
                }
                if (!string.IsNullOrEmpty(dr["加工门幅"].ToString()))
                {
                    PROCESSING_DOOR = decimal.Parse(dr["加工门幅"].ToString());
                 
                }
                else
                {
                    PROCESSING_DOOR = 0;
                }
                if (!string.IsNullOrEmpty(dr["图纸纸长"].ToString()))
                {
                    d1 = decimal.Parse(dr["图纸纸长"].ToString());
                    dr["加工长度"] = d1 + i1;
                }
                if (!string.IsNullOrEmpty(dr["加工长度"].ToString()))
                {
                    PROCESSING_LENGTH = decimal.Parse(dr["加工长度"].ToString());

                }
                else
                {
                    PROCESSING_LENGTH = 0;
                }

                if (!string.IsNullOrEmpty(dr["部品个数"].ToString()) && !string.IsNullOrEmpty(dr["拼模数"].ToString()))
                {
                    d1 = decimal.Parse(dr["部品个数"].ToString());
                    d2 = decimal.Parse(dr["拼模数"].ToString());
                    dr["部品总数"] = d1 / d2 * COUNT;

                }
                else
                {
                   // MessageBox.Show(dr["部品个数"].ToString() + "," + dr["拼模数"].ToString());
                }
            
                if (!string.IsNullOrEmpty(dr["部品总数"].ToString()))
                {
                    TOTAL_PRODUCT_NUMBER = decimal.Parse(dr["部品总数"].ToString());
                }
                else
                {
                    TOTAL_PRODUCT_NUMBER = 0;
                }
                DataTable dtt1 = bc.getdt(cprint_machine_size.sql + " WHERE A.MACHINE_TYPE='四开'");
                if (dtt1.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtt1.Rows[0]["最大宽"].ToString()))
                    {
                        d3 = decimal.Parse(dtt1.Rows[0]["最大宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtt1.Rows[0]["最大长"].ToString()))
                    {
                        d4 = decimal.Parse(dtt1.Rows[0]["最大长"].ToString());
                    }
                }
                DataTable dtt2 = bc.getdt(cprint_machine_size.sql + " WHERE A.MACHINE_TYPE='对开'");
                if (dtt2.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtt2.Rows[0]["最大宽"].ToString()))
                    {
                        d5 = decimal.Parse(dtt2.Rows[0]["最大宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtt2.Rows[0]["最大长"].ToString()))
                    {
                       d6 = decimal.Parse(dtt2.Rows[0]["最大长"].ToString());
                    }
                }
                DataTable dtt3 = bc.getdt(cprint_machine_size.sql + " WHERE A.MACHINE_TYPE='全开'");
                if (dtt3.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtt3.Rows[0]["最大宽"].ToString()))
                    {
                        d7 = decimal.Parse(dtt3.Rows[0]["最大宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtt3.Rows[0]["最大长"].ToString()))
                    {
                        d8 = decimal.Parse(dtt3.Rows[0]["最大长"].ToString());
                    }
                }
                DataTable dtt4 = bc.getdt(cprint_machine_size.sql + " WHERE A.MACHINE_TYPE='大全开'");
                if (dtt4.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtt4.Rows[0]["最大宽"].ToString()))
                    {
                        d9 = decimal.Parse(dtt4.Rows[0]["最大宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtt4.Rows[0]["最大长"].ToString()))
                    {
                        d10 = decimal.Parse(dtt4.Rows[0]["最大长"].ToString());
                    }
                }
                if (!string.IsNullOrEmpty(dr["部品总数"].ToString()) && !string.IsNullOrEmpty(dr["印刷选项"].ToString()))
                {
                    d1 = decimal.Parse(dr["加工门幅"].ToString());
                    d2 = decimal.Parse(dr["加工长度"].ToString());
                    if (Math.Min(d1, d2) <=d3 && Math.Max(d1, d2) <=d4)
                    {
                        dr["机器型号"] = "四开";
                    }
                    else if (Math.Min(d1, d2) <= d5 && Math.Max(d1, d2) <=d6)
                    {
                        dr["机器型号"] = "对开";
                    }
                    else if (Math.Min(d1, d2) <=d7 && Math.Max(d1, d2) <= d8)
                    {
                        dr["机器型号"] = "全开";
                    }
                    else if (Math.Min(d1, d2) <=d9 && Math.Max(d1, d2) <=d10)
                    {
                        dr["机器型号"] = "大全开";
                    }
                    else
                    {
                        dr["机器型号"] = "未发明";
                    }
                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.DIE_CUTTING/(1+B.TAX_RATE/100)
FROM
MACHINING_DET A
LEFT JOIN MACHINING_MST B ON A.MAID=B.MAID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                if (!string.IsNullOrEmpty(v1))
                {
                    DIE_CUTTING  = decimal.Parse(v1);
                }
                else
                {
                    DIE_CUTTING  = 0;
                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.MACHINE_FREE/(1+B.TAX_RATE/100)
FROM
MACHINING_DET A
LEFT JOIN MACHINING_MST B ON A.MAID=B.MAID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                if (!string.IsNullOrEmpty(v1))
                {
                    MACHINING_MACHINE_FREE  = decimal.Parse(v1);
                }
                else
                {
                    MACHINING_MACHINE_FREE  = 0;
                }
            
                if (!string.IsNullOrEmpty(dr["正面4C"].ToString()))
                {
                    POSITIVE_4C = decimal.Parse(dr["正面4C"].ToString());
                }
                else
                {
                    POSITIVE_4C = 0;
                }
                if (!string.IsNullOrEmpty(dr["反面4C"].ToString()))
                {
                    OPPOSITE_4C = decimal.Parse(dr["反面4C"].ToString());
                }
                else
                {
                    OPPOSITE_4C = 0;
                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.MACHINE_FREE/(1+B.TAX_RATE/100)
FROM
MACHINING_DET A
LEFT JOIN MACHINING_MST B ON A.MAID=B.MAID
WHERE B.MACHINE_TYPE='{0}'  AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),
                                                                                            bc.RETURN_CUSTOMER_TYPE (dr["项目号"].ToString ())));
                if (!string.IsNullOrEmpty(v1))
                {
                    SQUARE_OR_METRE_MIN = decimal.Parse(v1);
                }
                else
                {
                    SQUARE_OR_METRE_MIN  = 0;
                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.DIE_CUTTING/(1+B.TAX_RATE/100)
FROM
MACHINING_DET A
LEFT JOIN MACHINING_MST B ON A.MAID=B.MAID
WHERE B.MACHINE_TYPE='{0}'  AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", 
 dr["机器型号"].ToString(),bc.RETURN_CUSTOMER_TYPE (dr["项目号"].ToString ())));
                if (!string.IsNullOrEmpty(v1))
                {
                    DIE_CUTTING_PRICE = decimal.Parse(v1);
                }
                else
                {
                    DIE_CUTTING_PRICE = 0;
                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.MIN_PRINTING
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}' ", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                if (!string.IsNullOrEmpty(v1))
                {
                    MIN_PRINTING = decimal.Parse(v1);//起印数
                }
                else
                {
                    MIN_PRINTING = 0;//起印数
                }
                if (dr["机器型号"].ToString() != "")
                {
                    v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.CTP_EDITION/(1+B.TAX_RATE/100)
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                    dr["CTP单张价"] = v1;

                }
                if (!string.IsNullOrEmpty(v1))
                {
                    CTP_UNIT_PRICE = decimal.Parse(v1);
                }
                else
                {
                    CTP_UNIT_PRICE = 0;

                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.OUT_OF_PRINT/(1+B.TAX_RATE/100)
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                if (!string.IsNullOrEmpty(v1))
                {
                    OUT_OF_PRINT= decimal.Parse(v1);
                }
                else
                {
                    OUT_OF_PRINT= 0;
                }
                if (dr["机器型号"].ToString() != "")
                {
                    dr["超出单色单张价"] = v1;
                }
                if (!string.IsNullOrEmpty(dr["超出单色单张价"].ToString()))
                {
                    PASS_COLOR_UNIT_PRICE = decimal.Parse(dr["超出单色单张价"].ToString());
                }
                else
                {
                    PASS_COLOR_UNIT_PRICE = 0;
                }
                if (!string.IsNullOrEmpty(dr["面纸"].ToString()) && !string.IsNullOrEmpty(dr["面纸克重"].ToString()))
                {
                    v1 = bc.getOnlyString(string.Format(@"
SELECT
A.TON_PRICE
FROM TISSUE_SPEC_DET A
LEFT JOIN TISSUE_SPEC_MST B ON A.TSID=B.TSID 
WHERE
B.TISSUE_SPEC='{0}' AND A.WEIGHT='{1}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{2}'", dr["面纸"].ToString(), dr["面纸克重"].ToString(),CUSTOMER_TYPE ));
                    if (!string.IsNullOrEmpty(v1))
                    {
                        d1 = decimal.Parse(v1);
                        d2 = decimal.Parse(dr["面纸克重"].ToString());
                        dr["面纸单价"] = (d1 * d2 / 1000000).ToString();
                    }

                }
                d2 = 0;
                d3 = 0;
                if (dtx2.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["面纸内耗1到300"].ToString()))
                    {
                        d2 = decimal.Parse(dtx2.Rows[0]["面纸内耗1到300"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["面纸内耗大于300"].ToString()))
                    {
                        d3 = decimal.Parse(dtx2.Rows[0]["面纸内耗大于300"].ToString());
                    }
                }
                if (!string.IsNullOrEmpty(dr["部品总数"].ToString()) && !string.IsNullOrEmpty(dr["印刷选项"].ToString()))
                {
                    d1 = decimal.Parse(dr["部品总数"].ToString());
                    if (d1 <= 300)
                    {
                        dr["面纸内耗"] = d2;
                    }
                    else
                    {
                        dr["面纸内耗"] = d2 + (decimal.Parse(dr["部品总数"].ToString()) - 300) * d3 * 1 / 100;//面纸内耗
                    }
                    
                }
                if (!string.IsNullOrEmpty(dr["面纸内耗"].ToString()))
                {
                    TISSUE_INSIDE_LOSE = decimal.Parse(dr["面纸内耗"].ToString());
                }
                else
                {
                    TISSUE_INSIDE_LOSE = 0;
                }
                if (!string.IsNullOrEmpty(dr["部品总数"].ToString()) && !string.IsNullOrEmpty(dr["印刷选项"].ToString()))
                {

                    dr["面纸下单"]=Convert.ToDouble(dr["部品总数"].ToString()) + Convert.ToDouble(dr["面纸内耗"].ToString());

                }
                if (!string.IsNullOrEmpty(dr["面纸下单"].ToString()))
                {
                    TISSUE_ORDER = decimal.Parse(dr["面纸下单"].ToString());
                }
                else
                {
                    TISSUE_ORDER = 0;
                }
                NEED_COUNT = TISSUE_ORDER * 2;
           
                if (!string.IsNullOrEmpty(dr["机器型号"].ToString()) && !string.IsNullOrEmpty(dr["表面加工"].ToString()) && 
                    !string.IsNullOrEmpty(dr["表面次数"].ToString()))
                {


                    dr["表面处理单价"] = bc.getOnlyString(string.Format(@"
SELECT 
A.SURFACE_PROCESSING_PRICE/(1+B.TAX_RATE/100)
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND A.SURFACE_PROCESSING='{1}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{2}'", dr["机器型号"].ToString(), 
                                                                                                   dr["表面加工"].ToString(),CUSTOMER_TYPE ));
                    
                }

                if (!string.IsNullOrEmpty(dr["表面处理单价"].ToString()) && !string.IsNullOrEmpty(dr["加工门幅"].ToString()) &&
                    !string.IsNullOrEmpty(dr["加工长度"].ToString()))
                {
                    d1 = decimal.Parse(dr["加工门幅"].ToString());
                    d2 = decimal.Parse(dr["加工长度"].ToString());
                    
                    dr["表面处理用量"] = d1 / 1000 * d2 / 1000 * decimal.Parse(dr["表面次数"].ToString());
                }
                if (!string.IsNullOrEmpty(dr["表面处理单价"].ToString()) && !string.IsNullOrEmpty(dr["表面处理用量"].ToString()) &&
                     !string.IsNullOrEmpty(dr["部品总数"].ToString()))
                {
                    d1 = decimal.Parse(dr["加工门幅"].ToString());
                    d2 = decimal.Parse(dr["加工长度"].ToString());

                   
                }
                if (!string.IsNullOrEmpty(dr["表面处理单价"].ToString()) && !string.IsNullOrEmpty(dr["表面处理用量"].ToString()) &&
                !string.IsNullOrEmpty(dr["部品总数"].ToString()))
                {


                    v1= bc.getOnlyString(string.Format(@"
SELECT 
B.MACHINE_FREE/(1+B.TAX_RATE/100)
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND A.SURFACE_PROCESSING='{1}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{2}'", 
                                                     dr["机器型号"].ToString(), dr["表面加工"].ToString(),CUSTOMER_TYPE ));
                    if (!string.IsNullOrEmpty(v1))
                    {
                        d1 = decimal.Parse(v1);
                    }
                    else
                    {
                        d1 = 0;
                    }
                    d2=decimal .Parse (dr["表面处理用量"].ToString());
                    d3=decimal .Parse (dr["部品总数"].ToString());
                    d4 = decimal.Parse(dr["表面处理单价"].ToString());
                    //MessageBox.Show(d1.ToString()+","+d2.ToString() + "," + d3.ToString() + "," + d4.ToString()+","+(d2*d3*d4).ToString());
                    dr["表面加工小计"] = Math.Max(d1 * decimal.Parse(dr["表面次数"].ToString()), d2 *d3*d4 );

                }
                if (!string.IsNullOrEmpty(dr["表面加工小计"].ToString()))
                {
                    TOTAL_SURFACE_PROCESSING = decimal.Parse(dr["表面加工小计"].ToString());
                }
                else
                {
                    TOTAL_SURFACE_PROCESSING = 0;
                }
                if (dr["机器型号"].ToString()=="" && dr["正面4C"].ToString()=="" && dr["正面专色"].ToString()=="")
                {
        
                }
                else
                {
                    d1 = 0;
                    d2 = 0;
                    if(!string.IsNullOrEmpty (dr["正面4C"].ToString ()))
                    {
                        d1 = decimal.Parse(dr["正面4C"].ToString());
                    }
                    v1 = bc.getOnlyString("SELECT COLOR_COUNT FROM COLOR_PARAMETERS WHERE COLOR_PARAMETERS='" + dr["正面专色"].ToString() + "'");
                    if (!string.IsNullOrEmpty(v1))
                    {
                        d2 = decimal.Parse(v1);
                    }
                    if (d1 + d2 > 0)
                    {

                        dr["正面色数共计"] = d1 + d2;
                    }
                  
                }
                if (!string.IsNullOrEmpty(dr["正面色数共计"].ToString()))
                {
                    POSITIVE_COLOR_COUNT_TOTAL = decimal.Parse(dr["正面色数共计"].ToString());
                }
                else
                {
                    POSITIVE_COLOR_COUNT_TOTAL = 0;
                }
                if (dr["正面色数共计"].ToString()=="")
                {

                }
                else 
                {
                    if (dr["正面专色"].ToString() == "")
                    {
                        dr["正面CTP张数"] = dr["正面4C"].ToString();
                    }
                    else
                    {
                        v1 = bc.getOnlyString(string.Format(@"
SELECT 
CTP_EDITION
FROM
COLOR_PARAMETERS
WHERE COLOR_PARAMETERS='{0}' ", dr["正面专色"].ToString()));
                        if (!string.IsNullOrEmpty(v1))
                        {
                            dr["正面CTP张数"] =POSITIVE_4C + decimal.Parse(v1);
                        }
                        else
                        {
                            dr["正面CTP张数"] = dr["正面4C"].ToString();
                        }

                    }

                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
CTP_EDITION
FROM
COLOR_PARAMETERS
WHERE COLOR_PARAMETERS='{0}' ", dr["反面专色"].ToString()));
                if (!string.IsNullOrEmpty(v1))
                {
                    POSITIVE_CTP_EDITION_FOR_PARAMETERS = decimal.Parse(v1);
                }
                else
                {
                    POSITIVE_CTP_EDITION_FOR_PARAMETERS = 0;
                }
                if (!string.IsNullOrEmpty(dr["正面CTP张数"].ToString()))
                {
                    POSITIVE_CTP_COUNT = decimal.Parse(dr["正面CTP张数"].ToString());
                }
                else
                {
                    POSITIVE_CTP_COUNT = 0;
                }
                if (!string.IsNullOrEmpty(dr["CTP单张价"].ToString()) && !string.IsNullOrEmpty(dr["正面CTP张数"].ToString()))
                {
                    dr["正面CTP价计"] = CTP_UNIT_PRICE * POSITIVE_CTP_COUNT;
                }
                if (!string.IsNullOrEmpty(dr["正面CTP价计"].ToString()))
                {
                    POSITIVE_CTP_PRICE_TOTAL = decimal.Parse(dr["正面CTP价计"].ToString());
                }
                else
                {
                    POSITIVE_CTP_PRICE_TOTAL = 0;
                }
                if (dr["机器型号"].ToString() != "")
                {
                    v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.MONOCHROME_PRINTING/(1+B.TAX_RATE/100)
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                    dr["印工单色单价"] = v1;
               
                  
                }
                if (!string.IsNullOrEmpty(v1))
                {
                    PRINTING_UNIT_PRICE = decimal.Parse(dr["印工单色单价"].ToString());
                }
                else
                {
                    PRINTING_UNIT_PRICE = 0;
                }
                d1 = 0;
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
B.SUN_SCREEN_INK/(1+B.TAX_RATE/100)
FROM
PRINTING_TYPE_DET A
LEFT JOIN PRINTING_TYPE_MST B ON A.PTID=B.PTID
WHERE B.MACHINE_TYPE='{0}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{1}'", dr["机器型号"].ToString(),CUSTOMER_TYPE ));
                if (!string.IsNullOrEmpty(v1))
                {
                    
                    SUN_SCREEN_INK =decimal .Parse (v1);
                }
                else 
                {
                    SUN_SCREEN_INK =0;
                }
              
                if (dr["印刷选项"].ToString() == "" || dr["印刷选项"].ToString() == "不印刷" || dr["正面色数共计"].ToString() == "")
                {

                }
                else if (dr["正面防晒"].ToString() == "是" && (dr["印刷选项"].ToString() == "双纸画同" || 
                    dr["印刷选项"].ToString() == "单纸双同"))
                {
                   
                    dr["正面防晒合计"] = 2 * decimal.Parse(dr["面纸下单"].ToString()) * SUN_SCREEN_INK ;
                }
                else if (dr["正面防晒"].ToString() == "是")
                {
                    
                    dr["正面防晒合计"] =TISSUE_ORDER  * SUN_SCREEN_INK ;
                }
                else
                {
                    dr["正面防晒合计"] = "0";

                }
              
                if(!string .IsNullOrEmpty (dr["正面防晒合计"].ToString ()))
                {
                    POSITIVE_SUN_SCREEN_TOTAL = decimal.Parse(dr["正面防晒合计"].ToString());
                }
                else 
                {
                    POSITIVE_SUN_SCREEN_TOTAL = 0;
                }

                if (dr["机器型号"].ToString() == "" || dr["正面色数共计"].ToString() == "" || dr["印刷选项"].ToString() == "" ||
                    dr["印刷选项"].ToString() == "不印刷")
                {
                  
                }
                else if ((dr["印刷选项"].ToString() == "单纸双同" || dr["印刷选项"].ToString() == "双纸画同" ) &&
                    POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER*2<=MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * 4 + POSITIVE_SUN_SCREEN_TOTAL;
                    
                }
                else if ((dr["印刷选项"].ToString() == "单纸双同" || dr["印刷选项"].ToString() == "双纸画同") &&
                    POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER * 2 >MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * 4 + POSITIVE_SUN_SCREEN_TOTAL + (TISSUE_ORDER * 2 - MIN_PRINTING) * 
                        POSITIVE_COLOR_COUNT_TOTAL * PASS_COLOR_UNIT_PRICE;

                }
                else if ((dr["印刷选项"].ToString() == "单纸双同" || dr["印刷选项"].ToString() == "双纸画同") &&
                    POSITIVE_COLOR_COUNT_TOTAL> 4 && TISSUE_ORDER * 2 <= MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * POSITIVE_COLOR_COUNT_TOTAL + POSITIVE_SUN_SCREEN_TOTAL;

                }
                else if ((dr["印刷选项"].ToString() == "单纸双同" || dr["印刷选项"].ToString() == "双纸画同") &&
                    POSITIVE_COLOR_COUNT_TOTAL > 4 && TISSUE_ORDER * 2 > MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * POSITIVE_COLOR_COUNT_TOTAL + POSITIVE_SUN_SCREEN_TOTAL + (TISSUE_ORDER * 2 - MIN_PRINTING) *
                        POSITIVE_COLOR_COUNT_TOTAL * PASS_COLOR_UNIT_PRICE;

                }

                else if (POSITIVE_COLOR_COUNT_TOTAL <=4 && TISSUE_ORDER< MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * 4 + POSITIVE_SUN_SCREEN_TOTAL;

                }
                else if (POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER >MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * 4 + POSITIVE_SUN_SCREEN_TOTAL + (TISSUE_ORDER - MIN_PRINTING) *
                        POSITIVE_COLOR_COUNT_TOTAL * PASS_COLOR_UNIT_PRICE;
                

                                        
                }
                else if (POSITIVE_COLOR_COUNT_TOTAL > 4 && TISSUE_ORDER <= MIN_PRINTING )
                {

                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * POSITIVE_COLOR_COUNT_TOTAL + POSITIVE_SUN_SCREEN_TOTAL;
    
                }
                else
                {
                    dr["正面印工合计"] = PRINTING_UNIT_PRICE * POSITIVE_COLOR_COUNT_TOTAL + POSITIVE_SUN_SCREEN_TOTAL + (TISSUE_ORDER  - MIN_PRINTING) *
                     POSITIVE_COLOR_COUNT_TOTAL * PASS_COLOR_UNIT_PRICE;

               /*sqb = new StringBuilder();
 sqb.AppendFormat("部品名：{0},",dr["部品名"].ToString());
 sqb.AppendFormat("印工单色单价：{0},",PRINTING_UNIT_PRICE );
 sqb.AppendFormat("正面防晒合计：{0},", POSITIVE_SUN_SCREEN_TOTAL);
 sqb.AppendFormat("面纸下单：{0},", TISSUE_ORDER);
 sqb.AppendFormat("起印数：{0},", MIN_PRINTING);
 sqb.AppendFormat("正面色数共计：{0},", POSITIVE_COLOR_COUNT_TOTAL);
 sqb.AppendFormat("超出单色单张价：{0},", PASS_COLOR_UNIT_PRICE);
 sqb.AppendFormat("正面印工合计：{0},", PRINTING_UNIT_PRICE * POSITIVE_COLOR_COUNT_TOTAL + POSITIVE_SUN_SCREEN_TOTAL + (TISSUE_ORDER - MIN_PRINTING) *
                     POSITIVE_COLOR_COUNT_TOTAL * PASS_COLOR_UNIT_PRICE);
 MessageBox.Show(sqb.ToString());*/
                }
                if(!string .IsNullOrEmpty(dr["正面印工合计"].ToString ()))
                {
                    POSITIVE_PRINTING_TOTAL = decimal.Parse(dr["正面印工合计"].ToString());
                }
                else 
                {
                    POSITIVE_PRINTING_TOTAL = 0;
                }
                d2 = 0;
                d3 = 0;
                d4 = 0;
                d5 = 0;
                d6=0;
                d7=0;
                d8=0;
                d9=0;
                d10=0;
                d11=0;
                if (dtx2.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_A"].ToString()))
                    {
                        d2 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_A"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_B"].ToString()))
                    {
                        d3 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_B"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_C"].ToString()))
                    {
                        d4 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_C"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_D"].ToString()))
                    {
                        d5 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_D"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_E"].ToString()))
                    {
                        d6 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_E"].ToString());
                    }

                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_F"].ToString()))
                    {
                        d7 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_F"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_G"].ToString()))
                    {
                        d8 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_G"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_H"].ToString()))
                    {
                        d9 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_H"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_I"].ToString()))
                    {
                        d10 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_I"].ToString());
                    }
                     if (!string.IsNullOrEmpty(dtx2.Rows[0]["正面印刷纸张损耗_J"].ToString()))
                    {
                        d11 = decimal.Parse(dtx2.Rows[0]["正面印刷纸张损耗_J"].ToString());
                    }
                }
                if (dr["面纸下单"].ToString() == "" || dr["印刷选项"].ToString() == "不印刷" || dr["正面色数共计"].ToString() == "")
                {
                 
                }
                else if ((dr["印刷选项"].ToString() == "双纸画同" || dr["印刷选项"].ToString() == "单纸双同") &&
                    POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER*2<= 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {

                    dr["正面纸张损耗"] = d2 + POSITIVE_COLOR_COUNT_TOTAL * d3;
                }

                else if ((dr["印刷选项"].ToString() == "双纸画同" || dr["印刷选项"].ToString() == "单纸双同") &&
                    POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER * 2 >3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {

                    dr["正面纸张损耗"] = d4 + POSITIVE_COLOR_COUNT_TOTAL * d5 + (TISSUE_ORDER * 2 - 3000) * d6 / 100;
                }

                else if ((dr["印刷选项"].ToString() == "双纸画同" || dr["印刷选项"].ToString() == "单纸双同") &&
                    POSITIVE_COLOR_COUNT_TOTAL > 4 && TISSUE_ORDER * 2<= 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {

                    dr["正面纸张损耗"] = d7 + POSITIVE_COLOR_COUNT_TOTAL * d8;
                }
                else if ((dr["印刷选项"].ToString() == "双纸画同" || dr["印刷选项"].ToString() == "单纸双同") &&
                    POSITIVE_COLOR_COUNT_TOTAL > 4 && TISSUE_ORDER * 2 > 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {

                    dr["正面纸张损耗"] = d9 + POSITIVE_COLOR_COUNT_TOTAL * d10 + (TISSUE_ORDER * 2 - 3000) * d11 / 100;
                }
                else if (POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER <= 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {

                    dr["正面纸张损耗"] = d2 + POSITIVE_COLOR_COUNT_TOTAL * d3;
                }
                else if (POSITIVE_COLOR_COUNT_TOTAL <= 4 && TISSUE_ORDER > 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {
                    dr["正面纸张损耗"] = d4 + POSITIVE_COLOR_COUNT_TOTAL * d5 + (TISSUE_ORDER  - 3000) * d6 / 100;
                }
                else if (POSITIVE_COLOR_COUNT_TOTAL> 4 && TISSUE_ORDER <= 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {
                    dr["正面纸张损耗"] = d7 + POSITIVE_COLOR_COUNT_TOTAL * d8;
                }
                else if (POSITIVE_COLOR_COUNT_TOTAL > 4 && TISSUE_ORDER > 3000 && POSITIVE_COLOR_COUNT_TOTAL > 0)
                {
                    dr["正面纸张损耗"] = d9 + POSITIVE_COLOR_COUNT_TOTAL * d10 + (TISSUE_ORDER - 3000) * d11 / 100;
                }

                if(!string .IsNullOrEmpty (dr["正面纸张损耗"].ToString ()))
                {
                    POSITIVE_THE_PAPER_LOSE =decimal .Parse (dr["正面纸张损耗"].ToString ());
                }
                else 
                {
                    POSITIVE_THE_PAPER_LOSE =0;
                }
                v1 = bc.getOnlyString(string.Format(@"
SELECT 
COLOR_COUNT
FROM
COLOR_PARAMETERS
WHERE COLOR_PARAMETERS='{0}' ", dr["反面专色"].ToString()));
                if (!string.IsNullOrEmpty(v1))
                {
                    COLOR_COUNT = decimal.Parse(v1);
                }
                else
                {
                    COLOR_COUNT = 0;
                }
                if (dr["面纸下单"].ToString() == "" || dr["双面印刷"].ToString ()=="" || (string.IsNullOrEmpty (dr["反面4C"].ToString ()) &&
                    dr["反面专色"].ToString()==""))
                {
                }
                else if (dr["印刷选项"].ToString() == "双纸画异" || dr["印刷选项"].ToString() =="单纸双异")
                {
                    
                    dr["反面色数共计"] = OPPOSITE_4C  + COLOR_COUNT ;
                }
                if (!string.IsNullOrEmpty(dr["反面色数共计"].ToString()))
                {
                    OPPOSITE_COLOR_TOTAL = decimal.Parse(dr["反面色数共计"].ToString());
                }
                else
                {
                    OPPOSITE_COLOR_TOTAL = 0;

                }
                if (!string.IsNullOrEmpty (dr["反面色数共计"].ToString()))
                {
                    dr["反面CTP张数"] = OPPOSITE_4C + POSITIVE_CTP_EDITION_FOR_PARAMETERS;
                }
                d2 = 0;
                d3 = 0;
                d4 = 0;
                d5 = 0;
                d6 = 0;
                d7 = 0;
                d8 = 0;
                d9 = 0;
                d10 = 0;
                d11 = 0;
                if (dtx2.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_A"].ToString()))
                    {
                        d2 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_A"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_B"].ToString()))
                    {
                        d3 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_B"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_C"].ToString()))
                    {
                        d4 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_C"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_D"].ToString()))
                    {
                        d5 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_D"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_E"].ToString()))
                    {
                        d6 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_E"].ToString());
                    }

                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_F"].ToString()))
                    {
                        d7 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_F"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_G"].ToString()))
                    {
                        d8 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_G"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_H"].ToString()))
                    {
                        d9 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_H"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_I"].ToString()))
                    {
                        d10 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_I"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["反面印刷纸张损耗_J"].ToString()))
                    {
                        d11 = decimal.Parse(dtx2.Rows[0]["反面印刷纸张损耗_J"].ToString());
                    }
                }
                if (!string.IsNullOrEmpty(dr["面纸下单"].ToString()) && dr["反面色数共计"].ToString() != "" && dr["反面色数共计"].ToString() != "0")
                {   //AS27 面纸下单, BG27 反面色数共计,
                    if (decimal.Parse(dr["反面色数共计"].ToString()) <= 4 && decimal.Parse(dr["面纸下单"].ToString()) <= 3000)
                    {
                        dr["反面纸张损耗"] = d2 + decimal.Parse(dr["反面色数共计"].ToString()) * d3;
                    }
                    else if (decimal.Parse(dr["反面色数共计"].ToString()) <= 4 && decimal.Parse(dr["面纸下单"].ToString()) > 3000)
                    {
                        dr["反面纸张损耗"] = d4 + decimal.Parse(dr["反面色数共计"].ToString()) * d5 + (TISSUE_ORDER - 3000) * d6 / 100;
                    }
                    else if (decimal.Parse(dr["反面色数共计"].ToString()) > 4 && decimal.Parse(dr["面纸下单"].ToString()) <= 3000)
                    {
                        dr["反面纸张损耗"] = d7 + decimal.Parse(dr["反面色数共计"].ToString()) * d8;
                    }
                    else
                    {
                        dr["反面纸张损耗"] = d9 + decimal.Parse(dr["反面色数共计"].ToString()) * d10 + (TISSUE_ORDER - 3000) * d11 / 100;
                    }
                }
                if(!string.IsNullOrEmpty (dr["反面纸张损耗"].ToString ()))
                {
                    OPPOSITE_THE_PAPER_LOSE =decimal .Parse (dr["反面纸张损耗"].ToString ());
                }
                else 
                {
                    OPPOSITE_THE_PAPER_LOSE =0;
                }
                if (dr["反面CTP张数"].ToString() != "")
                {
                    if (dr["反面防晒"].ToString() == "是")
                    {
                        dr["反面防晒合计"] = TISSUE_ORDER * SUN_SCREEN_INK;
                    }
                    else
                    {
                        dr["反面防晒合计"] = 0;
                    }
                }
                if (!string.IsNullOrEmpty(dr["反面防晒合计"].ToString()))
                {
                    OPPOSITE_SUN_SCREEN_TOTAL = decimal.Parse(dr["反面防晒合计"].ToString());
                }
                else
                {
                    OPPOSITE_SUN_SCREEN_TOTAL = 0;
                }
                if (dr["CTP单张价"].ToString() != "" && dr["反面CTP张数"].ToString() != "")
                {
                    
                     dr["反面CTP价计"] = decimal.Parse(dr["CTP单张价"].ToString()) * decimal.Parse(dr["反面CTP张数"].ToString());
                }
                if (!string.IsNullOrEmpty(dr["反面CTP价计"].ToString()))
                {
                    OPPOSITE_CTP_PRICE_TOTAL = decimal.Parse(dr["反面CTP价计"].ToString());
                }
                else
                {
                    OPPOSITE_CTP_PRICE_TOTAL = 0;
                }
                if (dr["反面色数共计"].ToString() != "" && dr["反面色数共计"].ToString() != "0")
                {
                    if (decimal.Parse(dr["反面色数共计"].ToString()) <= 4 && TISSUE_ORDER <=MIN_PRINTING )
                    {
                        dr["反面印工合计"] = decimal.Parse(dr["印工单色单价"].ToString()) * 4 + decimal.Parse(dr["反面防晒合计"].ToString()); 
                    }
                    else if (decimal.Parse(dr["反面色数共计"].ToString()) <= 4 && TISSUE_ORDER > MIN_PRINTING)
                    {

                        dr["反面印工合计"] = PRINTING_UNIT_PRICE * 4 +
                            OPPOSITE_SUN_SCREEN_TOTAL  + (TISSUE_ORDER - MIN_PRINTING) *
                             PASS_COLOR_UNIT_PRICE * OPPOSITE_COLOR_TOTAL ;
                     
                    }
                    else if (TISSUE_ORDER <= MIN_PRINTING)
                    {
                        dr["反面印工合计"] = PRINTING_UNIT_PRICE * OPPOSITE_COLOR_TOTAL  +
                            OPPOSITE_SUN_SCREEN_TOTAL ;

                    }
                    else
                    {
                        dr["反面印工合计"] = PRINTING_UNIT_PRICE *OPPOSITE_COLOR_TOTAL  +
                       OPPOSITE_SUN_SCREEN_TOTAL  + (TISSUE_ORDER - MIN_PRINTING) *
                             PASS_COLOR_UNIT_PRICE * OPPOSITE_COLOR_TOTAL ;
                        /*StringBuilder sqb = new StringBuilder();
                        sqb.Append(dr["部品名"].ToString() + ",");
                        sqb.Append(dr["印工单色单价"].ToString() + ",");
                        sqb.Append(dr["反面色数共计"].ToString() + ",");
                        sqb.Append(dr["反面防晒合计"].ToString() + ",");
                        sqb.Append(TISSUE_ORDER + ",");
                        sqb.Append(MIN_PRINTING + ",");
                        sqb.Append(dr["超出单色单张价"].ToString() + ",");
                        MessageBox.Show(sqb.ToString ());*/
                    }
                }
                if (!string.IsNullOrEmpty(dr["反面印工合计"].ToString()))
                {
                    OPPOSITE_PRINTING_TOTAL = decimal.Parse(dr["反面印工合计"].ToString());
                }
                else
                {
                    OPPOSITE_PRINTING_TOTAL  = 0;
                }
                if (dr["正面CTP价计"].ToString() == "" && dr["反面CTP价计"].ToString() == "")
                {

                }
                else
                {
                    dr["正反CTP合计"] = POSITIVE_CTP_PRICE_TOTAL + OPPOSITE_CTP_PRICE_TOTAL;

                }
                if (dr["正面印工合计"].ToString() == "" && dr["反面印工合计"].ToString() == "")
                {

                }
                else
                {
                    dr["正反印工合计"] = POSITIVE_PRINTING_TOTAL + OPPOSITE_PRINTING_TOTAL;

                }
                if (!string.IsNullOrEmpty(dr["正反印工合计"].ToString()))
                {
                    TOTAL_POSITIVE_AND_OPPOSITE_PRINTING = decimal.Parse(dr["正反印工合计"].ToString());
                }
                else
                {
                    TOTAL_POSITIVE_AND_OPPOSITE_PRINTING = 0;
                }
                d2 = 0;
                d3 = 0;
                if (dtx2.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["无印刷用纸表面处理损耗_固定值"].ToString()))
                    {
                        d2 = decimal.Parse(dtx2.Rows[0]["无印刷用纸表面处理损耗_固定值"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["无印刷用纸表面处理损耗_百分比"].ToString()))
                    {
                        d3 = decimal.Parse(dtx2.Rows[0]["无印刷用纸表面处理损耗_百分比"].ToString());
                    }
                }
                if (dr["表面处理用量"].ToString() != "" && dr["正反印工合计"].ToString() == "")
                {
                    dr["无印刷表面处理损耗"] = SURFACE_NUMBER * (d2 + TISSUE_ORDER * d3 * 1 / 100);
                }
                else
                {
                    dr["无印刷表面处理损耗"] = 0;
                }
                if (dr["裱纸工艺"].ToString() == "" || dr["裱纸次数"].ToString() == "" || dr["机器型号"].ToString() == "")
                {

                }
                else
                {
                    v1 = bc.getOnlyString(string.Format(@"
SELECT 
LAMINATING_PROCESS_PRICE/(1+B.TAX_RATE/100)
FROM
MACHINING_DET A
LEFT JOIN MACHINING_MST B ON A.MAID=B.MAID
WHERE B.MACHINE_TYPE='{0}' AND A.LAMINATING_PROCESS='{1}' AND
SUBSTRING(B.CUSTOMER_TYPE,1,1)='{2}'", dr["机器型号"].ToString(), dr["裱纸工艺"].ToString(),CUSTOMER_TYPE ));
                    dr["裱工单价"] = v1;

                }
               
                if (!string.IsNullOrEmpty(dr["裱工单价"].ToString()))
                {
                    LAMINATING_PROCESS_PRICE = decimal.Parse(dr["裱工单价"].ToString());
                }
                else
                {
                    LAMINATING_PROCESS_PRICE = 0;
                }
                if (dr["裱工单价"].ToString() == "" || dr["加工门幅"].ToString() == "" || dr["加工长度"].ToString() == "")
                {

                }
                else
                {
                    dr["裱工用量"] = PROCESSING_DOOR / 1000 * PROCESSING_LENGTH / 1000 *LAMINATING_NUMBER ;
                }
                if (!string.IsNullOrEmpty(dr["裱工用量"].ToString()))
                {
                    LAMINATING_PROCESS_DOSAGE = decimal.Parse(dr["裱工用量"].ToString());
                }
                else
                {
                    LAMINATING_PROCESS_DOSAGE = 0;
                }
                if (dr["裱工单价"].ToString() == "" || dr["裱工用量"].ToString() == "" || dr["部品总数"].ToString() == "")
                {

                }
                else
                {
                    // MACHINING_MIN_PRINTING 起印数 ,LAMINATING_NUMBER 裱纸次数,LAMINATING_PROCESS_PRICE 裱工单价,
                    //LAMINATING_PROCESS_DOSAGE  裱工用量,TOTAL_PRODUCT_NUMBER 部品数
                    // 错的公式：max(起印数*裱纸次数,裱工单价*裱工用量*部品数);
                    // 对的公式：MAX(起机费，加工门幅/1000*加工纸长/1000*部品数*裱纸次数*裱工单价);
                    //MessageBox.Show(MACHINING_MIN_PRINTING.ToString() + "," + LAMINATING_NUMBER.ToString() + ",");
                    //MessageBox.Show(LAMINATING_PROCESS_PRICE + "," + LAMINATING_PROCESS_DOSAGE.ToString()+","+TOTAL_PRODUCT_NUMBER.ToString ());
                    dr["裱工小计"] = Math.Max(MACHINING_MACHINE_FREE * LAMINATING_NUMBER, LAMINATING_PROCESS_PRICE * 
                        LAMINATING_PROCESS_DOSAGE * TOTAL_PRODUCT_NUMBER);

                }
                if (!string.IsNullOrEmpty(dr["裱工小计"].ToString()))
                {
                    TOTAL_LAMINATING_PROCESS  = decimal.Parse(dr["裱工小计"].ToString());
                }
                else
                {
                    TOTAL_LAMINATING_PROCESS  = 0;
                }
                string sqln1=@"SELECT 
A.DIE_CUTTING AS 刀模,
A.TAX_UNIT_PRICE/(1+A.TAX_RATE/100) AS 未税单价,
CASE WHEN A.TAX_MACHINE_COST IS NOT NULL THEN A.TAX_MACHINE_COST/(1+A.TAX_RATE/100)      
ELSE ''
END AS 未税起机费,
A.TAX_UNIT_PRICE AS 含税单价,
A.UNIT AS 单位,
A.TAX_MACHINE_COST AS 含税起机费,
A.TAX_RATE AS 税率,
A.REMARK AS 说明,
B.ENAME AS 制单人,
A.DATE AS 制单日期
FROM DIE_CUTTING_COST A
LEFT JOIN EMPLOYEEINFO B ON A.MAKERID=B.EMID";
                d1 = 0;
                d2 = 0;
                DataTable dtn1 = bc.getdt(sqln1);
                if (dtn1.Rows.Count > 0)
                {
                    d1 = decimal.Parse(dtn1.Rows[0]["未税起机费"].ToString());
                    d2 = decimal.Parse(dtn1.Rows[0]["未税单价"].ToString());

                }
                if (dr["模切"].ToString() == "" || dr["模切"].ToString() == "否")//160107 修改模切为否时不计算模切小计
                {

                }
                else
                {
                   
                    dr["刀模小计"] = Math.Max(d1, d2 * PROCESSING_DOOR / 1000 * PROCESSING_LENGTH / 1000);
                }
                if (dr["机器型号"].ToString() == "" || dr["模切"].ToString() == "" || dr["模切"].ToString() == "否")
                {

                }
                else
                {
                    
                    dr["模切小计"] = Math.Max(MACHINING_MACHINE_FREE , DIE_CUTTING * PROCESSING_DOOR / 1000 * PROCESSING_LENGTH / 1000*TOTAL_PRODUCT_NUMBER );
                }
                if (!string.IsNullOrEmpty(dr["模切小计"].ToString()))
                {
                    TOTAL_DIE_CUTTING= decimal.Parse(dr["模切小计"].ToString());
                }
                else
                {
                    TOTAL_DIE_CUTTING = 0;
                }
                if (dr["面纸内耗"].ToString() == "")
                {

                }
                else if (dr["印刷选项"].ToString() == "不印刷")
                {
                    dr["面纸外耗"] = dr["无印刷表面处理损耗"].ToString();
                }
                else if (dr["印刷选项"].ToString() == "单纸双异")
                {
                    dr["面纸外耗"] =OPPOSITE_THE_PAPER_LOSE +POSITIVE_THE_PAPER_LOSE ;
                }
                else 
                {
                    dr["面纸外耗"] =POSITIVE_THE_PAPER_LOSE;
                }
                if(!string .IsNullOrEmpty (dr["面纸外耗"].ToString ()))
                {
                    TISSUE_OUTSIDE_LOSE =decimal .Parse (dr["面纸外耗"].ToString ());
                }
                else 
                {
                    TISSUE_OUTSIDE_LOSE =0;
                }
               
                TISSUE_DOSAGE = 0;
                TISSUE_DOSAGE  = TISSUE_OUTSIDE_LOSE + TISSUE_INSIDE_LOSE + TOTAL_PRODUCT_NUMBER;//面纸用量
                if (TISSUE_DOSAGE != 0)
                {
                    dr["面纸用量"] = TISSUE_DOSAGE;
                }
                else
                {
                    dr["面纸用量"] = DBNull.Value;
                }
                d1 = 0;
                if (dr["加工门幅"].ToString() == "" || dr["加工长度"].ToString() == "")
                {

                }
                else
                {
                    if(PROCESSING_DOOR !=0 && PROCESSING_LENGTH !=0)
                    {
                        d1 = Math.Truncate(787 / PROCESSING_DOOR) * Math.Truncate(1092 / PROCESSING_LENGTH);//正度可用
                    }
                }
              
                d2=0;
                if (d1 == 0)
                {
                }
                else
                {
                    d2 = (PROCESSING_DOOR * PROCESSING_LENGTH) * d1 / (787 * 1092);//正度利用率
                }
      
                d3 = 0;
                if (dr["加工门幅"].ToString() == "" || dr["加工长度"].ToString() == "")
                {

                }
                else if(PROCESSING_DOOR !=0 && PROCESSING_LENGTH !=0)
                {
                    d3 = Math.Truncate(889 / PROCESSING_DOOR) * Math.Truncate(1194 / PROCESSING_LENGTH);//大度可用

                }
              
                d4 = 0;
                if (d3 == 0)
                {
                }
                else
                {
                    d4 = (PROCESSING_DOOR * PROCESSING_LENGTH) * d3 / (889 * 1194);//大度利用率
                }
                d5 = 0;
                if (dr["机器型号"].ToString() == "" || dr["机器型号"].ToString() == "四开")
                {

                }
                else if (dr["面纸"].ToString() == "灰底白板")
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql+" WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE>={1}", dr["面纸"].ToString(), dr["加工门幅"].ToString());
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable  dtx =bc.getdt(sqb.ToString ());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//特规门幅
                    }
                   
                }

                else if (dr["面纸"].ToString() == "晨鸣白卡")
                {

                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE>={1}", dr["面纸"].ToString(), dr["加工门幅"].ToString());
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//特规门幅
                    }
                }

                else if (dr["面纸"].ToString() == "紫兴铜板")
                {

                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE>={1}", dr["面纸"].ToString(), dr["加工门幅"].ToString());
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//特规门幅
                    }
                }
                else if (dr["面纸"].ToString() == "酋长铜板")
                {

                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE>={1}", dr["面纸"].ToString(), dr["加工门幅"].ToString());
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//特规门幅
                    }
                }
                else
                {

                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE>={1}", dr["面纸"].ToString(), dr["加工门幅"].ToString());
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//特规门幅
                    }

                }
                //MessageBox.Show(dr["加工门幅"].ToString()+","+d5.ToString ());
                d6 = 0;
                if (d5 !=0)
                {
                    d6 = PROCESSING_LENGTH ;//特规纸长
                }
                d7 = 0;
                if (d5!=0)
                {
                    d7= Math.Truncate(d5 / PROCESSING_DOOR);//特规可用
                }
                if (dr["机器型号"].ToString() != "")
                {
                    if (d2 >= d4 && dr["机器型号"].ToString() == "四开")
                    {
                        dr["面纸门幅"] = 787;
                    }
                    else if (d2 < d4 && dr["机器型号"].ToString() == "四开")
                    {
                        dr["面纸门幅"] = 889;
                    }
                    else if (d5 != 0)
                    {

                        dr["面纸门幅"] = d5;
                    }
                    else
                    {
                        if (dr["面纸"].ToString() != "")
                        {
                            ErrowInfo = string.Format("项次 {0} 面纸暂无", i + 1);
                            break;  
                        }
                        dr["面纸门幅"] = DBNull.Value;
                      
                    }
                }

           
                if (!string.IsNullOrEmpty(dr["面纸门幅"].ToString()))
                {
                    TISSUE_DOOR = decimal.Parse(dr["面纸门幅"].ToString());
                }
                else
                {
                    TISSUE_DOOR = 0;

                }
                if (dr["机器型号"].ToString() != "")
                {
                    if (d2 >= d4 && dr["机器型号"].ToString() == "四开")
                    {
                        dr["面纸纸长"] = 1092;
                    }
                    else if (d2 < d4 && dr["机器型号"].ToString() == "四开")
                    {
                        dr["面纸纸长"] = 1194;
                    }
                    else if (d6 != 0)
                    {

                        dr["面纸纸长"] = d6;
                    }
                    else
                    {
                        dr["面纸纸长"] = DBNull.Value;
                    }
                }
                if (!string.IsNullOrEmpty(dr["面纸纸长"].ToString()))
                {
                    TISSUE_LENGTH = decimal.Parse(dr["面纸纸长"].ToString());
                }
                else
                {
                    TISSUE_LENGTH = 0;

                }
                if (dr["机器型号"].ToString() != "")
                {
                    if (d2 >= d4 && dr["机器型号"].ToString() == "四开")
                    {
                        dr["面纸可用"] = d1;
                    }
                    else if (d2 < d4 && dr["机器型号"].ToString() == "四开")
                    {
                        dr["面纸可用"] = d3;
                    }
                    else if (d7 != 0)
                    {
                        dr["面纸可用"] = d7;
                    }
                    else
                    {
                        dr["面纸可用"] = DBNull.Value;

                    }
                   
                }
                //MessageBox.Show(dr["报价编号"].ToString() + "," + TOTAL_PRODUCT_NUMBER.ToString() + "," + TISSUE_DOSAGE +","+ TISSUE_DOOR+","+TISSUE_LENGTH );
                if (dr["面纸门幅"].ToString() != "" && dr["面纸纸长"].ToString() != "" && dr["面纸可用"].ToString() != "" && TISSUE_DOOR >0 && TISSUE_LENGTH >0
                    && TOTAL_PRODUCT_NUMBER>0)
                {

                    dr["面纸单个用量"] = TISSUE_DOSAGE * TISSUE_DOOR / 1000 * TISSUE_LENGTH / 1000 /
                        decimal.Parse(dr["面纸可用"].ToString()) / COUNT;
                }
               
                if (dr["面纸单价"].ToString() == "" || dr["面纸单个用量"].ToString() == "")
                {

                }
                else
                {
                    dr["面纸小计"] = decimal.Parse(dr["面纸单价"].ToString()) * decimal.Parse(dr["面纸单个用量"].ToString()) * COUNT ;
                }
                if (!string.IsNullOrEmpty(dr["面纸小计"].ToString()))
                {
                    TOTAL_TISSUE = decimal.Parse(dr["面纸小计"].ToString());
                }
                else
                {
                    TOTAL_TISSUE = 0;
                }
                d1 = 0;//可用数A
                v1 = ""; //可用数A
            
                if (dr["芯纸"].ToString() == "")
                {

                }
                else if (dr["芯纸"].ToString() == "PVC或PET")
                {
                    d1 = Math.Truncate(600 / PROCESSING_DOOR);//可用数A
                    v1 = d1.ToString();
                    if (PROCESSING_DOOR > 915)
                    {
                        ErrowInfo = string.Format("项次：{0} 芯纸为：{1} 其加工门幅：{2} 大于 {3}", dr["项次"].ToString(), "PVC或PET", PROCESSING_DOOR,915);
                        break;
                    }
                }
                else if (dr["芯纸"].ToString() == "KT板")
                {
                    d1 = Math.Truncate(900 / PROCESSING_DOOR);//可用数A
                    v1 = d1.ToString();//可用数A
                    if (PROCESSING_DOOR > 1200)
                    {
                        ErrowInfo = string.Format("项次：{0} 芯纸为：{1} 其加工门幅：{2} 大于 {3}", dr["项次"].ToString(), "KT板", PROCESSING_DOOR,1200);
                        break;
                    }
                }
                else if (dr["芯纸"].ToString() == "双灰板")
                {
               
                    d1 =Math .Max ( Math.Truncate(787 / PROCESSING_DOOR)*Math .Truncate (1092/PROCESSING_LENGTH),
                        Math.Truncate(1092 / PROCESSING_DOOR) * Math.Truncate(787 / PROCESSING_LENGTH));//可用数A
                    v1 = d1.ToString();//可用数A
                    if (PROCESSING_DOOR > 889)
                    {
                        ErrowInfo = string.Format("项次：{0} 芯纸为：{1} 其加工门幅：{2} 大于 {3}", dr["项次"].ToString(), "双灰板", PROCESSING_DOOR,889);
                        break;
                    }
                    else if (PROCESSING_LENGTH > 1194)
                    {
                        ErrowInfo = string.Format("项次：{0} 芯纸为：{1} 其加工纸长：{2} 大于 {3}", dr["项次"].ToString(), "双灰板",PROCESSING_LENGTH,1194);
                        break;
                    }
                }
                else if (dr["芯纸"].ToString() == "AD板")
                {

                    d1= Math.Max(Math.Truncate(1220 / PROCESSING_DOOR) * Math.Truncate(2440 / PROCESSING_LENGTH),
                        Math.Truncate(2440 / PROCESSING_DOOR) * Math.Truncate(1220 / PROCESSING_LENGTH));
                    v1 = d1.ToString();//可用数A
                    if (PROCESSING_DOOR > 1220)
                    {
                        ErrowInfo = string.Format("项次：{0} 芯纸为：{1} 其加工门幅：{2} 大于 {3}", dr["项次"].ToString(), "AD板", PROCESSING_DOOR, 1220);
                        break;
                    }
                    else if (PROCESSING_LENGTH >2440)
                    {
                        ErrowInfo = string.Format("项次：{0} 芯纸为：{1} 其加工纸长：{2} 大于 {3}", dr["项次"].ToString(), "AD板", PROCESSING_LENGTH, 2440);
                        break;
                    }
                }
                else
                {
                    v1 = "瓦楞纸";//可用数A
                }
                d2 = 0;
                v2 = "";
                if (v1 == "" || v1 == "瓦楞纸")
                {

                }
                else if (dr["芯纸"].ToString() == "PVC或PET")
                {
                    d2 =d1*PROCESSING_DOOR /600 ;//利用率A
                    v2 = d2.ToString();
                }
                else if (dr["芯纸"].ToString() == "KT板")
                {
                    d2 = d1 * PROCESSING_DOOR / 910;//利用率A
                    v2 = d2.ToString();
                }
                else if (dr["芯纸"].ToString() == "双灰板")
                {
                    d2 = d1 * PROCESSING_DOOR * PROCESSING_LENGTH / (787 * 1092);//利用率A
                    v2 = d2.ToString();
                }
                else if (dr["芯纸"].ToString() == "AD板")
                {

                    d2 = d1 * PROCESSING_DOOR * PROCESSING_LENGTH / (1220 * 2440);//利用率A
                    v2 = d2.ToString();
                }
                else
                {
                    v2 = "瓦楞纸";
                }
                d3 = 0;
                v3 = "";
                if (dr["芯纸"].ToString() == "" )
                {

                }
                else if (dr["芯纸"].ToString() == "PVC或PET" && PROCESSING_DOOR!=0)
                {
                    d3 =Math .Truncate ( 915/PROCESSING_DOOR );//可用数B
                    v3 = d3.ToString();
                }
                else if (dr["芯纸"].ToString() == "KT板" && PROCESSING_DOOR != 0)
                {
                    d3 = Math.Truncate(1200 / PROCESSING_DOOR);//可用数B
                    v3 = d3.ToString();
                }
                else if (dr["芯纸"].ToString() == "双灰板")
                {
                    d3 = Math.Max(Math.Truncate(889 / PROCESSING_DOOR) * Math.Truncate(1194 / PROCESSING_LENGTH),
                        Math.Truncate(1194 / PROCESSING_DOOR) * Math.Truncate(889 / PROCESSING_LENGTH));
                    v3 = d3.ToString();
                }
                else if (dr["芯纸"].ToString() == "AD板")
                {

                    v3 = "AD板";//可用数B
                }
                else
                {
                    v3 = "瓦楞纸";//可用数B
                }
                d4 = 0;//利用率B

                if (v3 == "" || dr["芯纸"].ToString() == "瓦楞纸" || dr["芯纸"].ToString() == "坑纸" || dr["芯纸"].ToString() == "AD板")
                {

                }
                else if (dr["芯纸"].ToString() == "PVC或PET" )
                {
                    d4 =d3* PROCESSING_DOOR/915;//利用率B
             
                }
                else if (dr["芯纸"].ToString() == "KT板")
                {
                    d4 = d3 * PROCESSING_DOOR / 1200;//利用率B
                 
                }
                else 
                {
                    d4 = d3 * PROCESSING_DOOR*PROCESSING_LENGTH  / (889*1194);//利用率B
                
                }
                d5 = 0; //瓦楞纸门幅
                v5 = ""; //瓦楞纸门幅
                if (dr["芯纸"].ToString() == "" || dr["芯纸规格"].ToString() == "")
                {

                }
                else if (dr["芯纸"].ToString() == "瓦楞纸" && PROCESSING_DOOR > 0 && PROCESSING_DOOR < 433)
                {
                    d5 = 1300;
                    v5 = "1300";
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸" && PROCESSING_DOOR >= 433 && PROCESSING_DOOR < 650)
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE-20>={1}", dr["芯纸"].ToString(), PROCESSING_DOOR * 3);
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//瓦楞纸门幅
                        v5 = dtx.Rows[0]["值"].ToString();
                    }
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸" && PROCESSING_DOOR >= 650 && PROCESSING_DOOR < 1050)
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE-20>={1}", dr["芯纸"].ToString(), PROCESSING_DOOR * 2);
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//瓦楞纸门幅
                        v5 = dtx.Rows[0]["值"].ToString();
                    }
                  
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸")
                {
                   
                    sqb = new StringBuilder();
                    sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE-20>={1}", dr["芯纸"].ToString(), PROCESSING_DOOR);
                    sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                    DataTable dtx = bc.getdt(sqb.ToString());
                    if (dtx.Rows.Count > 0)
                    {
                        d5 = decimal.Parse(dtx.Rows[0]["值"].ToString());//瓦楞纸门幅
                        v5 = dtx.Rows[0]["值"].ToString();
                    }
                  
                    //MessageBox.Show(d5.ToString() + "," + PROCESSING_DOOR);
                }
                else
                {
                   
                    v5 = "非瓦楞";
                }
             
                if (dr["芯纸"].ToString() == "瓦楞纸" && dr["芯纸规格"].ToString ()!="" && d5 == 0)
                {
                    ErrowInfo = string.Format("项次 {0} 芯纸暂无或该客户类别下找不到匹配的材料门幅参数值", i + 1);
                    break;
                }
                d6 = 0;//瓦楞纸纸长
                if (v5 == "" || v5 == "非瓦楞")
                {

                }
                else
                {
                    d6 = decimal.Parse(dr["加工长度"].ToString());//瓦楞纸纸长
                }
                d7 = 0;//瓦楞可用
           
                if (v5 == "" || v5 == "暂无" || v5 == "非瓦楞")
                {

                }
                else if (PROCESSING_DOOR >= 433 && PROCESSING_DOOR < 650)
                {
                    d7 = 3;//瓦楞可用
                }
                else if (PROCESSING_DOOR >= 650 && PROCESSING_DOOR <= 1050)
                {
                    d7 = 2;//瓦楞可用
                }
                else if (PROCESSING_DOOR > 1050 && PROCESSING_DOOR <= 2500)
                {
                    d7 = 1;//瓦楞可用
                }
                else
                {
                    d7 = 1300 / PROCESSING_DOOR; //瓦楞可用
                    if (d7 > 0)
                    {
                        string  va1 = d7.ToString();
                         d7 = decimal.Parse(bc.RETURN_UNTIL_CHAR (va1,'.'));//16/01/07 只要整数部分且不用四舍五入的值
                    }
                }
            
                if (dr["芯纸"].ToString() == "")
                {

                }
                else  if (dr["芯纸"].ToString() == "坑纸")
                {
                    dr["芯纸门幅"] = dr["加工门幅"].ToString();
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸")
                {
                    dr["芯纸门幅"] = v5;
                }
                else if (dr["芯纸"].ToString() == "双灰板" && d2>=d4)
                {
                    dr["芯纸门幅"] = 787;
                }
                else if (dr["芯纸"].ToString() == "双灰板" && d2 < d4)
                {
                    dr["芯纸门幅"] = 889;
                }
                else if (dr["芯纸"].ToString() == "KT板" && d2 >= d4)
                {
                    dr["芯纸门幅"] = 900;
                }
                else if (dr["芯纸"].ToString() == "KT板" && d2 < d4)
                {
                    dr["芯纸门幅"] =1200;
                }
                else if (dr["芯纸"].ToString() == "PVC或PET" && d2 >= d4)
                {
                    dr["芯纸门幅"] = 600;
                }
                else if (dr["芯纸"].ToString() == "PVC或PET" && d2 < d4)
                {
                    dr["芯纸门幅"] = 915;
                }
                else
                {
                    dr["芯纸门幅"] = 1220;
                }
                sqb = new StringBuilder();
                sqb.AppendFormat(cdoor_parameters.sql + " WHERE  B.DOOR_PARAMETERS='{0}' AND A.PRICE>={1}", dr["芯纸"].ToString(), PROCESSING_DOOR);
                sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}' ORDER BY A.PRICE ASC", CUSTOMER_TYPE);
                dtx1 = bc.getdt(sqb.ToString());
                if (dtx1.Rows.Count > 0)
                {
                    d5 = decimal.Parse(dtx1.Rows[0]["值"].ToString());//瓦楞纸门幅
                    v5 = dtx1.Rows[0]["值"].ToString();
                }
           
                if(!string .IsNullOrEmpty (dr["芯纸门幅"].ToString ()))
                {
                    PAPER_CORE_DOOR =decimal .Parse (dr["芯纸门幅"].ToString ());
                }
                else 
                {
                    PAPER_CORE_DOOR =0;
                }
                if (dr["芯纸"].ToString() == "")
                {

                }
                else if (dr["芯纸"].ToString() == "坑纸" || dr["芯纸"].ToString() == "瓦楞纸" || dr["芯纸"].ToString() == "KT板" || dr["芯纸"].ToString() == "PVC或PET")
                {
                    dr["芯纸纸长"] = decimal.Parse(dr["加工长度"].ToString());
                }
                else if (dr["芯纸"].ToString() == "双灰板" && d2 >= d4)
                {
                    dr["芯纸纸长"] = 1092;
                }
                else if (dr["芯纸"].ToString() == "双灰板" && d2 < d4)
                {
                    dr["芯纸纸长"] = 1194;
                }
                else
                {
                    dr["芯纸纸长"] = 2440;
                }
                if(!string .IsNullOrEmpty (dr["芯纸纸长"].ToString ()))
                {
                    PAPER_CORE_LENGTH =decimal .Parse (dr["芯纸纸长"].ToString ());
                }
                else 
                {

                    PAPER_CORE_LENGTH =0;
                }
                if (dr["芯纸"].ToString() == "")
                {

                }
                else if (dr["芯纸"].ToString() == "坑纸")
                {
                    dr["芯纸可用"] = 1;
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸")
                {
                    dr["芯纸可用"] = d7;
                }
                else if (dr["芯纸"].ToString() == "AD板")
                {
                    dr["芯纸可用"] = v1;
                }
                else if (d2 >= d4 && (dr["芯纸"].ToString() == "双灰板" || dr["芯纸"].ToString() == "KT板" || dr["芯纸"].ToString() == "PVC或PET"))
                {
                    dr["芯纸可用"] = v1;
                }
                else
                {
                    dr["芯纸可用"] = v3;
                }
                if (bc.yesno(dr["芯纸可用"].ToString()) != 0 && !string.IsNullOrEmpty(dr["芯纸可用"].ToString()))
                {
                    PAPER_CORE_AVAILABLE = decimal.Parse(dr["芯纸可用"].ToString());
                }
               
                if (dr["芯纸"].ToString() == "" || dr["芯纸规格"].ToString() == "" || dr["芯纸门幅"].ToString() == "")
                {

                }
                else if (dr["芯纸"].ToString() == "KT板" && dr["芯纸门幅"].ToString() == "900")
                {
                  
                      dtx = bc.getdt(cpaper_core .sql  + string.Format(@" WHERE  B.PAPER_CORE='{0}' AND A.SPEC='{1}'  
                    AND A.PAPER_CORE_DOOR='{2}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{3}'",
                          dr["芯纸"].ToString(), dr["芯纸规格"].ToString(), dr["芯纸门幅"].ToString(),CUSTOMER_TYPE ));
                  
                    if (dtx.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单价"].ToString()))
                        {
                            dr["芯纸单价"] = decimal.Parse(dtx.Rows[0]["单价"].ToString());
                        }
                     
                    }
                }
                else if (dr["芯纸"].ToString() == "KT板" && dr["芯纸门幅"].ToString() == "1200")
                {
                    dtx = bc.getdt(cpaper_core.sql + string.Format(@" WHERE  B.PAPER_CORE='{0}' AND A.SPEC='{1}'  
AND A.PAPER_CORE_DOOR='{2}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{3}'",
                    dr["芯纸"].ToString(), dr["芯纸规格"].ToString(), dr["芯纸门幅"].ToString(),CUSTOMER_TYPE ));
                    if (dtx.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单价"].ToString()))
                        {
                            dr["芯纸单价"] = decimal.Parse(dtx.Rows[0]["单价"].ToString());
                        }

                    }
                }
                else
                {

                    dtx = bc.getdt(cpaper_core.sql + string.Format(" WHERE  B.PAPER_CORE='{0}' AND A.SPEC='{1}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{2}'",
                    dr["芯纸"].ToString(), dr["芯纸规格"].ToString(),CUSTOMER_TYPE ));

                    if (dtx.Rows.Count > 0 && !string.IsNullOrEmpty(dtx.Rows[0]["单价"].ToString()))
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[0]["单价"].ToString()))
                        {
                            dr["芯纸单价"] = decimal.Parse(dtx.Rows[0]["单价"].ToString());
                        }
                       //注根据输入时允许单价为空，所以此芯纸单价可能为空

                    }
                }
                d8 = 0;//内分步一
                v8 = "";//内分步一
                if (dr["芯纸单价"].ToString() == "" || dr["印刷选项"].ToString() == "" || TOTAL_PRODUCT_NUMBER == 0)
                {
                    
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸" && TOTAL_PRODUCT_NUMBER <= 300 && dr["面纸小计"].ToString() == "")
                {
                    d8 = 8;//内分步一
                    v8 = "8";//内分步一
                }
                else if (dr["芯纸"].ToString() == "瓦楞纸" && TOTAL_PRODUCT_NUMBER >300 && dr["面纸小计"].ToString() == "")
                {
                    /*if (!string.IsNullOrEmpty(dr["面纸小计"].ToString()) && bc.yesno(dr["面纸小计"].ToString()) != 0)
                    {
                        d8 = 8 + (TOTAL_PRODUCT_NUMBER - 300) * 1 / 2 * 1 / 100;//内分步一
                        v8 = d8.ToString();//内分步一
                    }151112 ago*/
                  
                      d8 = 8 + (TOTAL_PRODUCT_NUMBER - 300) * 1 / 2 * 1 / 100;//内分步一
                      v8 = d8.ToString();//内分步一
                }
                else if ((dr["芯纸"].ToString() == "瓦楞纸" || dr["芯纸"].ToString() == "坑纸" || dr["芯纸"].ToString() == "双灰板") && 
                    TOTAL_PRODUCT_NUMBER <= 300 )
                {
                    d8 = 15;//内分步一
                    v8 = "15";//内分步一
                }
                else if ((dr["芯纸"].ToString() == "瓦楞纸" || dr["芯纸"].ToString() == "坑纸" || dr["芯纸"].ToString() == "双灰板") &&
                    TOTAL_PRODUCT_NUMBER > 300)
                {
                    d8 = 15 + (TOTAL_PRODUCT_NUMBER  - 300) * 6 / 5 * 1 / 100;//内分步一
                    v8 = d8.ToString();//内分步一
                }
                else if ((dr["芯纸"].ToString() == "KT板" || dr["芯纸"].ToString() == "AD板") && TOTAL_PRODUCT_NUMBER <= 300)
                {
                    d8 = 5;//内分步一
                    v8 = "5";//内分步一
                }
                else
                {
                    v8 = "Next";//内分步一
                }
                //MessageBox.Show(v8);
                d2 = 0;
                d3 = 0;
                dtt = bc.getdt(cpaper_core_option .sql + string.Format(" WHERE A.PAPER_CORE='{0}'", dr["芯纸"].ToString()));
                if (dtt.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtt.Rows[0]["芯纸内耗1到300"].ToString()))
                    {
                        d2 = decimal.Parse(dtt.Rows[0]["芯纸内耗1到300"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtt.Rows[0]["芯纸内耗大于300"].ToString()))
                    {
                        d3 = decimal.Parse(dtt.Rows[0]["芯纸内耗大于300"].ToString());
                    }

                }
                if (dr["芯纸单价"].ToString() == "" || dr["印刷选项"].ToString() == "" || TOTAL_PRODUCT_NUMBER==0)
                {

                }
                else if(TOTAL_PRODUCT_NUMBER <=300)
                {
                    dr["芯纸内耗"] = d2;
                }
                else
                {
                    dr["芯纸内耗"] = d2 + (TOTAL_PRODUCT_NUMBER - 300) * d3 * 1 / 100;
                }
           
                if (!string.IsNullOrEmpty(dr["芯纸内耗"].ToString()) && bc.yesno(dr["芯纸内耗"].ToString()) != 0)
                {
                    dr["芯纸用量"] = decimal.Parse(dr["芯纸内耗"].ToString()) + TOTAL_PRODUCT_NUMBER;
                }
                if(!string .IsNullOrEmpty (dr["芯纸用量"].ToString ()))
                {
                    PAPER_CORE_DOSAGE =decimal .Parse (dr["芯纸用量"].ToString ());
                }
                else 
                {
                    PAPER_CORE_DOSAGE =0;
                }
            
                if (dr["芯纸门幅"].ToString() == "" || dr["芯纸纸长"].ToString() == "" || dr["芯纸可用"].ToString() == "")
                {

                }
                else if (dr["芯纸"].ToString() == "瓦楞纸" && TOTAL_PRODUCT_NUMBER >0 && PAPER_CORE_AVAILABLE >0)
                                                                        //  PAPER_CORE_DOOR 芯纸门幅,
                                                                       //PAPER_CORE_DOSAGE 芯纸用量, PAPER_CORE_LENGTH 芯纸纸长,
                                                                       //PAPER_CORE_AVAILABLE 芯纸可用
                {
                    dr["芯纸单个用量"] = Math.Max(PAPER_CORE_DOOR / 1000 * 100 /COUNT , 
                        PAPER_CORE_DOSAGE * PAPER_CORE_DOOR / 1000 * PAPER_CORE_LENGTH / 1000 / PAPER_CORE_AVAILABLE / COUNT );
                }
                else if (TOTAL_PRODUCT_NUMBER > 0 && PAPER_CORE_AVAILABLE > 0)
                {
                    dr["芯纸单个用量"] = PAPER_CORE_DOSAGE * PAPER_CORE_DOOR / 1000 * PAPER_CORE_LENGTH / 1000 / PAPER_CORE_AVAILABLE / COUNT ;
                }
                /*sqb = new StringBuilder();
                sqb.AppendFormat("部品名：{0},", dr["部品名"].ToString());
                sqb.AppendFormat("数量：{0},", COUNT );
                sqb.AppendFormat("部品总数：{0},", TOTAL_PRODUCT_NUMBER);
                sqb.AppendFormat("芯纸门幅：{0},", PAPER_CORE_DOOR);
                sqb.AppendFormat("芯纸用量：{0},", PAPER_CORE_DOSAGE );
                sqb.AppendFormat("芯纸纸长：{0},", PAPER_CORE_LENGTH);
                sqb.AppendFormat("芯纸可用：{0},", PAPER_CORE_AVAILABLE);
                sqb.AppendFormat("比较值一：{0},", PAPER_CORE_DOOR / 1000 * 100 / COUNT );
                sqb.AppendFormat("比较值二：{0},", PAPER_CORE_DOSAGE * PAPER_CORE_DOOR / 1000 * PAPER_CORE_LENGTH / 1000 / PAPER_CORE_AVAILABLE / COUNT);
               // MessageBox.Show(sqb.ToString());*/
                //16/01/06
                if (dr["芯纸单价"].ToString() == "" || dr["芯纸单个用量"].ToString() == "")
                {

                }
                else
                {
                    dr["芯纸小计"] = decimal.Parse(dr["芯纸单价"].ToString()) * decimal.Parse(dr["芯纸单个用量"].ToString())* COUNT ;
                }
                if (!string.IsNullOrEmpty(dr["芯纸小计"].ToString()))
                {
                    TOTAL_PAPAER_CORE  = decimal.Parse(dr["芯纸小计"].ToString());
                }
                else
                {
                    TOTAL_PAPAER_CORE = 0;
                }  
              
                if (dr["部品总数"].ToString() == "" || dr["印刷选项"].ToString() == "" || dr["底纸"].ToString() == "" || dr["底纸克重"].ToString() == "")
                {

                }
                else
                {
                    if (!string.IsNullOrEmpty(dr["底纸"].ToString()) && !string.IsNullOrEmpty(dr["底纸克重"].ToString()))
                    {
                        v1 = bc.getOnlyString(string.Format(@"
SELECT
A.TON_PRICE
FROM TISSUE_SPEC_DET A
LEFT JOIN TISSUE_SPEC_MST B ON A.TSID=B.TSID 
WHERE
B.TISSUE_SPEC='{0}' AND A.WEIGHT='{1}' AND SUBSTRING(B.CUSTOMER_TYPE,1,1)='{2}'", dr["底纸"].ToString(), dr["底纸克重"].ToString(),CUSTOMER_TYPE ));
                        if (!string.IsNullOrEmpty(v1))
                        {
                            d1 = decimal.Parse(v1);
                            d2 = decimal.Parse(dr["底纸克重"].ToString());
                            dr["底纸单价"] = (d1 * d2 / 1000000);
                        }

                    }
                   
                }
                d2 = 0;
                d3 = 0;
                if (dtx2.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["底纸内耗1到300"].ToString()))
                    {
                        d2 = decimal.Parse(dtx2.Rows[0]["底纸内耗1到300"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx2.Rows[0]["底纸内耗大于300"].ToString()))
                    {
                        d3 = decimal.Parse(dtx2.Rows[0]["底纸内耗大于300"].ToString());
                    }
                }
                if (dr["部品总数"].ToString() == "" || dr["印刷选项"].ToString() == "" || dr["底纸单价"].ToString() == "")
                {

                }
                else if (TOTAL_PRODUCT_NUMBER <= 300)
                {
                    dr["底纸内耗"] = d2;

                }
                else 
                {
                    dr["底纸内耗"] = d2 + (TOTAL_PRODUCT_NUMBER - 300) * d3 * 1 / 100;
             
                }
            
                if (!string.IsNullOrEmpty(dr["底纸内耗"].ToString()))
                {
                    BODY_PAPER_INSIDE_LOSE  = decimal.Parse(dr["底纸内耗"].ToString());
                }
                else
                {
                    BODY_PAPER_INSIDE_LOSE = 0;
                }

                dr["底纸下单"] = TOTAL_PRODUCT_NUMBER +BODY_PAPER_INSIDE_LOSE ;
                if (dr["底纸下单"].ToString() == "" || dr["底纸单价"].ToString() == "")
                {

                }
                else if (dr["印刷选项"].ToString() == "双纸画异")
                {
                    dr["底纸外耗"] = dr["反面纸张损耗"].ToString();
                }
                else
                {
                    dr["底纸外耗"] = 0;
                }
                if (!string.IsNullOrEmpty(dr["底纸外耗"].ToString()))
                {
                    BODY_PAPER_OUTSIDE_LOSE = decimal.Parse(dr["底纸外耗"].ToString());
                }
                else
                {
                    BODY_PAPER_OUTSIDE_LOSE = 0;
                }
                d1 = 0;
                d1 = TOTAL_PRODUCT_NUMBER + BODY_PAPER_INSIDE_LOSE+BODY_PAPER_OUTSIDE_LOSE  ;//底纸用量
                if (d1 != 0)
                {
                    dr["底纸用量"] = d1;
                }
                else
                {
                    dr["底纸用量"] = "";
                }
                decimal dx1 = 0;
                if (!string.IsNullOrEmpty(dr["面纸可用"].ToString()))
                {
                    dx1 = decimal.Parse(dr["面纸可用"].ToString());
                }
                else
                {
                    dx1 = 0;
                }
                if (dr["面纸门幅"].ToString() == "" || dr["面纸纸长"].ToString() == "" || d1 == 0)
                {

                }
                else if(TOTAL_PRODUCT_NUMBER >0 && TISSUE_DOOR >0 && TISSUE_LENGTH >0 && dx1>0)
                {
                
                    dr["底纸单个用量"] = d1 * TISSUE_DOOR  / 1000 * TISSUE_LENGTH  / 1000 
                        / decimal.Parse(dr["面纸可用"].ToString()) / COUNT  ;
                }
         
                if (dr["底纸单个用量"].ToString() == "" || dr["底纸单价"].ToString() == "" || TOTAL_PRODUCT_NUMBER ==0)
                {

                }
                else 
                {
                    dr["底纸小计"] = decimal.Parse(dr["底纸单个用量"].ToString())* decimal.Parse(dr["底纸单价"].ToString())*COUNT ;
                }
                if (!string.IsNullOrEmpty(dr["底纸小计"].ToString()))
                {
                    TOTAL_BODY_PAPER  = decimal.Parse(dr["底纸小计"].ToString());
                }
                else
                {
                   TOTAL_BODY_PAPER = 0;
                }
                
                if (TOTAL_PRODUCT_NUMBER == 0)
                {
                  
                }
                else if(dr["印刷选项"].ToString() != "不印刷" && dr["机器型号"].ToString() == "四开" && Math .Min (PROCESSING_DOOR ,PROCESSING_LENGTH )<=300)
                {
                    dr["部品总价"] = "四开小";
                }
                else if (dr["印刷选项"].ToString() != "不印刷" && dr["机器型号"].ToString() == "对开" && Math.Min(PROCESSING_DOOR, PROCESSING_LENGTH) <= 350)
                {
                    dr["部品总价"] = "对开小";
                }
                else if (dr["印刷选项"].ToString() != "不印刷" && dr["机器型号"].ToString() == "全开" && Math.Min(PROCESSING_DOOR, PROCESSING_LENGTH) <= 500)
                {
                    dr["部品总价"] = "全开小";
                }
                else if (dr["印刷选项"].ToString() != "不印刷" && dr["机器型号"].ToString() == "大全开" && Math.Min(PROCESSING_DOOR, PROCESSING_LENGTH) <= 350)
                 {
                     dr["部品总价"] = "大全开小";
                 }
                else if (PROCESSING_DOOR >1200 && PROCESSING_LENGTH >1620)
                {
                    dr["部品总价"] = "超出大全开";
                }
               else
               {
                  decimal dx = TOTAL_TISSUE + TOTAL_PAPAER_CORE + TOTAL_BODY_PAPER + TOTAL_POSITIVE_AND_OPPOSITE_PRINTING +
                     TOTAL_SURFACE_PROCESSING + TOTAL_LAMINATING_PROCESS + TOTAL_DIE_CUTTING + POSITIVE_CTP_PRICE_TOTAL + OPPOSITE_CTP_PRICE_TOTAL;
                  dr["部品总价"] = dx.ToString("#0");
               }
                if (dr["部品总数"].ToString() == "" || dr["部品总价"].ToString() == "" || COUNT ==0)
                {

                }
                else if (bc.yesno(dr["部品总价"].ToString()) != 0)//部品总价不含文字的才求部品单价
                {

                    dr["部品单价"] = decimal.Parse(dr["部品总价"].ToString()) / COUNT;

                }
                //MessageBox.Show(dr["表面加工小计"].ToString()+","+TOTAL_SURFACE_PROCESSING );
               // MessageBox.Show(TOTAL_TISSUE + "," + TOTAL_PAPAER_CORE + "," + TOTAL_BODY_PAPER + "," + TOTAL_POSITIVE_AND_OPPOSITE_PRINTING + "," +
                     //TOTAL_SURFACE_PROCESSING + "," + TOTAL_LAMINATING_PROCESS + "," + TOTAL_DIE_CUTTING);
                i = i + 1;
            }
            return dt;
        }
        #endregion
        #region RETURN_DT_TO_EXCEL_TOTAL
        public DataTable RETURN_DT_TO_EXCEL_TOTAL(DataTable dtt)
        {
            DataTable dt = dtt;
            DataRow dr2 = dt.NewRow();
            d1 = 0;
            d2 = 0;
            d3 = 0;
            d4 = 0;
            d5 = 0;
            d6 = 0;
            d7 = 0;
            d8 = 0;
            d9 = 0;
            d10 = 0;
            dr2["机器型号"] = "主体";
            dr2["面纸可用"] = "单项总价";
            if (!string.IsNullOrEmpty(dt.Compute("SUM(面纸小计)", "").ToString()))
            {
                d2 = decimal.Parse(dt.Compute("SUM(面纸小计)", "").ToString());
                dr2["面纸小计"] = d2.ToString();
            }
            TOTAL_COST_TISSUE = d2;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(芯纸小计)", "").ToString()))
            {
                d3 = decimal.Parse(dt.Compute("SUM(芯纸小计)", "").ToString());
                dr2["芯纸小计"] = d3.ToString();
            }
            TOTAL_COST_PAPAER_CORE = d3;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(底纸小计)", "").ToString()))
            {
                d4 = decimal.Parse(dt.Compute("SUM(底纸小计)", "").ToString());
                dr2["底纸小计"] = d4.ToString();
            }
            TOTAL_COST_BODY_PAPER = d4;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(表面加工小计)", "").ToString()))
            {
                d5 = decimal.Parse(dt.Compute("SUM(表面加工小计)", "").ToString());
                dr2["表面加工小计"] = d5.ToString();
            }
            TOTAL_COST_SURFACE_PROCESSING = d5;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(裱工小计)", "").ToString()))
            {
                d6 = decimal.Parse(dt.Compute("SUM(裱工小计)", "").ToString());
                dr2["裱工小计"] = d6.ToString();
            }
            TOTAL_COST_LAMINATING_PROCESS = d6;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(刀模小计)", "").ToString()))
            {
                d10 = decimal.Parse(dt.Compute("SUM(刀模小计)", "").ToString());
            }
            TOTAL_COST_CUTTING = d10;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(模切小计)", "").ToString()))
            {
                d7 = decimal.Parse(dt.Compute("SUM(模切小计)", "").ToString());

                dr2["模切小计"] = d7.ToString();
            }
            TOTAL_COST_DIE_CUTTING = d7;
            if (!string.IsNullOrEmpty(dt.Compute("SUM(正反CTP合计)", "").ToString()))
            {
                d9 = decimal.Parse(dt.Compute("SUM(正反CTP合计)", "").ToString());
                //dr2["正反CTP合计"] = d9.ToString("0");
            }
           
            if (!string.IsNullOrEmpty(dt.Compute("SUM(正反印工合计)", "").ToString()))
            {
                d8 = decimal.Parse(dt.Compute("SUM(正反印工合计)", "").ToString());
                dr2["正反印工合计"] = (d8+d9).ToString();
            }
            d8 = d8 + d9;
            if (d2 + d3 + d4 + d5 + d6 + d7 + d8 > 0)
            {
                d1 = d2 + d3 + d4 + d5 + d6 + d7 + d8;
                dr2["部品总价"] = d1.ToString();
            }
            TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP = d8;
            //MessageBox.Show(TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP.ToString() + "excel");
            dt.Rows.Add(dr2);

            DataRow dr3 = dt.NewRow();
            dr3["机器型号"] = "汇总";
            dr3["面纸可用"] = "单项单价";
            if (d1 != 0)
            {
                dr3["部品总价"] = (d1 / COUNT).ToString();
            }
            if (d2 != 0)
            {
                dr3["面纸小计"] = (d2 / COUNT).ToString();
            }
            if (d3 != 0)
            {
                dr3["芯纸小计"] = (d3 / COUNT).ToString();
            }
            if (d4 != 0)
            {
                dr3["底纸小计"] = (d4 / COUNT).ToString();
            }
            if (d5 != 0)
            {
                dr3["表面加工小计"] = (d5 / COUNT).ToString();
            }
            if (d6 != 0)
            {
                dr3["裱工小计"] = (d6 / COUNT).ToString();
            }
            if (d7 != 0)
            {
                dr3["模切小计"] = (d7 / COUNT).ToString();
            }
            if (d8 != 0)
            {
                dr3["正反印工合计"] = (d8 / COUNT).ToString();
            }
            dt.Rows.Add(dr3);
            return dt;
      
        }
        #endregion
        #region RETURN_DT_SHOW_HIDE
        public DataTable RETURN_DT_SHOW_HIDE(DataTable dtt)
        {
            DataTable dt = GetTableInfo_show_hide();
            int i = 1;
            foreach (DataRow dr1 in dtt.Rows)
            {
               
                DataRow dr = dt.NewRow();
                dr["序号"] = i;
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["客户"] = dr1["客户"].ToString();
                dr["品牌"] = dr1["品牌"].ToString();
                dr["AE"] = dr1["AE"].ToString();
                dr["数量"] = dr1["数量"].ToString();
                dr["项目号"] = dr1["项目号"].ToString();
                dr["报价编号"] = dr1["报价编号"].ToString();
                dr["报价"] = dr1["报价"].ToString();
                dr["日期"] = dr1["日期"].ToString();
                dr["部品名"] = dr1["部品名"].ToString();
                if (!string.IsNullOrEmpty(dr1["加工门幅"].ToString()))
                {
                    dr["加工门幅"] = dr1["加工门幅"].ToString();
                }
                else
                {
                    dr["加工门幅"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["加工长度"].ToString()))
                {
                    dr["加工长度"] = dr1["加工长度"].ToString();
                }
                else
                {
                    dr["加工长度"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["部品总数"].ToString()) && bc.yesno(dr1["部品总数"].ToString()) != 0)
                {
                    d1 = decimal.Parse(dr1["部品总数"].ToString());
                    dr["部品总数"] = d1.ToString();
                }
                else
                {
                    dr["部品总数"] = DBNull.Value;
                }
         
                dr["机器型号"] = dr1["机器型号"].ToString();
                if (!string.IsNullOrEmpty(dr1["部品单价"].ToString()) && bc.yesno(dr1["部品单价"].ToString()) != 0)
                {
                    d1 = decimal.Parse(dr1["部品单价"].ToString());
                    dr["部品单价"] = d1.ToString();
                }
                else
                {
                    dr["部品单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["部品总价"].ToString()))
                {
                    dr["部品总价"] = dr1["部品总价"].ToString();
                }
                else
                {
                    dr["部品总价"] = DBNull.Value;
                }

                if (!string.IsNullOrEmpty(dr1["面纸单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸单价"].ToString());
                    dr["面纸单价"] = d1.ToString();
                }
                else
                {
                    dr["面纸单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸用量"].ToString());
                    dr["面纸用量"] = d1.ToString();
                }
                else
                {
                    dr["面纸用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸内耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸内耗"].ToString());
                    dr["面纸内耗"] = d1.ToString();
                }
                else
                {
                    dr["面纸内耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸下单"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸下单"].ToString());
                    dr["面纸下单"] = d1.ToString();
                }
                else
                {
                    dr["面纸下单"] = DBNull.Value;
                }
                dr["面纸外耗"] = dr1["面纸外耗"].ToString();
                dr["面纸门幅"] = dr1["面纸门幅"].ToString();
                if (!string.IsNullOrEmpty(dr1["面纸纸长"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸纸长"].ToString());
                    dr["面纸纸长"] = d1.ToString();
                }
                else
                {
                    dr["面纸纸长"] = DBNull.Value;
                }
                dr["面纸可用"] = dr1["面纸可用"].ToString();
                if (!string.IsNullOrEmpty(dr1["面纸单个用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸单个用量"].ToString());
                    dr["面纸单个用量"] = d1.ToString();
                }
                else
                {
                    dr["面纸单个用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸小计"].ToString());
                    dr["面纸小计"] = d1.ToString();
                }
                else
                {

                    dr["面纸小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸单价"].ToString()))
                {

                    d1 = decimal.Parse(dr1["芯纸单价"].ToString());
                    dr["芯纸单价"] = d1.ToString();
                }
                else
                {
                    dr["芯纸单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸内耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸内耗"].ToString());
                    dr["芯纸内耗"] = d1.ToString();
                }
                else
                {
                    dr["芯纸内耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸用量"].ToString());
                    dr["芯纸用量"] = d1.ToString();
                }
                else
                {

                    dr["芯纸用量"] = DBNull.Value;
                }
                dr["芯纸门幅"] = dr1["芯纸门幅"].ToString();
                if (!string.IsNullOrEmpty(dr1["芯纸纸长"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸纸长"].ToString());
                    dr["芯纸纸长"] = d1.ToString();
                }
                else
                {
                    dr["芯纸纸长"] = DBNull.Value;
                }
                dr["芯纸可用"] = dr1["芯纸可用"].ToString();
                if (!string.IsNullOrEmpty(dr1["芯纸单个用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸单个用量"].ToString());
                    dr["芯纸单个用量"] = d1.ToString();
                }
                else
                {
                    dr["芯纸单个用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸小计"].ToString());
                    dr["芯纸小计"] = d1.ToString();
                }
                else
                {

                    dr["芯纸小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸单价"].ToString());
                    dr["底纸单价"] = d1.ToString();
                }
                else
                {
                    dr["底纸单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸用量"].ToString());
                    dr["底纸用量"] = d1.ToString();
                }
                else
                {
                    dr["底纸用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸内耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸内耗"].ToString());
                    dr["底纸内耗"] = d1.ToString();
                }
                else
                {

                    dr["底纸内耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸下单"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸下单"].ToString());
                    dr["底纸下单"] = d1.ToString();
                }
                else
                {

                    dr["底纸下单"] = DBNull.Value;
                }
             
                dr["底纸外耗"] = dr1["底纸外耗"].ToString();
                if (!string.IsNullOrEmpty(dr1["底纸单个用量"].ToString()))
                {

                    d1 = decimal.Parse(dr1["底纸单个用量"].ToString());
                    dr["底纸单个用量"] = d1.ToString();
                }
                else
                {
                    dr["底纸单个用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸小计"].ToString());
                    dr["底纸小计"] = d1.ToString();
                }
                else
                {

                    dr["底纸小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["印工单色单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["印工单色单价"].ToString());
                    dr["印工单色单价"] = d1.ToString();
                }
                else
                {
                    dr["印工单色单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["超出单色单张价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["超出单色单张价"].ToString());
                    dr["超出单色单张价"] = d1.ToString();
                }
                else
                {
                    dr["超出单色单张价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["CTP单张价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["CTP单张价"].ToString());
                    dr["CTP单张价"] = d1.ToString();
                }
                else
                {
                    dr["CTP单张价"] = DBNull.Value;
                }
             
                dr["正面色数共计"] = dr1["正面色数共计"].ToString();
                dr["正面CTP张数"] = dr1["正面CTP张数"].ToString();
                dr["正面纸张损耗"] = dr1["正面纸张损耗"].ToString();
                if (!string.IsNullOrEmpty(dr1["正面防晒合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面防晒合计"].ToString());
                    dr["正面防晒合计"] = d1.ToString();
                }
                else
                {
                    dr["正面防晒合计"] = DBNull.Value;
                }

                if (!string.IsNullOrEmpty(dr1["正面CTP价计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面CTP价计"].ToString());
                    dr["正面CTP价计"] = d1.ToString();
                }
                else
                {
                    dr["正面CTP价计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正面印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面印工合计"].ToString());
                    dr["正面印工合计"] = d1.ToString();
                }
                else
                {

                    dr["正面印工合计"] = DBNull.Value;
                }
       
                dr["反面色数共计"] = dr1["反面色数共计"].ToString();
                dr["反面CTP张数"] = dr1["反面CTP张数"].ToString();
                dr["反面纸张损耗"] = dr1["反面纸张损耗"].ToString();
                if (!string.IsNullOrEmpty(dr1["反面防晒合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["反面防晒合计"].ToString());
                    dr["反面防晒合计"] = d1.ToString();
                }
                else
                {
                    dr["反面防晒合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["反面CTP价计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["反面CTP价计"].ToString());
                    dr["反面CTP价计"] = d1.ToString();
                }
                else
                {
                    dr["反面CTP价计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["反面印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["反面印工合计"].ToString());
                    dr["反面印工合计"] = d1.ToString();
                }
                else
                {
                    dr["反面印工合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正反CTP合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正反CTP合计"].ToString());
                    dr["正反CTP合计"] = d1.ToString();
                }
                else
                {
                    dr["正反CTP合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正反印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正反印工合计"].ToString());
                    dr["正反印工合计"] = d1.ToString();
                }
                else
                {
                    dr["正反印工合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["表面处理单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面处理单价"].ToString());
                    dr["表面处理单价"] = d1.ToString();
                }
                else
                {
                    dr["表面处理单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["无印刷表面处理损耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["无印刷表面处理损耗"].ToString());
                    dr["无印刷表面处理损耗"] = d1.ToString();
                }
                else
                {

                    dr["无印刷表面处理损耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["表面处理用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面处理用量"].ToString());
                    dr["表面处理用量"] = d1.ToString();
                }
                else
                {
                    dr["表面处理用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["表面加工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面加工小计"].ToString());
                    dr["表面加工小计"] = d1.ToString();
                }
                else
                {

                    dr["表面加工小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["裱工单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工单价"].ToString());
                    dr["裱工单价"] = d1.ToString();
                }
                else
                {
                    dr["裱工单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["裱工用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工用量"].ToString());
                    dr["裱工用量"] = d1.ToString();
                }
                else
                {
                    dr["裱工用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["裱工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工小计"].ToString());
                    dr["裱工小计"] = d1.ToString();
                }
                else
                {

                    dr["裱工小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["刀模小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["刀模小计"].ToString());
                    dr["刀模小计"] = d1.ToString();
                }
                else
                {
                    dr["刀模小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["模切小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["模切小计"].ToString());
                    dr["模切小计"] = d1.ToString();
                }
                else
                {
                    dr["模切小计"] = DBNull.Value;
                }
                dt.Rows.Add(dr);
                i = i + 1;
            }

            if (RETURN_IF_EXISTS_ABOVE_PROJECT(dt).Rows.Count > 1)
            {

            }
            else
            {
                dt = RETURN_DT_TO_EXCEL_TOTAL(dt);
            }
            return dt;
        }
        #endregion
        #region RETURN_DT_SHOW_HIDE_FORM
        public DataTable RETURN_DT_SHOW_HIDE_FORM(DataTable dtt)//用于显示数据时只显示需要显示的小数位
        {
            DataTable dt = GetTableInfo_show_hide();
            int i = 1;
            DataTable  dtt1 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "面纸可用 NOT IN ('单项总价','单项单价')");
            foreach (DataRow dr1 in dtt1.Rows)
            {
          
                DataRow dr = dt.NewRow();
                dr["序号"] = i;
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["客户"] = dr1["客户"].ToString();
                dr["品牌"] = dr1["品牌"].ToString();
                dr["AE"] = dr1["AE"].ToString();
                dr["数量"] = dr1["数量"].ToString();
                dr["项目号"] = dr1["项目号"].ToString();
                dr["报价编号"] = dr1["报价编号"].ToString();
                dr["报价"] = dr1["报价"].ToString();
                dr["日期"] = dr1["日期"].ToString();
                dr["部品名"] = dr1["部品名"].ToString();
                dr["报价ID"] = dr1["报价ID"].ToString();
                dr["印刷选项"] = dr1["印刷选项"].ToString();
                dr["模切"] = dr1["模切"].ToString();

                dr["面纸"] = dr1["面纸"].ToString();
                if (!string.IsNullOrEmpty(dr1["面纸克重"].ToString()))
                {
                    dr["面纸克重"] = dr1["面纸克重"].ToString();
                }
                else
                {
                    dr["面纸克重"] = DBNull.Value;
                }
                dr["芯纸"] = dr1["芯纸"].ToString();
                dr["芯纸规格"] = dr1["芯纸规格"].ToString();
                dr["底纸"] = dr1["底纸"].ToString();
                if (!string.IsNullOrEmpty(dr1["底纸克重"].ToString()))
                {
                    dr["底纸克重"] = dr1["底纸克重"].ToString();
                }
                else
                {
                    dr["底纸克重"] = DBNull.Value;
                }
               
                dr["表面加工"] = dr1["表面加工"].ToString();
               
                dr["裱纸工艺"] = dr1["裱纸工艺"].ToString();
            

                if (!string.IsNullOrEmpty(dr1["加工门幅"].ToString()))
                {
                    dr["加工门幅"] = dr1["加工门幅"].ToString();
                }
                else
                {
                    dr["加工门幅"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["加工长度"].ToString()))
                {
                    dr["加工长度"] = dr1["加工长度"].ToString();
                }
                else
                {
                    dr["加工长度"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["部品总数"].ToString()) && bc.yesno(dr1["部品总数"].ToString()) != 0)
                {
                    d1 = decimal.Parse(dr1["部品总数"].ToString());
                    dr["部品总数"] = d1.ToString("0");
                }
                else
                {
                    dr["部品总数"] = DBNull.Value;
                }

                dr["机器型号"] = dr1["机器型号"].ToString();
                if (!string.IsNullOrEmpty(dr1["部品单价"].ToString()) && bc.yesno(dr1["部品单价"].ToString()) != 0)
                {
                    d1 = decimal.Parse(dr1["部品单价"].ToString());
                    dr["部品单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["部品单价"] = DBNull.Value;
                }
  
                if (!string.IsNullOrEmpty(dr1["部品总价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["部品总价"].ToString());
                    dr["部品总价"] = d1.ToString("0");
                }
                else
                {
                    dr["部品总价"] = DBNull.Value;
                }

                if (!string.IsNullOrEmpty(dr1["面纸单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸单价"].ToString());
                    dr["面纸单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["面纸单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸用量"].ToString());
                    dr["面纸用量"] = d1.ToString("0");
                }
                else
                {
                    dr["面纸用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸内耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸内耗"].ToString());
                    dr["面纸内耗"] = d1.ToString("0");
                }
                else
                {
                    dr["面纸内耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸下单"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸下单"].ToString());
                    dr["面纸下单"] = d1.ToString("0");
                }
                else
                {
                    dr["面纸下单"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸外耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸外耗"].ToString());
                    dr["面纸外耗"] = d1.ToString("0");
                }
                else
                {
                    dr["面纸外耗"] = DBNull.Value;
                }
                dr["面纸门幅"] = dr1["面纸门幅"].ToString();
                if (!string.IsNullOrEmpty(dr1["面纸纸长"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸纸长"].ToString());
                    dr["面纸纸长"] = d1.ToString("0");
                }
                else
                {
                    dr["面纸纸长"] = DBNull.Value;
                }
                dr["面纸可用"] = dr1["面纸可用"].ToString();
                if (!string.IsNullOrEmpty(dr1["面纸单个用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸单个用量"].ToString());
                    dr["面纸单个用量"] = d1.ToString("0.000");
                }
                else
                {
                    dr["面纸单个用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["面纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸小计"].ToString());
                    dr["面纸小计"] = d1.ToString("0.00");
                }
                else
                {

                    dr["面纸小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸单价"].ToString()))
                {

                    d1 = decimal.Parse(dr1["芯纸单价"].ToString());
                    dr["芯纸单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["芯纸单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸内耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸内耗"].ToString());
                    dr["芯纸内耗"] = d1.ToString("0");
                }
                else
                {
                    dr["芯纸内耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸用量"].ToString());
                    dr["芯纸用量"] = d1.ToString("0");
                }
                else
                {

                    dr["芯纸用量"] = DBNull.Value;
                }
                dr["芯纸门幅"] = dr1["芯纸门幅"].ToString();
                if (!string.IsNullOrEmpty(dr1["芯纸纸长"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸纸长"].ToString());
                    dr["芯纸纸长"] = d1.ToString();
                }
                else
                {
                    dr["芯纸纸长"] = DBNull.Value;
                }
                dr["芯纸可用"] = dr1["芯纸可用"].ToString();
                if (!string.IsNullOrEmpty(dr1["芯纸单个用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸单个用量"].ToString());
                    dr["芯纸单个用量"] = d1.ToString("0.000");
                }
                else
                {
                    dr["芯纸单个用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["芯纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸小计"].ToString());
                    dr["芯纸小计"] = d1.ToString("0.00");
                }
                else
                {

                    dr["芯纸小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸单价"].ToString());
                    dr["底纸单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["底纸单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸用量"].ToString());
                    dr["底纸用量"] = d1.ToString("0");
                }
                else
                {
                    dr["底纸用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸内耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸内耗"].ToString());
                    dr["底纸内耗"] = d1.ToString("0");
                }
                else
                {

                    dr["底纸内耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸下单"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸下单"].ToString());
                    dr["底纸下单"] = d1.ToString("0");
                }
                else
                {

                    dr["底纸下单"] = DBNull.Value;
                }

                if (!string.IsNullOrEmpty(dr1["底纸外耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸外耗"].ToString());
                    dr["底纸外耗"] = d1.ToString("0");
                }
                else
                {
                    dr["底纸外耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸单个用量"].ToString()))
                {

                    d1 = decimal.Parse(dr1["底纸单个用量"].ToString());
                    dr["底纸单个用量"] = d1.ToString("0.000");
                }
                else
                {
                    dr["底纸单个用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["底纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸小计"].ToString());
                    dr["底纸小计"] = d1.ToString("0");
                }
                else
                {

                    dr["底纸小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["印工单色单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["印工单色单价"].ToString());
                    dr["印工单色单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["印工单色单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["超出单色单张价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["超出单色单张价"].ToString());
                    dr["超出单色单张价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["超出单色单张价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["CTP单张价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["CTP单张价"].ToString());
                    dr["CTP单张价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["CTP单张价"] = DBNull.Value;
                }

                dr["正面色数共计"] = dr1["正面色数共计"].ToString();
                dr["正面CTP张数"] = dr1["正面CTP张数"].ToString();
                if (!string.IsNullOrEmpty(dr1["正面纸张损耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面纸张损耗"].ToString());
                    dr["正面纸张损耗"] = d1.ToString("0");
                }
                else
                {
                    dr["正面纸张损耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正面防晒合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面防晒合计"].ToString());
                    dr["正面防晒合计"] = d1.ToString("0");
                }
                else
                {
                    dr["正面防晒合计"] = DBNull.Value;
                }

                if (!string.IsNullOrEmpty(dr1["正面CTP价计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面CTP价计"].ToString());
                    dr["正面CTP价计"] = d1.ToString("0.00");
                }
                else
                {
                    dr["正面CTP价计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正面印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正面印工合计"].ToString());
                    dr["正面印工合计"] = d1.ToString("0.00");
                }
                else
                {

                    dr["正面印工合计"] = DBNull.Value;
                }

                dr["反面色数共计"] = dr1["反面色数共计"].ToString();
                dr["反面CTP张数"] = dr1["反面CTP张数"].ToString();
                dr["反面纸张损耗"] = dr1["反面纸张损耗"].ToString();
                if (!string.IsNullOrEmpty(dr1["反面防晒合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["反面防晒合计"].ToString());
                    dr["反面防晒合计"] = d1.ToString("0");
                }
                else
                {
                    dr["反面防晒合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["反面CTP价计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["反面CTP价计"].ToString());
                    dr["反面CTP价计"] = d1.ToString("0.00");
                }
                else
                {
                    dr["反面CTP价计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["反面印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["反面印工合计"].ToString());
                    dr["反面印工合计"] = d1.ToString("0.000");
                }
                else
                {
                    dr["反面印工合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正反CTP合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正反CTP合计"].ToString());
                    dr["正反CTP合计"] = d1.ToString("0");
                }
                else
                {
                    dr["正反CTP合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["正反印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正反印工合计"].ToString());
                    dr["正反印工合计"] = d1.ToString("0");
                }
                else
                {
                    dr["正反印工合计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["表面处理单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面处理单价"].ToString());
                    dr["表面处理单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["表面处理单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["无印刷表面处理损耗"].ToString()))
                {
                    d1 = decimal.Parse(dr1["无印刷表面处理损耗"].ToString());
                    dr["无印刷表面处理损耗"] = d1.ToString();
                }
                else
                {

                    dr["无印刷表面处理损耗"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["表面处理用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面处理用量"].ToString());
                    dr["表面处理用量"] = d1.ToString("0");
                }
                else
                {
                    dr["表面处理用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["表面加工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面加工小计"].ToString());
                    dr["表面加工小计"] = d1.ToString("0");
                }
                else
                {

                    dr["表面加工小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["裱工单价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工单价"].ToString());
                    dr["裱工单价"] = d1.ToString("0.00");
                }
                else
                {
                    dr["裱工单价"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["裱工用量"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工用量"].ToString());
                    dr["裱工用量"] = d1.ToString("0");
                }
                else
                {
                    dr["裱工用量"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["裱工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工小计"].ToString());
                    dr["裱工小计"] = d1.ToString("0");
                }
                else
                {

                    dr["裱工小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["刀模小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["刀模小计"].ToString());
                    dr["刀模小计"] = d1.ToString("0");
                }
                else
                {
                    dr["刀模小计"] = DBNull.Value;
                }
                if (!string.IsNullOrEmpty(dr1["模切小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["模切小计"].ToString());
                    dr["模切小计"] = d1.ToString("0");
                }
                else
                {
                    dr["模切小计"] = DBNull.Value;
                }
                dt.Rows.Add(dr);
                i = i + 1;
            }
            if (dt.Rows.Count > 0)
            {
                RETURN_DT_TO_EXCEL_TOTAL_FORM(dt, dtt);
            }

            return dt;
        }
        #endregion
        #region RETURN_DT_TO_EXCEL_TOTAL_FORM
        public DataTable RETURN_DT_TO_EXCEL_TOTAL_FORM(DataTable dt, DataTable dtt)
        {
            DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "面纸可用  IN ('单项总价')");
            d1 = 0;
            foreach (DataRow dr1 in dtt1.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["机器型号"] = "主体";
                dr["面纸可用"] = "单项总价";
                if (!string.IsNullOrEmpty(dr1["面纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸小计"].ToString());
                    dr["面纸小计"] = d1.ToString("0");
                }
                if (!string.IsNullOrEmpty(dr1["芯纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸小计"].ToString());
                    dr["芯纸小计"] = d1.ToString("0");
                }

                if (!string.IsNullOrEmpty(dr1["底纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸小计"].ToString());
                    dr["底纸小计"] = d1.ToString("0");
                }

                if (!string.IsNullOrEmpty(dr1["表面加工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面加工小计"].ToString());
                    dr["表面加工小计"] = d1.ToString("0");
                }

                if (!string.IsNullOrEmpty(dr1["裱工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工小计"].ToString());
                    dr["裱工小计"] = d1.ToString("0");
                }
                if (!string.IsNullOrEmpty(dr1["模切小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["模切小计"].ToString());
                    dr["模切小计"] = d1.ToString("0");
                }
                if (!string.IsNullOrEmpty(dr1["正反印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正反印工合计"].ToString());
                    dr["正反印工合计"] = d1.ToString("0");
                }
                if (!string.IsNullOrEmpty(dr1["部品总价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["部品总价"].ToString());
                    dr["部品总价"] = d1.ToString("0");
                }
                dt.Rows.Add(dr);

            }
            DataTable dtt2 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "面纸可用  IN ('单项单价')");
            d1 = 0;
            foreach (DataRow dr1 in dtt2.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["机器型号"] = "汇总";
                dr["面纸可用"] = "单项单价";
                if (!string.IsNullOrEmpty(dr1["面纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["面纸小计"].ToString());
                    dr["面纸小计"] = d1.ToString("0.00");
                }
                if (!string.IsNullOrEmpty(dr1["芯纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["芯纸小计"].ToString());
                    dr["芯纸小计"] = d1.ToString("0.00");
                }

                if (!string.IsNullOrEmpty(dr1["底纸小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["底纸小计"].ToString());
                    dr["底纸小计"] = d1.ToString("0.00");
                }

                if (!string.IsNullOrEmpty(dr1["表面加工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["表面加工小计"].ToString());
                    dr["表面加工小计"] = d1.ToString("0.00");
                }

                if (!string.IsNullOrEmpty(dr1["裱工小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["裱工小计"].ToString());
                    dr["裱工小计"] = d1.ToString("0.00");
                }
                if (!string.IsNullOrEmpty(dr1["模切小计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["模切小计"].ToString());
                    dr["模切小计"] = d1.ToString("0.00");
                }
                if (!string.IsNullOrEmpty(dr1["正反印工合计"].ToString()))
                {
                    d1 = decimal.Parse(dr1["正反印工合计"].ToString());
                    dr["正反印工合计"] = d1.ToString("0.00");

                }
                if (!string.IsNullOrEmpty(dr1["部品总价"].ToString()))
                {
                    d1 = decimal.Parse(dr1["部品总价"].ToString());
                    dr["部品总价"] = d1.ToString("0.00");
                }
                dt.Rows.Add(dr);

            }
            return dt;

        }
        #endregion
        #region RETURN_COST_TOTAL_DT_FORM
        public DataTable RETURN_COST_TOTAL_DT_FORM(DataTable dtt)
        {
            DataTable dt = emptydatatable();
            DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "项目 NOT IN ('无外购采购比')");
            int i = 1;
            foreach (DataRow dr1 in dtt1.Rows)
            {
                DataRow dr = dt.NewRow();
                if (i == dtt1.Rows.Count - 1 || i == dtt1.Rows.Count)
                {
                }
                else
                {
                    dr["序号"] = i.ToString();
                }
                dr["项目"] = dr1["项目"].ToString();
                if (!string.IsNullOrEmpty(dr1["元套"].ToString()) && bc.yesno(dr1["元套"].ToString()) != 0)
                {
                    d1 = decimal.Parse(dr1["元套"].ToString());
                    dr["元套"] = d1.ToString("0.00");
                  
                }
                if (!string.IsNullOrEmpty(dr1["批量小计"].ToString()) && bc.yesno(dr1["批量小计"].ToString()) != 0)
                {
                    d1 = decimal.Parse(dr1["批量小计"].ToString());
                    dr["批量小计"] = d1.ToString("0");
                    
                }
                dr["主件用量"] = dr1["主件用量"].ToString();
                dt.Rows.Add(dr);
                i = i + 1;
            }
            DataTable dtt2 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "项目 IN ('无外购采购比')");//此为最后一行的数据
            foreach (DataRow dr1 in dtt2.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["项目"] = dr1["项目"].ToString();
                dr["批量小计"] = dr1["批量小计"].ToString();//此单元格值为文字“外购采购比“160131
                if (!string.IsNullOrEmpty(dr1["元套"].ToString()))
                {
                    d1 = decimal.Parse(bc.RETURN_UNTIL_CHAR (dr1["元套"].ToString(),'%'));
                    dr["元套"] = d1.ToString("0.0") +'%';

                }
                /*sqb = new StringBuilder();
                sqb.AppendFormat("项目：{0},", dr1["项目"].ToString());
                sqb.AppendFormat("元套：{0},", dr1["元套"].ToString());
                sqb.AppendFormat("批量小计：{0},", dr1["批量小计"].ToString());
                sqb.AppendFormat("主件用量：{0},", dr1["主件用量"].ToString());
                MessageBox.Show(sqb.ToString());*/

                if (!string.IsNullOrEmpty(dr1["主件用量"].ToString()))
                {
                    d1 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dr1["主件用量"].ToString(), '%'));
                    dr["主件用量"] = d1.ToString("0.0")+'%' ;

                }
              
                dt.Rows.Add(dr);
            }
          
        
            return dt;

        }
        #endregion
        #region RETURN_IF_EXISTS_ABOVE_PROJECT
        public DataTable RETURN_IF_EXISTS_ABOVE_PROJECT(DataTable dt)
        {

            DataTable dtt = new DataTable();
            dtt.Columns.Add("报价编号", typeof(string));
            if (dt.Rows.Count > 0)
            {
                dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "报价编号 IS NOT NULL");
                foreach (DataRow dr1 in dt.Rows)
                {
                    DataTable dtt1 = bc.GET_DT_TO_DV_TO_DT(dtt, "", string.Format("报价编号='{0}'", dr1["报价编号"].ToString()));
                    if (dtt1.Rows.Count > 0)
                    {

                    }
                    else
                    {
                        DataRow dr = dtt.NewRow();
                        dr["报价编号"] = dr1["报价编号"].ToString();
                        dtt.Rows.Add(dr);
                    }
                }
            }

            return dtt;
        }
        #endregion
        #region RETURN_HAVE_ID_DT
        public DataTable RETURN_HAVE_ID_DT(DataTable dtx)
        {
            DataTable dt = GetTableInfo_show_all();
            int i = 1;
            foreach (DataRow dr1 in dtx.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["序号"] = i.ToString();
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["数量"] = dr1["数量"].ToString();
                dr["项目号"] = dr1["项目号"].ToString();
                dr["报价编号"] = dr1["报价编号"].ToString();
                dr["报价"] = dr1["报价"].ToString();
                dr["日期"] = dr1["日期"].ToString();
                dr["部品名"] = dr1["部品名"].ToString();
                dr["加工门幅"] = dr1["加工门幅"].ToString();
                dr["加工长度"] = dr1["加工长度"].ToString();
                dr["部品总数"] = dr1["部品总数"].ToString();
                dr["机器型号"] = dr1["机器型号"].ToString();
                dr["部品单价"] = dr1["部品单价"].ToString();
                dr["部品总价"] = dr1["部品总价"].ToString();
                dr["面纸单价"] = dr1["面纸单价"].ToString();
                dr["面纸用量"] = dr1["面纸用量"].ToString();
                dr["面纸内耗"] = dr1["面纸内耗"].ToString();
                dr["面纸下单"] = dr1["面纸下单"].ToString();
                dr["面纸外耗"] = dr1["面纸外耗"].ToString();
                dr["面纸门幅"] = dr1["面纸门幅"].ToString();
                dr["面纸纸长"] = dr1["面纸纸长"].ToString();
                dr["面纸可用"] = dr1["面纸可用"].ToString();
                dr["面纸单个用量"] = dr1["面纸单个用量"].ToString();
                dr["面纸小计"] = dr1["面纸小计"].ToString();
                dr["芯纸单价"] = dr1["芯纸单价"].ToString();
                dr["芯纸内耗"] = dr1["芯纸内耗"].ToString();
                dr["芯纸用量"] = dr1["芯纸用量"].ToString();
                dr["芯纸门幅"] = dr1["芯纸门幅"].ToString();
                dr["芯纸纸长"] = dr1["芯纸纸长"].ToString();
                dr["芯纸可用"] = dr1["芯纸可用"].ToString();
                dr["芯纸单个用量"] = dr1["芯纸单个用量"].ToString();
                dr["芯纸小计"] = dr1["芯纸小计"].ToString();
                dr["底纸单价"] = dr1["底纸单价"].ToString();
                dr["底纸用量"] = dr1["底纸用量"].ToString();
                dr["底纸内耗"] = dr1["底纸内耗"].ToString();
                dr["底纸下单"] = dr1["底纸下单"].ToString();
                dr["底纸外耗"] = dr1["底纸外耗"].ToString();
                dr["底纸单个用量"] = dr1["底纸单个用量"].ToString();
                dr["底纸小计"] = dr1["底纸小计"].ToString();
                dr["印工单色单价"] = dr1["印工单色单价"].ToString();
                dr["超出单色单张价"] = dr1["超出单色单张价"].ToString();
                dr["CTP单张价"] = dr1["CTP单张价"].ToString();
                dr["正面色数共计"] = dr1["正面色数共计"].ToString();
                dr["正面CTP张数"] = dr1["正面CTP张数"].ToString();
                dr["正面纸张损耗"] = dr1["正面纸张损耗"].ToString();
                dr["正面防晒合计"] = dr1["正面防晒合计"].ToString();
                dr["正面CTP价计"] = dr1["正面CTP价计"].ToString();
                dr["正面印工合计"] = dr1["正面印工合计"].ToString();
                dr["反面色数共计"] = dr1["反面色数共计"].ToString();
                dr["反面CTP张数"] = dr1["反面CTP张数"].ToString();
                dr["反面纸张损耗"] = dr1["反面纸张损耗"].ToString();
                dr["反面防晒合计"] = dr1["反面防晒合计"].ToString();
                dr["反面CTP价计"] = dr1["反面CTP价计"].ToString();
                dr["反面印工合计"] = dr1["反面印工合计"].ToString();
                dr["正反CTP合计"] = dr1["正反CTP合计"].ToString();
                dr["正反印工合计"] = dr1["正反印工合计"].ToString();
                dr["表面处理单价"] = dr1["表面处理单价"].ToString();
                dr["无印刷表面处理损耗"] = dr1["无印刷表面处理损耗"].ToString();
                dr["表面处理用量"] = dr1["表面处理用量"].ToString();
                dr["表面加工小计"] = dr1["表面加工小计"].ToString();
                dr["裱工单价"] = dr1["裱工单价"].ToString();
                dr["裱工用量"] = dr1["裱工用量"].ToString();
                dr["裱工小计"] = dr1["裱工小计"].ToString();
                dr["刀模小计"] = dr1["刀模小计"].ToString();
                dr["模切小计"] = dr1["模切小计"].ToString();

                dt.Rows.Add(dr);
                i = i + 1;
            }
            return dt;
        }
        #endregion
        #region RETURN_SEARCH
        public DataTable RETURN_SEARCH(DataTable dtt)
        {
            DataTable dt = GetTableInfo_search();
            DataTable dtx = bc.RETURN_NOHAVE_REPEAT_DT(dtt, "项目号");
            DataTable dtx4=new DataTable ();
            i = 0;
            if (dtx.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dtx.Rows)
                {
                    bool b1 = false;
                    bool b2 = false;
                    DataTable dtx1 = bc.GET_DT_TO_DV_TO_DT(bc.getdt(sqlfi + " ORDER BY A.SAMPLE_ID ASC"), "", "项目号='" + dr1["VALUE"].ToString() + "'");//打样单号
                    DataTable dtx2 = bc.GET_DT_TO_DV_TO_DT(bc.getdt(sqlsi), "", "项目号='" + dr1["VALUE"].ToString() + "'");//报价编号
                    DataTable dtx3 = bc.GET_DT_TO_DV_TO_DT(dtt, "", "项目号='" + dr1["VALUE"].ToString() + "'");
                    if (dtx1.Rows.Count > 0)
                    {
                        b1 = true;
                    }
                    if (dtx2.Rows.Count > 0)
                    {
                        b2 = true;
                    }
                    if (b1 == true && b2 == false)
                    {
                        if (dtx1.Rows.Count > 0)
                        {
                            foreach (DataRow dr2 in dtx1.Rows)
                            {
                                DataRow dr = dt.NewRow();
                                dr["序号"] = i;
                                dr["项目名称"] = dr2["项目名称"].ToString();
                                dr["项目号"] = dr2["项目号"].ToString();
                                dr["打样单号"] =dr2["打样单号"].ToString();
                                dr["客户"] = dr2["客户"].ToString();
                                dr["品牌"] = dr2["品牌"].ToString();
                                dr["AE"] = dr2["AE"].ToString();
                                dtx4 = csample_rely_list.RETURN_DT(bc.getdt(csample_rely_list.sql + " WHERE A.SAMPLE_ID='" + dr2["打样单号"].ToString() +
                                    "'"));

                                dr["打样金额"] = dtx4.Rows[0]["样板计费"].ToString();
                                dt.Rows.Add(dr);
                                i = i + 1;
                            }
                        }
                    }
                    else if (b1 == false && b2 == true)
                    {
                        if (dtx2.Rows.Count > 0)
                        {
                            foreach (DataRow dr2 in dtx2.Rows)
                            {
                                DataRow dr = dt.NewRow();
                                dr["序号"] = i;
                                dr["项目名称"] = dr2["项目名称"].ToString();
                                dr["项目号"] = dr2["项目号"].ToString();
                                dr["报价编号"] = dr2["报价编号"].ToString();
                                dr["客户"] = dr2["客户"].ToString();
                                dr["品牌"] = dr2["品牌"].ToString();
                                dr["AE"] = dr2["AE"].ToString();
                                dr["报价数量"] = dr2["报价数量"].ToString();
                                dr["审核批注"] = dr2["审核批注"].ToString();
                                dr["制单人"] = dr2["制单人"].ToString();
                                dr["制单日期"] = dr2["制单日期"].ToString();
                                dr["修改日期"] = dr2["修改日期"].ToString();
                                dt.Rows.Add(dr);
                                i = i + 1;
                            }
                        }
                    }
                    else if (b1 == true && b2 == true)
                    {
                       
                        if (dtx1.Rows.Count >= dtx2.Rows.Count)
                        {
                            int[] array1 = new int[dtx1.Rows.Count];
                            int i1 = 0;
                            //MessageBox.Show("ok");
                            foreach (DataRow dr2 in dtx1.Rows)
                            {
                                DataRow dr = dt.NewRow();
                                dr["序号"] = i;
                                dr["项目名称"] = dr2["项目名称"].ToString();
                                dr["项目号"] = dr2["项目号"].ToString();
                                dr["打样单号"] = dr2["打样单号"].ToString();
                                dr["客户"] = dr2["客户"].ToString();
                                dr["品牌"] = dr2["品牌"].ToString();
                                dr["AE"] = dr2["AE"].ToString();
                                dtx4 = csample_rely_list.RETURN_DT(bc.getdt(csample_rely_list.sql + " WHERE A.SAMPLE_ID='" + dr2["打样单号"].ToString() +
                                    "'"));

                                dr["打样金额"] = dtx4.Rows[0]["样板计费"].ToString();
                                dt.Rows.Add(dr);
                                array1[i1] = i;/*用数组记住需要更新的序号*/
                                i = i + 1;
                                i1 += 1;
                            }
                            for (int j = 0; j < dtx2.Rows.Count; j++)
                            {
                                dt.Rows[array1[j]]["报价编号"] = dtx2.Rows[j]["报价编号"].ToString();
                                dt.Rows[array1[j]]["报价数量"] = dtx2.Rows[j]["报价数量"].ToString();
                                dt.Rows[array1[j]]["审核批注"] = dtx2.Rows[j]["审核批注"].ToString();
                                dt.Rows[array1[j]]["制单人"] = dtx2.Rows[j]["制单人"].ToString();
                                dt.Rows[array1[j]]["制单日期"] = dtx2.Rows[j]["制单日期"].ToString();
                                dt.Rows[array1[j]]["修改日期"] = dtx2.Rows[j]["修改日期"].ToString();
                            }
                        }
                 
                        else
                        {
                            int[] array1 = new int[dtx2.Rows.Count];
                            int i1 = 0;
                            foreach (DataRow dr2 in dtx2.Rows)
                            {
                                DataRow dr = dt.NewRow();
                                dr["序号"] = i;
                                dr["项目名称"] = dr2["项目名称"].ToString();
                                dr["项目号"] = dr2["项目号"].ToString();
                                dr["报价编号"] = dr2["报价编号"].ToString();
                                dr["报价数量"] = dr2["报价数量"].ToString();
                                dr["客户"] = dr2["客户"].ToString();
                                dr["品牌"] = dr2["品牌"].ToString();
                                dr["AE"] = dr2["AE"].ToString();
                                dr["审核批注"] = dr2["审核批注"].ToString();
                                dr["制单人"] = dr2["制单人"].ToString();
                                dr["制单日期"] = dr2["制单日期"].ToString();
                                dr["修改日期"] = dr2["修改日期"].ToString();
                                dt.Rows.Add(dr);
                                array1[i1] = i;
                                i = i + 1;
                                i1 += 1;
                            }
                            for (int j = 0; j < dtx1.Rows.Count; j++)
                            {

                                dt.Rows[array1[j]]["打样单号"] = dtx1.Rows[j]["打样单号"].ToString();
                                dtx4 = csample_rely_list.RETURN_DT(bc.getdt(csample_rely_list.sql + @" 
WHERE A.SAMPLE_ID='" + dtx1.Rows[j]["打样单号"].ToString() +
                               "'"));

                                dt.Rows[array1[j]]["打样金额"] = dtx4.Rows[0]["样板计费"].ToString();
                            }
                        }
                    }
             
                }

                foreach (DataRow dr in dt.Rows)
                {
                    d1 = 0;
                    if (dr["报价编号"].ToString() == "")
                    {
                    }
                    else
                    {
                        dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.OFFER_ID='" + dr["报价编号"].ToString() + "'");
                        if (dtx.Rows.Count > 0 && dtx.Rows .Count ==20)//发现行数有不为20的费用TABLE PRINT_COST_TOTAL 需先加这个宽屏判断条件
                        {
                            //MessageBox.Show(dr["报价编号"].ToString());
                            if (!string.IsNullOrEmpty(dtx.Rows[dtx.Rows .Count -2]["元套"].ToString()))
                            {
                                d1 = decimal.Parse(dtx.Rows[dtx.Rows .Count -2]["元套"].ToString());
                                dr["报出价"] = d1.ToString("0.00");
                            }
                            
                        }
                    }
                   
                }
            }
            return dt;
        }
        #endregion
        #region RETURN_SEARCH_HAVE_ID_DT
        public DataTable RETURN_SEARCH_HAVE_ID_DT(DataTable dtx)
        {
            DataTable dt = GetTableInfo_search();
            int i = 1;
            foreach (DataRow dr1 in dtx.Rows)
            {
                DataRow dr = dt.NewRow();
                dr["序号"] = i.ToString();
                dr["项目号"] = dr1["项目号"].ToString();
                dr["项目名称"] = dr1["项目名称"].ToString();
                dr["客户"] = dr1["客户"].ToString();
                dr["品牌"] = dr1["品牌"].ToString();
                dr["AE"] = dr1["AE"].ToString();
                dr["打样单号"] = dr1["打样单号"].ToString();
                dr["打样金额"] = dr1["打样金额"].ToString();
                dr["报价编号"] = dr1["报价编号"].ToString();
                dr["报价数量"] = dr1["报价数量"].ToString();
                dr["报出价"] = dr1["报出价"].ToString();
                dr["审核批注"] = dr1["审核批注"].ToString();
                dr["制单人"] = dr1["制单人"].ToString();
                dr["制单日期"] = dr1["制单日期"].ToString();
                dr["修改日期"] = dr1["修改日期"].ToString();
                dt.Rows.Add(dr);
                i = i + 1;
            }
            return dt;
        }
        #endregion
        #region RETURN_OFFER_PRICE
        public decimal  RETURN_OFFER_PRICE(string OFFER_ID)
        {
            decimal d1 = 0;
            DataTable dt = RETURN_SEARCH(bc.getdt(sqlse + " WHERE B.OFFER_ID='"+OFFER_ID +"'"));
            if (dt.Rows.Count > 0)
            {
                d1 = decimal.Parse(dt.Rows[0]["报出价"].ToString());
            }
            return d1;
        }
        #endregion
        #region emptydatatable
        public DataTable emptydatatable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("项目", typeof(string));
            dt.Columns.Add("元套", typeof(string));
            dt.Columns.Add("批量小计", typeof(string));
            dt.Columns.Add("主件用量", typeof(string));
            return dt;
        }
        #endregion
        #region RETURN_COST_TOTAL_DT
        public DataTable RETURN_COST_TOTAL_DT(string PFID, DataTable dtt)
        {
            //MessageBox.Show("4");
            decimal de1 = 0, de2 = 0, de3 = 0, de4 = 0, de5 = 0, de6 = 0, de7 = 0, de8 = 0, de9 = 0, de10 = 0, de11 = 0, de12 = 0,de13_a=0,de14_a=0,de15_a=0,de16=0;
            DataTable dt = emptydatatable();
            List<string> list = new List<string>();
            list.Add("纸张");//0
            list.Add("芯纸");//1
            list.Add("印刷");//2
            list.Add("表面");//3
            list.Add("裱纸");//4
            list.Add("模切");//5
            list.Add("刀模");//6

            list.Add("写真");//7
            list.Add("配件");//8
            list.Add("包装");//9
            list.Add("运输");//10
            list.Add("人工");//11

            list.Add("成本合计");//12
            list.Add("管理");//13
            list.Add("利润");//14
            list.Add("外购件");//15
            list.Add("代购费");//16
            /*用于费用汇总表修改百分比时产生公式计算所需数据 START*/
            if (RETURN_BATCH_SUBTOTAL_COST_SET != 0)
            {
                dtt = bc.getdt(sql + " WHERE B.PFID='" +PFID + "'");
                if (dtt.Rows.Count > 0)
                {
                    dtt =RETURN_DT(dtt);
                    dtt =bind2(dtt, 0, "");
                    dtt = RETURN_DT_SHOW_HIDE(dtt);
                }
            }
            /*用于费用汇总表修改百分比时产生公式计算所需数据 END*/
            COUNT = 0;

            if (!string.IsNullOrEmpty(dtt.Rows[0]["数量"].ToString()))
            {
                COUNT = decimal.Parse(dtt.Rows[0]["数量"].ToString());
            }
            for (int j = 0; j < list.Count; j++)
            {
                DataRow dr = dt.NewRow();
                dr["序号"] = j + 1;
                dr["项目"] = list[j].ToString();
                dt.Rows.Add(dr);
            }
            //MessageBox.Show("5");
            if (dtt.Rows.Count > 0)
            {

                dt.Rows[0]["元套"] = 0;
                dt.Rows[1]["元套"] = 0;
                dt.Rows[2]["元套"] = 0;
                dt.Rows[3]["元套"] = 0;
                dt.Rows[4]["元套"] = 0;
                dt.Rows[5]["元套"] = 0;
                if (!string.IsNullOrEmpty(dtt.Rows[0]["数量"].ToString()))
                {
                    COUNT = decimal.Parse(dtt.Rows[0]["数量"].ToString());
                }
                //MessageBox.Show(this.TOTAL_COST_TISSUE + ","+this.TOTAL_COST_PAPAER_CORE +","+ this.TOTAL_COST_BODY_PAPER);
                d1 = this.TOTAL_COST_TISSUE + this.TOTAL_COST_BODY_PAPER;

                if (d1 != 0)
                {
                    dt.Rows[0]["批量小计"] = d1.ToString();
                    d2 = d1 / COUNT;
                    dt.Rows[0]["元套"] = d2.ToString();
                    de1 = d2;
                }

                d1 = this.TOTAL_COST_PAPAER_CORE;
                if (d1 != 0)
                {
                    dt.Rows[1]["批量小计"] = d1.ToString();
                    d2 = d1 / COUNT;
                    dt.Rows[1]["元套"] = d2.ToString();
                    de2 = d2;
                }
      
                d1 = this.TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP;
                if (this.TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP != 0)
                {
                    //MessageBox.Show(TOTAL_COST_POSITIVE_AND_OPPOSITE_PRINTING_AND_CTP.ToString ()+"total");
                    dt.Rows[2]["批量小计"] = d1.ToString();
                    d2 = d1 / COUNT;
                    dt.Rows[2]["元套"] = d2.ToString();
                    de3 = d2;
                }

                d1 = this.TOTAL_COST_SURFACE_PROCESSING;
                if (d1 != 0)
                {
                    dt.Rows[3]["批量小计"] = d1.ToString();
                    d2 = d1 / COUNT;
                    dt.Rows[3]["元套"] = d2.ToString();
                    de4 = d2;
                }
                d1 = this.TOTAL_COST_LAMINATING_PROCESS;
                if (d1 != 0)
                {
                    dt.Rows[4]["批量小计"] = d1.ToString();
                    d2 = d1 / COUNT;
                    dt.Rows[4]["元套"] = d2.ToString();
                    de5 = d2;
                }
                d1 = this.TOTAL_COST_DIE_CUTTING;
                if (d1 != 0)
                {
                    dt.Rows[5]["批量小计"] = d1.ToString();
                    d2 = d1 / COUNT;
                    dt.Rows[5]["元套"] = d2.ToString();
                    de6 = d2;
                }
            }
            //MessageBox.Show("51");
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "主件用量");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[0]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }
            dt.Rows[1]["主件用量"] = "主件单价";
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "主件单价");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[2]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }
            dt.Rows[3]["主件用量"] = "辅材单价";
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "辅材单价");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[4]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }
            dt.Rows[5]["主件用量"] = "报价管理";
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "报价管理");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[6]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }

            dt.Rows[7]["主件用量"] = "报价利润";
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "报价利润");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[8]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "管理");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[13]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }

            if (RECEPTION_USE)
            {
                dt.Rows[13]["主件用量"] = MAIN_MANAGE;

            }
            //MessageBox.Show("52");
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "利润");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[14]["主件用量"] = dtx.Rows[0]["客户比例"].ToString();
            }
            if (RECEPTION_USE)
            {
                dt.Rows[14]["主件用量"] = MAIN_PROFIT;
            }
            //MessageBox.Show("55");
            dtx = bc.getdt(cprint_die_cutting.sql + " WHERE A.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[6]["批量小计"] = dtx.Rows[2]["圆孔个数"].ToString();
                dt.Rows[6]["元套"] = dtx.Rows[2]["小计"].ToString();
                if (!string.IsNullOrEmpty(dtx.Rows[2]["小计"].ToString()))
                {
                    de7 = decimal.Parse(dtx.Rows[2]["小计"].ToString());
                }
            }
            //MessageBox.Show("56");
            dtx = bc.getdt(cprint_portray.sql + " WHERE A.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(dtx.Compute("SUM(小计)", "").ToString()))
                {
                    dt.Rows[7]["批量小计"] = dtx.Compute("SUM(小计)", "");
                }
                dt.Rows[7]["元套"] = dtx.Rows[dtx.Rows.Count - 1]["单价"].ToString();

                if (!string.IsNullOrEmpty(dtx.Rows[dtx.Rows.Count - 1]["单价"].ToString()))
                {
                    de8 = decimal.Parse(dtx.Rows[dtx.Rows.Count - 1]["单价"].ToString());
                }
            }
            //MessageBox.Show("54");
            RETURN_PARTS_AUXILIAR_DT(PFID, dtt);
            if (TOTAL_COST_PARTS_AUXILIARY != 0)
            {
                dt.Rows[8]["批量小计"] = (TOTAL_COST_PARTS_AUXILIARY * COUNT).ToString();
                dt.Rows[8]["元套"] = TOTAL_COST_PARTS_AUXILIARY.ToString();
                de9 = TOTAL_COST_PARTS_AUXILIARY;

            }
            else
            {
                dt.Rows[8]["元套"] = 0;

            }
            //
            //MessageBox.Show("53");
            RETURN_PACK_MATERIAL_DT(PFID, dtt);
            if (TOTAL_COST_PACK_MATERIAL != 0)
            {
                dt.Rows[9]["批量小计"] = (TOTAL_COST_PACK_MATERIAL * COUNT).ToString();
                dt.Rows[9]["元套"] = TOTAL_COST_PACK_MATERIAL.ToString();
                de10 = TOTAL_COST_PACK_MATERIAL;
            }
            else
            {
                dt.Rows[9]["元套"] = 0;
            }//
            //MessageBox.Show("1");
            dtx = bc.getdt(cprint_transport.sql + " WHERE A.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
                //MessageBox.Show("2");
                if (!string.IsNullOrEmpty(dtx.Compute("SUM(小计)", "").ToString()))
                {
                    dt.Rows[10]["批量小计"] = decimal.Parse(dtx.Compute("SUM(小计)", "").ToString()).ToString("0");
                    de11 = decimal.Parse(dtx.Compute("SUM(小计)", "").ToString()) / COUNT;
                }
                dt.Rows[10]["元套"] = de11.ToString();
                //MessageBox.Show("3");
            }
            else
            {
                dt.Rows[10]["元套"] = 0;
            }

            RETURN_ARTIFICIAL_DT(PFID, dtt);
            if (TOTAL_COST_ARTIFICIAL != 0)
            {
                dt.Rows[11]["批量小计"] = (TOTAL_COST_ARTIFICIAL * COUNT).ToString();
                dt.Rows[11]["元套"] = TOTAL_COST_ARTIFICIAL.ToString();
                de12 = TOTAL_COST_ARTIFICIAL;
            }
            else
            {
                dt.Rows[11]["元套"] = 0;
            }
            d1 = 0;
            d2 = 0;
            d4 = 0;
            de12 = de1 + de2 + de3 + de4 + de5 + de6 + de7 + de8 + de9 + de10 + de11 + de12;
            dt.Rows[12]["批量小计"] = (de12 * COUNT).ToString();
            de13_a = de12 * COUNT;
            dt.Rows[12]["元套"] = (de12).ToString();

            if (!string.IsNullOrEmpty(dt.Rows[13]["主件用量"].ToString()))
            {
                d1 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dt.Rows[13]["主件用量"].ToString(), '%')) / 100;//
            }
            if (!string.IsNullOrEmpty(dt.Rows[14]["主件用量"].ToString()))
            {
                d2 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dt.Rows[14]["主件用量"].ToString(), '%')) / 100;//
            }
            double do1 = 0,do2=0,do3=0,do4=0,do5=0;
            do1 = Convert.ToDouble(d1);
            do2 = Convert.ToDouble(d2);
            do3 = Convert.ToDouble(de12);
            do4 = Convert.ToDouble(COUNT);
            do5 = ((do3 * do4) / (1 - do1 - do2 - (11.5 * 1 / 100)) * do1);//160119 根据需求修改公式
            de14_a = decimal.Parse(do5.ToString());//160119 根据需求修改公式
            dt.Rows[13]["批量小计"] = de14_a;//160119 根据需求修改公式
            RETURN_BATCH_SUBTOTAL_MANAGE = 0;
            if (!string.IsNullOrEmpty(dt.Rows[13]["批量小计"].ToString()))
            {
                RETURN_BATCH_SUBTOTAL_MANAGE = decimal.Parse(dt.Rows[13]["批量小计"].ToString());
            }
            dt.Rows[13]["元套"] = de14_a /COUNT ;//16/01/19 根据需求修改公式
            YUAN_SET_MANAGE = de14_a / COUNT;//16/01/19 根据需求修改公式
            /*sqb = new StringBuilder();
            sqb.AppendFormat("de1={0},", de1 * COUNT);
            sqb.AppendFormat("de2={0},", de2 * COUNT);
            sqb.AppendFormat("de3={0},", de3 * COUNT);
            sqb.AppendFormat("de4={0},", de4 * COUNT);
            sqb.AppendFormat("de5={0},", de5 * COUNT);
            sqb.AppendFormat("de6={0},", de6 * COUNT);
            sqb.AppendFormat("de7={0},", de7 * COUNT);
            sqb.AppendFormat("de8={0},", de8 * COUNT);
            sqb.AppendFormat("de9={0},", de9 * COUNT);
            sqb.AppendFormat("de10={0},", de10 * COUNT);
            sqb.AppendFormat("de11={0},", de11* COUNT);
            sqb.AppendFormat("de12={0},", de12 * COUNT);
            sqb.AppendFormat("de14改后={0},", de14_a );
            sqb.AppendFormat("百分比={0},", do5);
            //MessageBox.Show(sqb.ToString());*/
            de15_a = decimal .Parse (((do3* do4) / (1 - do1 - do2- (11.5 * 1 / 100)) * do2).ToString ());//16/01/19 根据需求修改公式
            dt.Rows[14]["批量小计"] = de15_a;//16/01/19 根据需求修改公式
            RETURN_MAIN_DOSAGE_PROFIT = de15_a;//16/01/19 根据需求修改公式
            dt.Rows[14]["元套"] = de15_a /COUNT ;//16/01/19 根据需求修改公式
            YUAN_SET_PROFIT =de15_a /COUNT ;//16/01/19 根据需求修改公式
       
            dt.Rows[15]["主件用量"] = "代购费率";
            dtx = bc.getdt(cprint_purchase.sql + " WHERE A.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
                dt.Rows[15]["批量小计"] = decimal.Parse(dtx.Rows[dtx.Rows .Count -1]["管理费二"].ToString()) * COUNT;
                dt.Rows[15]["元套"] = dtx.Rows[dtx.Rows .Count -1]["管理费二"].ToString();
                TOTAL_COST_PURCHASE = decimal.Parse(dtx.Rows[dtx.Rows .Count -1]["管理费二"].ToString());
                de16 = decimal.Parse(dtx.Rows[dtx.Rows .Count -1]["管理费二"].ToString());
            }
            else
            {
                dt.Rows[15]["元套"] = 0;
            }
            if (RECEPTION_USE && MAIN_PURCHASE != 0)
            {

                dt.Rows[16]["主件用量"] = MAIN_PURCHASE;
            }
            d3 = 0;
            if (!string.IsNullOrEmpty(dt.Rows[16]["主件用量"].ToString()))
            {
                d3 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dt.Rows[16]["主件用量"].ToString(), '%')) / 100;//
            }
            RETURN_MAIN_DOSAGE_PURCHASE_COST = 0;//没有赋值的属性在使用前先初始化，不然多个报价编号调用时值被错误使用 15/12/19
            d4 = 0;
            d5 = 0;
            d6 = 0;
            d7 = 0;
            if (!string.IsNullOrEmpty(dtx.Rows[0]["管理费一"].ToString()))
            {
                d4 = decimal.Parse(dtx.Rows[0]["管理费一"].ToString());
            }
            if (!string.IsNullOrEmpty(dtx.Rows[0]["管理费二"].ToString()))
            {
                d5 = decimal.Parse(dtx.Rows[0]["管理费二"].ToString());
            }
            if (!string.IsNullOrEmpty(dtx.Rows[1]["管理费一"].ToString()))
            {
                d6 = decimal.Parse(dtx.Rows[1]["管理费一"].ToString());
            }
            d7 = d4 + d5 + d6;
            if (TOTAL_COST_PURCHASE == 0)
            {


            }
            else if (dt.Rows[16]["主件用量"].ToString() == "")
            {
                dt.Rows[16]["批量小计"] = (d7 * COUNT).ToString();
                RETURN_MAIN_DOSAGE_PURCHASE_COST = (d7) * COUNT;
            }
            else
            {
                if (!string.IsNullOrEmpty(dt.Rows[15]["批量小计"].ToString()))
                {
                    d7 = decimal.Parse(dt.Rows[15]["批量小计"].ToString ());
                }
                dt.Rows[16]["批量小计"] = (d7 / (1 - d3) * d3).ToString();
                RETURN_MAIN_DOSAGE_PURCHASE_COST = (d7 / (1 - d3) * d3);
            }
            /*sqb = new StringBuilder();
            sqb.AppendFormat("d3={0},", d3);
            sqb.AppendFormat("d5={0},", d5);
            sqb.AppendFormat("d6={0},", d6);
            sqb.AppendFormat("d7={0},", d7);
            MessageBox.Show(sqb.ToString());*/
            dt.Rows[16]["元套"] = (RETURN_MAIN_DOSAGE_PURCHASE_COST / COUNT).ToString();
            RETURN_YUAN_SET_PURCHASE_COST = (RETURN_MAIN_DOSAGE_PURCHASE_COST / COUNT);
            d1 = 0;
            d2 = 0;
            d3 = 0;
            d4 = 0;
            d5 = 0;
            d6 = 0;
 
            if (!string.IsNullOrEmpty(dt.Rows[12]["元套"].ToString()))
            {
                d1 = decimal.Parse(dt.Rows[12]["元套"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[13]["元套"].ToString()))
            {
                d2 = decimal.Parse(dt.Rows[13]["元套"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[14]["元套"].ToString()))
            {
                d3 = decimal.Parse(dt.Rows[14]["元套"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[15]["元套"].ToString()))
            {
                d4 = decimal.Parse(dt.Rows[15]["元套"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[16]["元套"].ToString()))
            {
                d5 = decimal.Parse(dt.Rows[16]["元套"].ToString());

            }
            DataRow dr1 = dt.NewRow();
            dr1["项目"] = "未税";
        
                dr1["元套"] = (d1 + d2 + d3 + d4 + d5).ToString();
                YUAN_SET_NO_TAX = (d1 + d2 + d3 + d4 + d5);
          
            dr1["批量小计"] = ((d1 + d2 + d3 + d4 + d5) * COUNT).ToString();
            BATCH_TOTAL_NO_TAX = ((d1 + d2 + d3 + d4 + d5) * COUNT);
            dt.Rows.Add(dr1);
            d10 = 0;
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql + " WHERE A.PROJECT_NAME='{0}'", "税金");
            sqb.AppendFormat(" AND C.CNAME='{0}'", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            dtx = bc.getdt(sqb.ToString());
            if (dtx.Rows.Count > 0)
            {
                //MessageBox.Show(dtx.Rows[0]["客户比例"].ToString());
                if (!string.IsNullOrEmpty(dtx.Rows[0]["客户比例"].ToString()))
                {
                    d10 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[0]["客户比例"].ToString(), '%')) / 100;
                  
                }

            }
            DataRow dr2 = dt.NewRow();
            dr2["项目"] = "含税";
            dr2["元套"] = ((d1 + d2 + d3 + d4 + d5) * (1 + d10)).ToString();
            YUAN_SET_HAVE_TAX = ((d1 + d2 + d3 + d4 + d5) * (1 + d10));
            //sqb = new StringBuilder();
            //sqb.AppendFormat("d1='{0}',d2='{1}',d3='{2}',d4='{3}',d5='{4}'", d1,d2,d3,d4,d5);

            dr2["批量小计"] = (((d1 + d2 + d3 + d4 + d5) * COUNT) * (1 + d10)).ToString();
            BATCH_TOTAL_HAVE_TAX = (((d1 + d2 + d3 + d4 + d5) * COUNT) * (1 + d10));
            dt.Rows.Add(dr2);
            DataRow dr3 = dt.NewRow();
            dr3["项目"] = "无外购采购比";
            decimal dx3 = 0;
            dx3 = (de1 + de2 + de3 + de4 + de7 + de8 + de9 + de10 + de11);
            decimal dx1 =dx3 * COUNT;
            decimal dx2 = (de13_a + de14_a + de15_a);
            if (d1 + d2 + d3 != 0)
            {
                dr3["元套"] = (dx1 / dx2) * 100 + "%";
                YUAN_SET_PURCHASE_PERCENT = (dx1 / dx2) * 100;
            }
            dr3["批量小计"] = "外购采购比";
            if (YUAN_SET_NO_TAX != 0)
            {
                dr3["主件用量"] = ((dx3 + de16) / YUAN_SET_NO_TAX) * 100 + "%";
                MAIN_DOSAGE_PURCHASE_PERCENT = ((dx3 + de16) / YUAN_SET_NO_TAX) * 100;
            }
            dt.Rows.Add(dr3);
            return dt;
        }
        #endregion
        #region RETURN_DIE_CUTTING_PRICE_DT
        public DataTable RETURN_DIE_CUTTING_PRICE_DT(DataTable dtt,DataGridView dgv2)//dtt为DT_EXCEL_TOTAL_DT
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("项目", typeof(string));
            dt.Columns.Add("刀模长米", typeof(string));
            dt.Columns.Add("元米", typeof(string));
            dt.Columns.Add("圆孔个数", typeof(string));
            dt.Columns.Add("元个", typeof(string));
            dt.Columns.Add("小计", typeof(decimal));
            if (dgv2.Rows.Count > 0)
            {
                for(i=0;i<dgv2.Rows.Count ;i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["项目"] = dgv2["项目", i].FormattedValue.ToString();
                    dr["刀模长米"] = dgv2["刀模长米", i].FormattedValue.ToString();
                    dr["圆孔个数"] = dgv2["圆孔个数", i].FormattedValue.ToString();
                    if (i == 0)
                    {
                        dtx = bc.getdt(cdie_cutting_cost.sql + string.Format(@" WHERE A.DIE_CUTTING='{0}' 
", dr["项目"].ToString()));
                        if (dr["项目"].ToString() != "" && dtx.Rows.Count > 0)
                        {

                            dr["元米"] = dtx.Rows[0]["未税单价"].ToString();//第一行调用预先在属性管理设定的单价 16/01/13

                        }
                        dtx = bc.getdt(cdie_cutting_cost.sql + string.Format(@" WHERE A.DIE_CUTTING='{0}' 
", "圆孔"));
                        if (dr["项目"].ToString() != "" &&
                            dr["圆孔个数"].ToString() != "" && dtx.Rows.Count > 0)
                        {
                            dr["元个"] = dtx.Rows[0]["未税单价"].ToString();//第一行调用预先在属性管理设定的单价 16/01/13
                        }

                    }
                    else
                    {
                        dr["元米"] = dgv2["元米", i].FormattedValue.ToString();//第二行为手输入单价 16/01/13
                        dr["元个"] = dgv2["元个", i].FormattedValue.ToString();///第二行为手输入单价 16/01/13

                    }

                    dt.Rows.Add(dr);
                }
                d1 = 0;
                d2 = 0;
                d3 = 0;
                d4 = 0;
                d5 = 0;
                d6 = 0;
                if (dt.Rows[0]["刀模长米"].ToString() == "" && dt.Rows[0]["圆孔个数"].ToString() == "")
                {
                }
                else
                {
                    DataTable dtx1 = bc.getdt(cdie_cutting_cost.sql + " WHERE A.DIE_CUTTING='" + dt.Rows[0]["项目"].ToString() + "'");
                    if (dtx1.Rows.Count > 0)
                    {

                        if (!string.IsNullOrEmpty(dtx1.Rows[0]["未税起机费"].ToString()))
                        {
                            d1 = decimal.Parse(dtx1.Rows[0]["未税起机费"].ToString());
                        }

                    }
                    if (!string.IsNullOrEmpty(dt.Rows[0]["刀模长米"].ToString()))
                    {
                        d2 = decimal.Parse(dt.Rows[0]["刀模长米"].ToString());
                    }

                    if (!string.IsNullOrEmpty(dt.Rows[0]["元米"].ToString()))
                    {
                        d3 = decimal.Parse(dt.Rows[0]["元米"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[0]["圆孔个数"].ToString()))
                    {
                        d4 = decimal.Parse(dt.Rows[0]["圆孔个数"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[0]["元个"].ToString()))
                    {
                        d5 = decimal.Parse(dt.Rows[0]["元个"].ToString());
                    }
                    d6 = Math.Max(d1, d2 * d3) + d4 * d5;
                    if (Math.Max(d1, d2 * d3) + d4 * d5 != 0)
                    {
                        dt.Rows[0]["小计"] = d6;
                    }
                }
                d1 = 0;
                d2 = 0;
                d3 = 0;
                d4 = 0;
                d5 = 0;
                d7 = 0;
                d8 = 0;
              //1
                if (dt.Rows[1]["刀模长米"].ToString() == "" && dt.Rows[1]["圆孔个数"].ToString() == "" &&
                    dt.Rows[1]["元米"].ToString() == "" && dt.Rows[1]["元个"].ToString() == "")
                {
                }
                else
                {
                    if (!string.IsNullOrEmpty(dt.Rows[1]["刀模长米"].ToString()))
                    {
                        d2 = decimal.Parse(dt.Rows[1]["刀模长米"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[1]["元米"].ToString()))
                    {
                        d3 = decimal.Parse(dt.Rows[1]["元米"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[1]["圆孔个数"].ToString()))
                    {
                        d4 = decimal.Parse(dt.Rows[1]["圆孔个数"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[1]["元个"].ToString()))
                    {
                        d5 = decimal.Parse(dt.Rows[1]["元个"].ToString());
                    }
                    d7=d2*d3+d4*d5;
                    dt.Rows[1]["小计"] = d7;
                }
                //t2MessageBox.Show("k");

                if (dt.Rows[2]["元米"].ToString() == "按平方")
                {
                    //MessageBox.Show(cprint_cutting_price.sql + " WHERE C.OFFER_ID='" + OFFER_ID + "'");
                    dt.Rows[2]["圆孔个数"] = dtt.Compute("SUM(刀模小计)", "");
                    string v1 = dtt.Compute("SUM(刀模小计)", "").ToString();
                    if (!string.IsNullOrEmpty(v1))
                    {
                        d8 = decimal.Parse(v1);
                    }

                }
                else
                {
                    dt.Rows[2]["圆孔个数"] = dt.Compute("SUM(小计)", "");
                    d8 = d6 + d7;
                }
                if (d8 != 0 && dtt.Rows[0]["数量"].ToString() != "")
                {
                    dt.Rows[2]["小计"] = (d8 / decimal.Parse(dtt.Rows[0]["数量"].ToString())).ToString("0.00");
                }
            }
            return dt;
        }
        #endregion
        #region RETURN_PORTRAY_DT
        public DataTable RETURN_PORTRAY_DT(DataTable dtt, DataGridView dgv)//dtt为DT_EXCEL_TOTAL_DT
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("写真类型", typeof(string));
            dt.Columns.Add("长", typeof(string));
            dt.Columns.Add("宽", typeof(string));
            dt.Columns.Add("总数量", typeof(string));
            dt.Columns.Add("单价", typeof(decimal));
            dt.Columns.Add("小计", typeof(decimal));
            if (dgv.Rows.Count > 0)
            {
                d8 = 0;
                for (i = 0; i < dgv.Rows.Count; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["写真类型"] = dgv["写真类型", i].FormattedValue.ToString();
                    dr["长"] = dgv["长", i].FormattedValue.ToString();
                    dr["宽"] = dgv["宽", i].FormattedValue.ToString();
                    dr["总数量"] = dgv["总数量", i].FormattedValue.ToString();
                    d1 = 0;
                    d2 = 0;
                    d3 = 0;
                    d4 = 0;
                    d5 = 0;
                    d6 = 0;
                    d7 = 0;
                    d9 = 0;
                    d10 = 0;
                    decimal d11 = 0;
                    decimal d12 = 0;
                    DataTable dtx1 = bc.getdt("SELECT * FROM PORTRAY A WHERE A.PORTRAY_TYPE='" + dr["写真类型"].ToString() +
                     "' AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='" + bc.RETURN_CUSTOMER_TYPE(dtt.Rows[0]["项目号"].ToString()) + "'");
                    if (dtx1.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(dtx1.Rows[0]["PUT_THE_NUMBER"].ToString()))//放数
                        {
                            d1 = decimal.Parse(dtx1.Rows[0]["PUT_THE_NUMBER"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx1.Rows[0]["ATTRITION_RATE"].ToString()))
                        {
                            d7 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx1.Rows[0]["ATTRITION_RATE"].ToString(), '%'));//损耗率
                        }
                        if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_MACHINE_COST"].ToString()))//含税起机费
                        {
                            d9 = decimal.Parse(dtx1.Rows[0]["TAX_MACHINE_COST"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))//税率
                        {
                            d10 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))//含税单价
                        {
                            d11 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                        }
                        d12 = d11 / (1 + d10 / 100);//未税单价 16/01/13
                    }
                    if (i == 0 || i == 1 || i == 2 || i == 3 || i == 4 || i==5 || i==6)
                    {
                        if (d12 != 0)
                        {
                            dr["单价"] = d12;//前7行的单价取自属性管理里预先设好的单价 16/01/13
                        }
                        else
                        {
                            dr["单价"] = DBNull.Value;
                        }
                    }
                    else if (!string.IsNullOrEmpty(dgv["单价", i].FormattedValue.ToString()))
                    {
                        dr["单价"] = dgv["单价", i].FormattedValue.ToString();
                    }
                    else
                    {
                        dr["单价"] = DBNull.Value;
                    }

                    if (!string.IsNullOrEmpty(dr["长"].ToString()))
                    {
                        d2 = decimal.Parse(dr["长"].ToString());
                    }

                    if (!string.IsNullOrEmpty(dr["宽"].ToString()))
                    {
                        d3 = decimal.Parse(dr["宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr["总数量"].ToString()))
                    {
                        d4 = decimal.Parse(dr["总数量"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr["单价"].ToString()))
                    {
                        d5 = decimal.Parse(dr["单价"].ToString());
                    }
                    /*sqb = new StringBuilder();
                    sqb.AppendFormat("LENGTH:{0},", d2);
                    sqb.AppendFormat("WIDTH:{0},", d3);
                    sqb.AppendFormat("TOTAL_COUNT:{0},", d4);
                    sqb.AppendFormat("PRICE:{0},", d5);
                    sqb.AppendFormat("未税起机费：{0}", d9 / (1 + d10 / 100));
                    sqb.AppendFormat("另一比效值：{0}", (d2 + 30) / 1000 * (d3 + 30) / 1000 * d5 * (d4 + d1 + d4 * d7 / 100));
                    MessageBox.Show(sqb.ToString());*/
                    d6 = Math.Max(d9 / (1 + d10 / 100), (d2 + 30) / 1000 * (d3 + 30) / 1000 * d5 * (d4 + d1 + d4 * d7 / 100));//16/01/08 修正公式加入最大值的比较
                   
                    if (d2 == 0 || d3 == 0 || d4 == 0 || d5 == 0)
                    {
                    }
                    else
                    {
                        dr["小计"] = d6;
                    }
                    d8 = d8 + d6;
                    dt.Rows.Add(dr);
                }
                d1 = 0;
                sqb = new StringBuilder();
                sqb.AppendFormat("SELECT A.TAX_UNIT_PRICE/(1+A.TAX_RATE/100) AS 未税单价  FROM PORTRAY A ");
                sqb.AppendFormat(" WHERE A.PORTRAY_TYPE='{0}' ", "批次写真运费");
                sqb.AppendFormat(" AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='{0}'", bc.RETURN_CUSTOMER_TYPE(dtt.Rows[0]["项目号"].ToString()));
                dtx = bc.getdt(sqb.ToString());
                if (dtx.Rows.Count > 0)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[0][0].ToString()))
                    {
                        d1 = decimal.Parse(dtx.Rows[0][0].ToString());
                    }
                }
                DataRow dr2 = dt.NewRow();
                dr2["总数量"] = "汇总";
                if (d8 > 0)//证明存在写真项目
                {
                    TOTAL_COST_PORTRAY = (d8+d1) / decimal.Parse(dtt.Rows[0]["数量"].ToString());//在有写真时将运费加入到写真TOTAL中 16/01/05
                    dr2["单价"] = TOTAL_COST_PORTRAY;

                }
                else
                {
                    dr2["单价"] = 0;
                    TOTAL_COST_PORTRAY = 0;
                }
                dt.Rows.Add(dr2);
            }
  
            return dt;
        }
        #endregion
        #region RETURN_PARTS_AUXILIAR_DT
        public DataTable RETURN_PARTS_AUXILIAR_DT(string PFID, DataTable dtt)//dtt为DT_EXCEL_TOTAL_DT
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("配件名", typeof(string));
            dt.Columns.Add("用量", typeof(string));
            dt.Columns.Add("单价", typeof(decimal));
            dt.Columns.Add("单位", typeof(string));
            dt.Columns.Add("小计", typeof(decimal));
            dt.Columns.Add("备注", typeof(string));
    
            dtx = bc.getdt(cprint_parts_auxiliary.sql  + " WHERE C.PFID='" + PFID  + "'");
            d8 = 0;
            if (dtx.Rows.Count > 0)
            {
                for (i = 0; i < dtx.Rows.Count; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["配件名"] = dtx.Rows[i]["配件名"].ToString();
                    dr["用量"] = dtx.Rows[i]["用量"].ToString();
                    dr["备注"] = dtx.Rows[i]["备注"].ToString();
                    dr["单位"] = dtx.Rows[i]["单位"].ToString();
                    d1 = 0;
                    d2 = 0;
                    d3 = 0;
                    d4 = 0;
                    d5 = 0;
                    d6 = 0;
                    d7 = 0;
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                    {
                        dr["单价"] = dtx.Rows[i]["单价"].ToString();//此单价为已经存在数据库的单价 16/01/10
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                    {
                        //此单价为已经存在数据库的单价 16/01/10
                    }
                    else
                    {
                        DataTable dtx1 = bc.getdt("SELECT * FROM PARTS_AUXILIARY A WHERE A.PARTS_AUXILIARY='" + dtx.Rows[i]["配件名"].ToString() + "'");
                        if (dtx1.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))
                            {
                                d1 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                            }
                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                            {

                                d7 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                            }

                            dr["单价"] = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()) / (1 + decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString()) / 100);
                            //此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10
                        }
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["用量"].ToString()))
                    {
                        d2 = decimal.Parse(dtx.Rows[i]["用量"].ToString());
                    }

                    if (!string.IsNullOrEmpty(dr["单价"].ToString()))
                    {
                        d3 = decimal.Parse(dr["单价"].ToString());
                    }
                    if (d2 == 0 || d3 == 0)
                    {
                    }
                    else
                    {
                        d6 = d2 * d3;
                    }
                    if (d6 != 0)
                    {
                        dr["小计"] = d6;
                    }
                    d8 = d8 + d6;
                    dt.Rows.Add(dr);
                 
                    
                }
                DataRow dr2 = dt.NewRow();
                dr2["单位"] = "汇总";
                dr2["小计"] = d8;
                dt.Rows.Add(dr2);
                TOTAL_COST_PARTS_AUXILIARY = d8;
           

            }
            return dt;
        }
        #endregion
        #region RETURN_PACK_MATERIAL_DT
        public DataTable RETURN_PACK_MATERIAL_DT(string PFID, DataTable dtt)//dtt为DT_EXCEL_TOTAL_DT
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("项目", typeof(string));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("长", typeof(string));
            dt.Columns.Add("宽", typeof(string));
            dt.Columns.Add("高", typeof(string));
            dt.Columns.Add("箱形", typeof(string));
            dt.Columns.Add("材质", typeof(string));
            dt.Columns.Add("单价", typeof(decimal));
            dt.Columns.Add("小计", typeof(decimal));
            PACK_LENGTH = 0;
            PACK_WIDTH = 0;
            PACK_HEIGHT = 0;
            dtx = bc.getdt(cprint_pack_material .sql  + " WHERE C.PFID='" + PFID + "'");
            d8 = 0;
            //MessageBox.Show("34");
            if (dtx.Rows.Count > 0)
            {
                
                for (i = 0; i < dtx.Rows.Count; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["项目"] = dtx.Rows[i]["项目"].ToString();
                    dr["数量"] = dtx.Rows[i]["数量"].ToString();
                    dr["箱形"] = dtx.Rows[i]["箱形"].ToString();
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["长"].ToString()))
                    {
                        PACK_LENGTH = decimal.Parse(dtx.Rows[i]["长"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["宽"].ToString()))
                    {
                        PACK_WIDTH = decimal.Parse(dtx.Rows[i]["宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["高"].ToString()))
                    {
                        PACK_HEIGHT = decimal.Parse(dtx.Rows[i]["高"].ToString());
                    }
                    dr["长"] = dtx.Rows[i]["长"].ToString();
                    dr["宽"] = dtx.Rows[i]["宽"].ToString();
                    dr["高"] = dtx.Rows[i]["高"].ToString();
                    dr["箱形"] = dtx.Rows[i]["箱形"].ToString();
                    dr["材质"] = dtx.Rows[i]["材质"].ToString();
                    d1 = 0;
                    d2 = 0;
                    d3 = 0;
                    d4 = 0;
                    d5 = 0;
                    d6 = 0;
                    d7 = 0;
                    if (i == 0 || i == 1 || i == 2 || i == 3)
                    {
                        if ((i == 0 || i == 1 || i == 2) && ( dr["数量"].ToString() == "" || dr["长"].ToString() == ""
                            || dr["宽"].ToString() == "" || dr["高"].ToString() == "" || dr["箱形"].ToString() == "" || dr["材质"].ToString() == ""))
                        {

                        }
                        else if ((i ==3) && (dr["数量"].ToString() == ""  || dr["长"].ToString() == ""
                            || dr["宽"].ToString() == "" || dr["材质"].ToString() == ""))
                        {

                        }
                        else
                        {
                         
                            DataTable dtx1 = bc.getdt(@"
SELECT * FROM PACK_MATERIAL A WHERE A.PACK_MATERIAL='" + dtx.Rows[i]["材质"].ToString() + "'AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='"+
                                                       bc.RETURN_CUSTOMER_TYPE (dtt.Rows [0]["项目号"].ToString ())+"' ");
                            if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                            {
                                dr["单价"] = dtx.Rows[i]["单价"].ToString();//此单价为已经存在数据库的单价 16/01/10
                            }
                            else
                            {
                                if (dtx1.Rows.Count > 0)
                                {
                                    if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))
                                    {
                                        d1 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                                    }
                                    if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                                    {

                                        d7 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                                    }

                                    dr["单价"] = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()) /
                                        (1 + decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString()) / 100);//此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10

                                }
                            }
                           
                                if (!string.IsNullOrEmpty (dr["单价"].ToString()))
                                {

                                    d1 = decimal.Parse(dr["长"].ToString());
                                    d2 = decimal.Parse(dr["宽"].ToString());
                                    if (i != 3)
                                    {
                                        d3 = decimal.Parse(dr["高"].ToString());
                                    }
                                    d4 = decimal.Parse(dr["数量"].ToString());
                                    d5 = decimal.Parse(dr["单价"].ToString());
                                    if (i == 0 || i == 1 || i == 2)
                                    {
                                        if (dr["箱形"].ToString() == "天地盖")
                                        {

                                            dr["小计"] = (d1 + d3 * 2 + 80) / 1000 * (d2 + d3 * 2 + 60) / 1000 * 2 * d5 * d4;
                                        }
                                        else
                                        {
                                            dr["小计"] = (d1 + d2 + 80) / 1000 * (d2 + d3 + 60) / 1000 * 2 * d5 * d4;

                                        }
                                        
                                    }
                                    else if (i == 3)
                                    {
                                        dr["小计"] = d5 * d4 * d1 / 1000 * d2 / 1000;
                                    }
                                    
                                }
                        }
                    }
                   
                    if (i == 4 || i == 5 || i == 6)
                    {
                        //MessageBox.Show("35");
                        d1 = 0;
                        d2 = 0;
                        d7 = 0;
                        //MessageBox.Show(dtt.Rows[0]["项目号"].ToString());
                        DataTable dtx1 = bc.getdt(@"
SELECT * FROM PACK_MATERIAL A WHERE A.PACK_MATERIAL='" + dtx.Rows[i]["箱形"].ToString() + "'AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='" +
                       bc.RETURN_CUSTOMER_TYPE(dtt.Rows[0]["项目号"].ToString()) + "' ");
                        //MessageBox.Show("40");
                        if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                        {
                            dr["单价"] = dtx.Rows[i]["单价"].ToString();//此单价为已经存在数据库的单价 16/01/10
                        }
                        else
                        {
                            if (dtx1.Rows.Count > 0)
                            {
                                //MessageBox.Show("38");
                                if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))
                                {
                                    d1 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                                }
                                if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                                {

                                    d7 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                                }

                                dr["单价"] = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()) /
                                    (1 + decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString()) / 100);
                                //此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10
                            }
                        }
                            //MessageBox.Show("37");
                            if (!string.IsNullOrEmpty (dr["单价"].ToString()) && dr["数量"].ToString ()!="")
                            {
                                dr["小计"] = decimal.Parse(dr["单价"].ToString()) * decimal.Parse(dr["数量"].ToString());
                            }
                            //MessageBox.Show("36");
                     }
                    if (i == 7 || i == 8 || i == 9)
                    {
                        if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                        {
                            dr["单价"] = dtx.Rows[i]["单价"].ToString();
                        }
                        else
                        {
                            dr["单价"] = DBNull.Value;
                        }
                      
                        if (dr["单价"].ToString() != "" && dr["数量"].ToString() != "")
                        {
                            dr["小计"] = decimal.Parse(dr["单价"].ToString()) * decimal.Parse(dr["数量"].ToString());
                        }
                    }
                    if (dr["小计"].ToString() != "")
                    {
                      d2=decimal.Parse(dr["小计"].ToString());
                    }
                      d8 = d8 +d2;
                    dt.Rows.Add(dr);

                    }
                if (dtt.Rows[0]["数量"].ToString() != "")
                {
                    DataRow dr2 = dt.NewRow();
                    dr2["材质"] = "汇总";
                    dr2["小计"] = d8.ToString("0.00");
                    dt.Rows.Add(dr2);
                    TOTAL_COST_PACK_MATERIAL = d8;
                }
            
                }
            
             return dt;
        }
        #endregion
        #region RETURN_TRANSPORT_DT
        public DataTable RETURN_TRANSPORT_DT(DataTable dtt,DataGridView dgv)//dtt为DT_EXCEL_TOTAL_DT
        {
           
            DataTable dt = new DataTable();
            dt.Columns.Add("长", typeof(string));
            dt.Columns.Add("宽", typeof(string));
            dt.Columns.Add("高", typeof(string));
            dt.Columns.Add("总箱数", typeof(string));
            dt.Columns.Add("总立方数", typeof(decimal));
            dt.Columns.Add("运输方式", typeof(string));
            dt.Columns.Add("单价", typeof(decimal));
            dt.Columns.Add("小计", typeof(decimal));
            d9 = 0;
            d10 = 0;
            decimal d11 = 0;
            if (dgv.Rows.Count > 0)
            {
                for (i = 0; i < dgv.Rows.Count; i++)
                {
                  
                    PACK_LENGTH = 0;
                    PACK_WIDTH = 0;
                    PACK_HEIGHT = 0;
                    DataRow dr = dt.NewRow();
                    dr["长"] = dgv["长", i].FormattedValue.ToString();
                    dr["宽"] = dgv["宽", i].FormattedValue.ToString();
                    dr["高"] = dgv["高", i].FormattedValue.ToString();
          
                    dr["总箱数"] = dgv["总箱数", i].FormattedValue.ToString();
                    dr["运输方式"] = dgv["运输方式", i].FormattedValue.ToString();
                    d1 = 0;
                    d2 = 0;
                    d3 = 0;
                    d4 = 0;
                    d5 = 0;
                    d6 = 0;
                    d7 = 0;
                    d8 = 0;
                    if (!string.IsNullOrEmpty(dr["长"].ToString()))
                    {
                        PACK_LENGTH = decimal.Parse(dr["长"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr["宽"].ToString()))
                    {
                        PACK_WIDTH = decimal.Parse(dr["宽"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr["高"].ToString()))
                    {
                        PACK_HEIGHT = decimal.Parse(dr["高"].ToString());
                    }
                    if (dr["高"].ToString() == "" || dr["总箱数"].ToString() == "" || dr["长"].ToString() == ""
                            || dr["宽"].ToString() == "")
                    {

                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(dr["总箱数"].ToString()))
                        {
                            d1 = decimal.Parse(dr["总箱数"].ToString());
                        }
                 
                        d11 = (PACK_LENGTH + 10) / 1000 * (PACK_WIDTH + 10) / 1000 * (PACK_HEIGHT + 10) / 1000 * d1;
                        if (d11 != 0)
                        {
                            dr["总立方数"] = d11.ToString("0.00");
                        }
                    
                        d2=(PACK_LENGTH + 10) / 1000 * (PACK_WIDTH + 10) / 1000 * (PACK_HEIGHT + 10) / 1000 * d1;
                    }
                    if (i == 2)
                    {
                        if (!string.IsNullOrEmpty(dgv["总立方数", i].FormattedValue.ToString()))
                        {
                            d11 = decimal.Parse(dgv["总立方数", i].FormattedValue.ToString());
                            dr["总立方数"] = d11.ToString("0.00");
                        }
                        else
                        {
                            dr["总立方数"] = DBNull.Value;
                        }
                    }
                    /*sqb = new StringBuilder();
                    sqb.AppendFormat("{0}, ", dr["运输方式"].ToString());
                    sqb.AppendFormat("{0}, ", dr["总立方数"].ToString());
                    sqb.AppendFormat("第 {0} 行, ", i+1);
                    MessageBox.Show(sqb.ToString ());*/
                    if (dr["运输方式"].ToString() != "" && !string.IsNullOrEmpty(dr["总立方数"].ToString()) && decimal.Parse(dr["总立方数"].ToString()) != 0 && 
                       ( i==0 || i==1))
                    {
                      
                            DataTable dtx1 = bc.getdt(@"
SELECT * FROM TRANSPORT A WHERE A.TRANSPORT='" + dr["运输方式"].ToString() + "'AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='" +
                                                       bc.RETURN_CUSTOMER_TYPE(dtt.Rows[0]["项目号"].ToString()) + "' ");
                            if (dtx1.Rows.Count > 0)
                            {
                                if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                                {
                                    d4 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                                }
                                if (d2 < 50)
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
                                dr["单价"] = (d3 / (1 + d4 / 100)).ToString("0.00");//单价保留两位小数即可 16/01/13
                              
                                if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_TRANSPORT_COST"].ToString()))
                                {
                                    d5 = decimal.Parse(dtx1.Rows[0]["TAX_TRANSPORT_COST"].ToString());
                                }
                                d6 = d5 / (1 + d4 / 100);
                            }
                         
                        }
                    if (i == 2)
                    {
                        if (!string.IsNullOrEmpty(dgv["单价", i].FormattedValue.ToString()))
                        {
                            dr["单价"] = dgv["单价", i].FormattedValue.ToString();
                        }
                        else
                        {
                            dr["单价"] = DBNull.Value;
                        }
                       
                        if (!string.IsNullOrEmpty(dr["单价"].ToString()) && dr["总立方数"].ToString() != "")
                        {
                            d7 = decimal.Parse(dr["单价"].ToString()) * decimal.Parse(dr["总立方数"].ToString());
                            dr["小计"] = d7.ToString("0");

                        }
                    }
                    else  if (dr["单价"].ToString() != "" && dr["总立方数"].ToString()!="")
                    {
                        d7 = Math.Max(d2 * d3 / (1 + d4 / 100), d6);
                        dr["小计"] = d7.ToString("0");
                    }
                    d9 = d9 + d7;
                    dt.Rows.Add(dr);
                }
           
                if (dtt.Rows[0]["数量"].ToString() != "")
                {
                 
                    DataRow dr2 = dt.NewRow();
                    dr2["运输方式"] = "产品运输成本总价为";
                    dr2["单价"] = d9.ToString("0.00");
                    dt.Rows.Add(dr2);
                    DataRow dr3 = dt.NewRow();
                    dr3["长"] = "运输";
                    if (d9 != 0)
                    {
                        d10 = d9 / decimal.Parse(dtt.Rows[0]["数量"].ToString());
                        dr3["单价"] = d10.ToString("0.00");
                    }
                    dt.Rows.Add(dr3);
                   
                }

            }
          
            return dt;
        }
        #endregion
        #region RETURN_ARTIFICIAL_DT
        public DataTable RETURN_ARTIFICIAL_DT(string PFID, DataTable dtt)//dtt为DT_EXCEL_TOTAL_DT
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("项目", typeof(string));
            dt.Columns.Add("单价", typeof(decimal));
            dt.Columns.Add("数量", typeof(string));
            dt.Columns.Add("小计", typeof(decimal));
            dt.Columns.Add("元套", typeof(decimal));
       
            dtx = bc.getdt(cprint_artificial.sql + " WHERE C.PFID='" + PFID + "'");
            d9 = 0;
            d10 = 0;
            if (dtx.Rows.Count > 0)
            {
                for (i = 0; i < dtx.Rows.Count; i++)
                {
                   
                    DataRow dr = dt.NewRow();
                    dr["项目"] = dtx.Rows[i]["项目"].ToString();
                    dr["数量"] = dtx.Rows[i]["数量"].ToString();
                    d1 = 0;
                    d2 = 0;
                    d3 = 0;
                    d4 = 0;
                    d5 = 0;
                    d6 = 0;
                    d7 = 0;
                    d8 = 0;

                    if (!string.IsNullOrEmpty(dtx.Rows[i]["数量"].ToString()))
                    {
                        d3 = decimal.Parse(dtx.Rows[i]["数量"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                    {
                        dr["单价"] = decimal.Parse(dtx.Rows[i]["单价"].ToString());//此单价为已经存在数据库的单价 16/01/10
                        d6 = decimal.Parse(dr["单价"].ToString());
                    }
                    else
                    {
                        DataTable dtx1 = bc.getdt(@"
SELECT * FROM ARTIFICIAL A WHERE A.ARTIFICIAL='" + dtx.Rows[i]["项目"].ToString() + "'AND SUBSTRING(A.CUSTOMER_TYPE,1,1)='" +
                                 bc.RETURN_CUSTOMER_TYPE(dtt.Rows[0]["项目号"].ToString()) + "' ");
                        if (dtx1.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                            {
                                d4 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                            }
                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString()))
                            {
                                d5 = decimal.Parse(dtx1.Rows[0]["TAX_UNIT_PRICE"].ToString());
                            }
                            d6 = d5 / (1 + d4 / 100);
                            dr["单价"] = d6;//此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10
                        }
                    }
                    if (!string.IsNullOrEmpty (dr["单价"].ToString()) && dr["数量"].ToString() != "")
                    {
                        d7 = d3 * d6;
                        dr["小计"] = d7;
                    }
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                    {
                        dr["单价"] = decimal.Parse(dtx.Rows[i]["单价"].ToString());//此单价为已经存在数据库的单价 16/01/10
                    }
                    else
                    {
                        if (i == 1)
                        {
                            if (!string.IsNullOrEmpty(dtx.Rows[i]["单价"].ToString()))
                            {
                                dr["单价"] = dtx.Rows[i]["单价"].ToString();//此单价为新增作业时由属性管理相关作业调入的参数产生单价 16/01/10
                            }
                            else
                            {
                                dr["单价"] = DBNull.Value;
                            }

                        }
                    }
                    if (i == 1 && !string.IsNullOrEmpty (dr["单价"].ToString()) && !string.IsNullOrEmpty (dr["数量"].ToString()))
                    {
                        d7 = decimal.Parse(dr["单价"].ToString()) * decimal.Parse(dr["数量"].ToString());
                        dr["小计"] = d7;
                    }
                    d9 = d9 + d7;
                    dt.Rows.Add(dr);

                }
                if (dtt.Rows[0]["数量"].ToString() != "")
                {
                    dt.Rows[0]["元套"] = d9;
                    TOTAL_COST_ARTIFICIAL = d9;
                }

            }

            return dt;
        }
        #endregion
        #region RETURN_PURCHASE_DT
        public DataTable RETURN_PURCHASE_DT(DataTable dtt,DataGridView dgv)//dtt为DT_EXCEL_TOTAL_DT
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("类型一", typeof(string));
            dt.Columns.Add("外购价一", typeof(string));
            dt.Columns.Add("管理费一", typeof(decimal));
            dt.Columns.Add("小计一", typeof(decimal));
            dt.Columns.Add("类型二", typeof(string));
            dt.Columns.Add("外购价二", typeof(string));
            dt.Columns.Add("管理费二", typeof(decimal));
            dt.Columns.Add("小计二", typeof(decimal));
            d9 = 0;
            d10 = 0;
            decimal d12 = 0;
            if (dgv.Rows.Count > 0)
            {
                for (i = 0; i < dgv.Rows.Count; i++)
                {

                    DataRow dr = dt.NewRow();
                    DataTable dtx1 = new DataTable();
                    dr["类型一"] = dgv["类型一", i].FormattedValue.ToString();
                    dr["外购价一"] = dgv["外购价一", i].FormattedValue.ToString();
                    dr["类型二"] = dgv["类型二", i].FormattedValue.ToString();
                    dr["外购价二"] = dgv["外购价二", i].FormattedValue.ToString();
                    d1 = 0;
                    d2 = 0;
                    d3 = 0;
                    d4 = 0;
                    d5 = 0;
                    d6 = 0;
                    d7 = 0;
                    d8 = 0;
                    decimal d11 = 0;
           
                        if (!string.IsNullOrEmpty(dr["外购价一"].ToString()))
                        {
                            d3 = decimal.Parse(dr["外购价一"].ToString());
                        }

                        dtx1 = bc.getdt(@"
SELECT * FROM PURCHASE A WHERE A.PURCHASE='" + dr["类型一"].ToString() + "'");
                        if (dtx1.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                            {
                                d5 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                            }

                            d7 = d3 / (1 - d5 / 100) * d5 / 100;
                            dr["管理费一"] = d7;
                            d8 = d7 + d3;//公式外购价+管理费 16/01/13
                            dr["小计一"] = d8;
                        }

                        if (!string.IsNullOrEmpty(dr["外购价二"].ToString()))
                        {
                            d4 = decimal.Parse(dr["外购价二"].ToString());
                        }
                        dtx1 = bc.getdt(@"
SELECT * FROM PURCHASE A WHERE A.PURCHASE='" + dr["类型二"].ToString() + "'");
                        if (dtx1.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtx1.Rows[0]["TAX_RATE"].ToString()))
                            {
                                d1 = decimal.Parse(dtx1.Rows[0]["TAX_RATE"].ToString());
                            }

                            d11 = d4 / (1 - d1 / 100) * d1 / 100;
                            dr["管理费二"] = d11;
                            d8 = d11 + d4;
                            dr["小计二"] = d8;
                        }
                        d12 = d12 + d7 + d3 + d11 + d4;
                        d9 = d9 + d3 + d4;
                        dt.Rows.Add(dr);

                    
                }
                DataRow dr2 = dt.NewRow();
                dr2["类型二"] = "外购件合计";
                dr2["管理费二"] = d9;
                dt.Rows.Add(dr2);
                TOTAL_COST_PURCHASE = d9;
                TOTAL_COST_PURCHASE_TWO = d12;
            }
   
            return dt;
        }
        #endregion
        #region GetTableInfo_PAPER
        public DataTable GetTableInfo_PAPER()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("面纸", typeof(string));
            dt.Columns.Add("面纸克重", typeof(string));
            dt.Columns.Add("面纸单价", typeof(string));
            dt.Columns.Add("面纸单个用量", typeof(decimal));
            dt.Columns.Add("面纸小计", typeof(decimal));
            return dt;
        }

        #endregion
        #region RETURN_PAPER_TOTAL
        public DataTable RETURN_PAPER_TOTAL(DataTable dt)
        {
         
            DataTable dtx2 = GetTableInfo_PAPER();
            DataTable dtx5 = GetTableInfo_PAPER();
            DataTable dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dt, "面纸", "面纸克重");
            d1 = 0;
            d2 = 0;
            d3 = 0;
            i = 1;
            if (dtx3.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dtx3.Rows)
                {
                    DataRow dr = dtx2.NewRow();
                    dr["序号"] = i;
                    dr["面纸"] = dr1["VALUE1"].ToString();
                    dr["面纸克重"] = dr1["VALUE2"].ToString();
                    DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt, "", string.Format (" 面纸='{0}' AND 面纸克重='{1}'",dr1["VALUE1"].ToString(),
                        dr1["VALUE2"].ToString()));
                    if (dtx4.Rows.Count > 0)
                    {
                        dr["面纸单价"] = dtx4.Rows [0]["面纸单价"].ToString();
                        dr["面纸单个用量"] = dtx4.Compute("SUM(面纸单个用量)", "");
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(面纸小计)", "").ToString()))
                        {
                            d3 = decimal.Parse(dtx4.Compute("SUM(面纸小计)", "").ToString());
                            dr["面纸小计"] = d3.ToString();
                        }
                      
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(面纸小计)", "").ToString ()))
                        {
                            d1 = decimal.Parse(dtx4.Compute("SUM(面纸小计)", "").ToString ()) + d1;
                        }
                    }
                    dtx2.Rows.Add(dr);
                    i = i + 1;
                }
                dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dt, "底纸", "底纸克重");
                if (dtx3.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dtx3.Rows)
                    {
                        DataRow dr = dtx2.NewRow();
                        dr["序号"] = i;
                        dr["面纸"] = dr1["VALUE1"].ToString();
                        dr["面纸克重"] = dr1["VALUE2"].ToString();
                        DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt, "", string.Format(" 底纸='{0}' AND 底纸克重='{1}'", dr1["VALUE1"].ToString(),
                        dr1["VALUE2"].ToString()));
                        dr["面纸单价"] = dtx4.Rows[0]["底纸单价"].ToString();
                        dr["面纸单个用量"] = dtx4.Compute("SUM(底纸单个用量)", "");
                        //MessageBox.Show(dtx4.Compute("SUM(底纸单个用量)", "").ToString ());
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(底纸小计)", "").ToString()))
                        {
                            d3 = decimal.Parse(dtx4.Compute("SUM(底纸小计)", "").ToString());
                            dr["面纸小计"] = d3.ToString();
                        }
                     
                        if (!string.IsNullOrEmpty(dr["面纸小计"].ToString()))
                        {
                            d1 = decimal.Parse(dr["面纸小计"].ToString()) + d1;
                        }
                        dtx2.Rows.Add(dr);
                        i = i + 1;      
                    }
                }
           
                dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dtx2, "面纸", "面纸克重");
                i = 1;
                if (dtx3.Rows.Count > 0)
                {
                  
                    foreach (DataRow dr in dtx3.Rows)
                    {
                        DataRow dr1 = dtx5.NewRow();
                        dr1["序号"] = i;
                        dr1["面纸"] = dr["VALUE1"].ToString();
                        dr1["面纸克重"] = dr["VALUE2"].ToString();
                        DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dtx2, "", string.Format(" 面纸='{0}' AND 面纸克重='{1}'", dr["VALUE1"].ToString(),
                        dr["VALUE2"].ToString()));
                        if (dtx4.Rows.Count > 0)
                        {
                            dr1["面纸单价"] = dtx4.Rows[0]["面纸单价"].ToString();
                        }
                        dr1["面纸单个用量"] = dtx4.Compute("SUM(面纸单个用量)", "").ToString();
                        dr1["面纸小计"] = dtx4.Compute("SUM(面纸小计)", "").ToString();
                        i = i + 1;
                        dtx5.Rows.Add(dr1);
                    }
                }
            }
            d3 = 0;
            DataRow dr2 = dtx5.NewRow();
            dr2["面纸克重"] = "用纸汇总";
            dr2["面纸单价"] = "合计";
            dr2["面纸小计"] = d1;
            dtx5.Rows.Add(dr2);
            DataRow dr3 = dtx5.NewRow();
            dr3["面纸单价"] = "小计";
            if (!string.IsNullOrEmpty(dt.Rows[0]["数量"].ToString()))
            {
                d3 = d1 / decimal.Parse(dt.Rows[0]["数量"].ToString());
                dr3["面纸小计"] = d3.ToString();
              
            }
            dtx5.Rows.Add(dr3);
            return dtx5;

        }
        #endregion
        #region RETURN_PAPER_CORE_TOTAL
        public DataTable RETURN_PAPER_CORE_TOTAL(DataTable dt)
        {
            d1 = 0;
            d2 = 0;
            i = 1;
            DataTable dtx2 = new DataTable();
            dtx2.Columns.Add("序号", typeof(string));
            dtx2.Columns.Add("芯纸", typeof(string));
            dtx2.Columns.Add("芯纸规格", typeof(string));
            dtx2.Columns.Add("芯纸单价", typeof(string));
            dtx2.Columns.Add("芯纸单个用量", typeof(string));
            dtx2.Columns.Add("芯纸小计", typeof(string));
            DataTable dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dt, "芯纸", "芯纸规格");
            if (dtx3.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dtx3.Rows)
                {
                    DataRow dr = dtx2.NewRow();
                    dr["序号"] = i;
                    dr["芯纸"] = dr1["VALUE1"].ToString();
                    dr["芯纸规格"] = dr1["VALUE2"].ToString();
                    DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt, "", string.Format(" 芯纸='{0}' AND 芯纸规格='{1}'", dr1["VALUE1"].ToString(),
                        dr1["VALUE2"].ToString()));
                    if (dtx4.Rows.Count > 0)
                    {
                        dr["芯纸单价"] = dtx4.Rows[0]["芯纸单价"].ToString();
                        dr["芯纸单个用量"] = dtx4.Compute("SUM(芯纸单个用量)", "");
                        dr["芯纸小计"] = dtx4.Compute("SUM(芯纸小计)", "");
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(芯纸小计)", "").ToString()))
                        {
                            d1 = decimal.Parse(dtx4.Compute("SUM(芯纸小计)", "").ToString()) + d1;
                        }
                    }
                    dtx2.Rows.Add(dr);
                    i = i + 1;
                }
            dtx2 = bc.GET_DT_TO_DV_TO_DT(dtx2, "芯纸,芯纸规格 ASC", "");
            i = 1;
            if (dtx2.Rows.Count > 0)
            {
                foreach (DataRow dr in dtx2.Rows)
                {
                    dr["序号"] = i;
                    dr["芯纸"] = dr["芯纸"].ToString();
                    dr["芯纸规格"] = dr["芯纸规格"].ToString();
                    dr["芯纸单价"] = dr["芯纸单价"].ToString();
                    dr["芯纸单个用量"] = dr["芯纸单个用量"].ToString();
                    dr["芯纸小计"] = dr["芯纸小计"].ToString();
                    i = i + 1;
                }
            }
            }
            DataRow dr2 = dtx2.NewRow();
            dr2["芯纸规格"] ="芯纸汇总";
            dr2["芯纸单价"] = "合计";
            dr2["芯纸小计"] = d1;
            dtx2.Rows.Add(dr2);
            DataRow dr3 = dtx2.NewRow();
            dr3["芯纸单价"] = "小计";
            if (!string.IsNullOrEmpty(dt.Rows[0]["数量"].ToString()))
            {
                dr3["芯纸小计"] = d1 / decimal.Parse(dt.Rows[0]["数量"].ToString());
            }
            dtx2.Rows.Add(dr3);
            return dtx2;

        }
        #endregion
        #region RETURN_PRINTING_TOTAL
        public DataTable RETURN_PRINTING_TOTAL(DataTable dt)
        {
            d3 = 0;
            DataTable dtx2 = new DataTable();
            dtx2.Columns.Add("序号", typeof(string));
            dtx2.Columns.Add("机器型号", typeof(string));
            dtx2.Columns.Add("几款", typeof(string));
            dtx2.Columns.Add("小计", typeof(string));

            //MessageBox.Show(dt.Rows[0]["数量"].ToString());
            sqb = new StringBuilder();
            sqb.AppendFormat("机器型号 NOT IN ('主体','汇总')");
            sqb.AppendFormat(" AND 机器型号 IS NOT NULL AND 机器型号<>''");
            sqb.AppendFormat(" AND 印刷选项<>'不印刷'");//16/01/07 修改当印刷选项为不印刷时不统计机器型号
            DataTable dt1 = bc.GET_DT_TO_DV_TO_DT(dt, "",sqb.ToString ());
            DataTable dtx3 = bc.RETURN_NOHAVE_REPEAT_DT (dt1,"机器型号");
            if (dtx3.Rows.Count > 0)
            {
                i = 1;
                foreach (DataRow dr1 in dtx3.Rows)
                {
                    d1 = 0;
                    d2 = 0;
                    DataRow dr = dtx2.NewRow();
                    DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt1, "", "机器型号='" + dr1["VALUE"].ToString() + "' ");
                    if (dtx4.Rows.Count > 0)
                    {
                        dr["几款"] = dtx4.Rows.Count;
                        dr["机器型号"] = dr1["VALUE"].ToString();
                        dr["序号"] = i;
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(正反印工合计)", "").ToString()))
                        {
                            d1 = decimal.Parse(dtx4.Compute("SUM(正反印工合计)", "").ToString());
                        }
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(正反CTP合计)", "").ToString()))
                        {
                            d2 = decimal.Parse(dtx4.Compute("SUM(正反CTP合计)", "").ToString());
                        }
                        dr["小计"] = d1 + d2;
                    }
                    d3 = d3+d1+d2;
                    dtx2.Rows.Add(dr);
                    i = i + 1; 
                }
     
            }
          
            DataRow dr2 = dtx2.NewRow();
            dr2["机器型号"] = "印刷汇总";
            dr2["小计"] = d3;
            dtx2.Rows.Add(dr2);
            DataRow dr3 = dtx2.NewRow();
            if (!string.IsNullOrEmpty(dt.Rows[0]["数量"].ToString()))
            {
                dr3["小计"] = d3 / decimal.Parse(dt.Rows[0]["数量"].ToString());//此时的DT需为含所有内容的DT,含不印刷的情况，
                                                                               //如果只有一行且是不印刷，去除该行将0行无值 16/01/20
            }
            dtx2.Rows.Add(dr3);
       
            return dtx2;
        }
        #endregion
        #region RETURN_SURFACE_MACHINING
        public DataTable RETURN_SURFACE_MACHINING(DataTable dt)
        {
            d3 = 0;
            DataTable dtx2 = new DataTable();
            dtx2.Columns.Add("序号", typeof(string));
            dtx2.Columns.Add("表面加工", typeof(string));
            dtx2.Columns.Add("几款", typeof(string));
            dtx2.Columns.Add("小计", typeof(string));
            DataTable dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dt, "表面加工");
            if (dtx3.Rows.Count > 0)
            {
                i = 1;
                foreach (DataRow dr1 in dtx3.Rows)
                {

                    d1 = 0;
                    d2 = 0;
                    DataRow dr = dtx2.NewRow();
                    dr["表面加工"] = dr1["VALUE"].ToString();
                    dr["序号"] = i;
                    DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt, "", "表面加工='" + dr1["VALUE"].ToString() + "'");
                    if (dtx4.Rows.Count > 0)
                    {
                        dr["几款"] = dtx4.Rows.Count;
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(表面加工小计)", "").ToString()))
                        {
                            d1 = decimal.Parse(dtx4.Compute("SUM(表面加工小计)", "").ToString());
                        }
                     
                        dr["小计"] = d1;
                    }
                    d3 = d1 + d3;
                    dtx2.Rows.Add(dr);
                    i = i + 1;
                }
            }
            DataRow dr2 = dtx2.NewRow();
            dr2["表面加工"] = "表面处理";
            dr2["小计"] = d3;
            dtx2.Rows.Add(dr2);
            DataRow dr3 = dtx2.NewRow();
            if (!string.IsNullOrEmpty(dt.Rows[0]["数量"].ToString()))
            {
                dr3["小计"] = d3 / decimal.Parse(dt.Rows[0]["数量"].ToString());
            }
            dtx2.Rows.Add(dr3);
            return dtx2;
        }
        #endregion
        #region RETURN_LAMINATING_PROCESS
        public DataTable RETURN_LAMINATING_PROCESS(DataTable dt)
        {
            d3 = 0;
            DataTable dtx2 = new DataTable();
            dtx2.Columns.Add("序号", typeof(string));
            dtx2.Columns.Add("裱纸工艺", typeof(string));
            dtx2.Columns.Add("几款", typeof(string));
            dtx2.Columns.Add("小计", typeof(string));
            DataTable dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dt, "裱纸工艺");
            if (dtx3.Rows.Count > 0)
            {
                i = 1;
                foreach (DataRow dr1 in dtx3.Rows)
                {

                    d1 = 0;
                    d2 = 0;
                    DataRow dr = dtx2.NewRow();
                    dr["裱纸工艺"] = dr1["VALUE"].ToString();
                    dr["序号"] = i;
                    DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt, "", "裱纸工艺='" + dr1["VALUE"].ToString() + "'");
                    if (dtx4.Rows.Count > 0)
                    {
                        dr["几款"] = dtx4.Rows.Count;
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(裱工小计)", "").ToString()))
                        {
                            d1 = decimal.Parse(dtx4.Compute("SUM(裱工小计)", "").ToString());
                        }

                        dr["小计"] = d1;
                    }
                    d3 = d1 + d3;
                    dtx2.Rows.Add(dr);
                    i = i + 1;
                }
            }
            DataRow dr2 = dtx2.NewRow();
            dr2["裱纸工艺"] = "裱纸汇总";
            dr2["小计"] = d3;
            dtx2.Rows.Add(dr2);
            DataRow dr3 = dtx2.NewRow();
            if (!string.IsNullOrEmpty(dt.Rows[0]["数量"].ToString()))
            {
                dr3["小计"] = d3 / decimal.Parse(dt.Rows[0]["数量"].ToString());
            }
            dtx2.Rows.Add(dr3);
            return dtx2;
        }
        #endregion
        #region RETURN_DIE_CUTTING
        public DataTable RETURN_DIE_CUTTING(DataTable dt)
        {
            d3 = 0;
            DataTable dtx2 = new DataTable();
            dtx2.Columns.Add("序号", typeof(string));
            dtx2.Columns.Add("机器型号", typeof(string));
            dtx2.Columns.Add("几款", typeof(string));
            dtx2.Columns.Add("小计", typeof(string));
            DataTable dt1 = bc.GET_DT_TO_DV_TO_DT(dt, "", "机器型号 NOT IN ('主体','汇总') AND 模切='是'");
            DataTable dtx3 = bc.RETURN_NOHAVE_REPEAT_DT(dt1, "机器型号");
            if (dtx3.Rows.Count > 0)
            {
                i = 1;
                foreach (DataRow dr1 in dtx3.Rows)
                {

                    d1 = 0;
                    d2 = 0;
                    DataRow dr = dtx2.NewRow();
                    dr["机器型号"] = dr1["VALUE"].ToString();
                    dr["序号"] = i;
                    DataTable dtx4 = bc.GET_DT_TO_DV_TO_DT(dt1, "", "机器型号='" + dr1["VALUE"].ToString() + "'");
                    if (dtx4.Rows.Count > 0)
                    {
                        dr["几款"] = dtx4.Rows.Count;
                        if (!string.IsNullOrEmpty(dtx4.Compute("SUM(模切小计)", "").ToString()))
                        {
                            d1 = decimal.Parse(dtx4.Compute("SUM(模切小计)", "").ToString());
                        }

                        dr["小计"] = d1;
                    }
                    d3 = d1 + d3;
                    dtx2.Rows.Add(dr);
                    i = i + 1;
                }
            }
            DataRow dr2 = dtx2.NewRow();
            dr2["机器型号"] = "印刷汇总";
            dr2["小计"] = d3;
            dtx2.Rows.Add(dr2);
            DataRow dr3 = dtx2.NewRow();
            if (!string.IsNullOrEmpty(dt.Rows[0]["数量"].ToString()))
            {
                dr3["小计"] = d3 / decimal.Parse(dt.Rows[0]["数量"].ToString());
                //此时的DT需为含所有内容的DT,含模切为空或否的情况，
                //如果只有一行且模切为空或是否，去除该行将0行无值 16/01/20
            }
            dtx2.Rows.Add(dr3);
            return dtx2;
        }
        #endregion
        #region ExcelPrint_FOR_BASEINFO_PURCHASE
        public void ExcelPrint_FOR_BASEINFO_PURCHASE(DataTable dtt, string BillName, string Printpath)
        {
            //int j;
            PFID = dtt.Rows[0]["报价ID"].ToString();
            dt = bc.getdt(sql + " WHERE A.PFID='"+PFID +"'");
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);
            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
            dt2 = bc.getdt(cprint_die_cutting.sql + " WHERE A.PFID='" + PFID + "'");/*dgv2 start*/
            for (i = 0; i < dt2.Rows.Count - 1; i++)
            {
              
                worksheet.Cells[4 + i, "B"] = dt2.Rows[i]["项目"].ToString();
                worksheet.Cells[4 + i, "C"] = dt2.Rows[i]["刀模长米"].ToString();
                worksheet.Cells[4 + i, "D"] = dt2.Rows[i]["元米"].ToString();
                worksheet.Cells[4 + i, "E"] = dt2.Rows[i]["圆孔个数"].ToString();
                worksheet.Cells[4 + i, "F"] = dt2.Rows[i]["元个"].ToString();
                worksheet.Cells[4 + i, "G"] = dt2.Rows[i]["小计"].ToString();
            }
         
            worksheet.Cells[6, "E"] = dt2.Rows[2]["圆孔个数"].ToString();
            worksheet.Cells[6, "D"] = dt2.Rows[2]["元米"].ToString();/*dgv2 end */
            //MessageBox.Show(PFID);
            dt3 = bc.getdt(cprint_portray.sql + " WHERE A.PFID='" + PFID + "'");/*dgv3 start*/
            for (i = 0; i < dt3.Rows.Count - 1; i++)
            {
                worksheet.Cells[8 + i, "B"] = dt3.Rows[i]["写真类型"].ToString();
                worksheet.Cells[8 + i, "C"] = dt3.Rows[i]["长"].ToString();
                worksheet.Cells[8 + i, "D"] = dt3.Rows[i]["宽"].ToString();
                worksheet.Cells[8 + i, "E"] = dt3.Rows[i]["总数量"].ToString();
                worksheet.Cells[8 + i, "F"] = dt3.Rows[i]["单价"].ToString();
                worksheet.Cells[8 + i, "G"] = dt3.Rows[i]["小计"].ToString();
            }
            if (dt3.Rows.Count > 0)
            {
                worksheet.Cells[8 + i, "F"] = dt3.Rows[i]["单价"].ToString();
            }
         
            /*dgv3 end */
            dt4 = RETURN_PARTS_AUXILIAR_DT (PFID , dtt);/*dgv4 start*/
            for (i = 0; i < dt4.Rows.Count - 1; i++)
            {
                worksheet.Cells[4 + i, "I"] = dt4.Rows[i]["配件名"].ToString();
                worksheet.Cells[4 + i, "J"] = dt4.Rows[i]["用量"].ToString();
                worksheet.Cells[4 + i, "K"] = dt4.Rows[i]["单价"].ToString();
                worksheet.Cells[4 + i, "L"] = dt4.Rows[i]["单位"].ToString();
                worksheet.Cells[4 + i, "M"] = dt4.Rows[i]["小计"].ToString();
                worksheet.Cells[4 + i, "N"] = dt4.Rows[i]["备注"].ToString();
            }/*dgv4 end */
            dt5 =RETURN_PACK_MATERIAL_DT (PFID , dtt);/*dgv5 start*/
            for (i = 0; i < dt5.Rows.Count - 1; i++)
            {
                worksheet.Cells[4 + i, "Q"] = dt5.Rows[i]["项目"].ToString();
                worksheet.Cells[4 + i, "R"] = dt5.Rows[i]["数量"].ToString();
                worksheet.Cells[4 + i, "S"] = dt5.Rows[i]["长"].ToString();
                worksheet.Cells[4 + i, "T"] = dt5.Rows[i]["宽"].ToString();
                worksheet.Cells[4 + i, "U"] = dt5.Rows[i]["高"].ToString();
                worksheet.Cells[4 + i, "V"] = dt5.Rows[i]["箱形"].ToString();
                worksheet.Cells[4 + i, "W"] = dt5.Rows[i]["材质"].ToString();
                worksheet.Cells[4 + i, "X"] = dt5.Rows[i]["单价"].ToString();
                worksheet.Cells[4 + i, "Y"] = dt5.Rows[i]["小计"].ToString();
            }/*dgv5 end */
            dt6 =RETURN_ARTIFICIAL_DT (PFID, dtt);/*dgv6 start*/
            for (i = 0; i < dt6.Rows.Count; i++)
            {
                worksheet.Cells[19 + i, "C"] = dt6.Rows[i]["项目"].ToString();
                worksheet.Cells[19 + i, "D"] = dt6.Rows[i]["单价"].ToString();
                worksheet.Cells[19 + i, "E"] = dt6.Rows[i]["数量"].ToString();
                worksheet.Cells[19 + i, "F"] = dt6.Rows[i]["小计"].ToString();
                worksheet.Cells[19 + i, "G"] = dt6.Rows[i]["元套"].ToString();
      
            }/*dgv6 end */
            dt7 = bc.getdt(cprint_purchase.sql + " WHERE A.PFID='" + PFID + "'");/*dgv7 start*/
            for (i = 0; i < dt7.Rows.Count - 1; i++)
            {
                worksheet.Cells[19 + i, "H"] = dt7.Rows[i]["类型一"].ToString();
                worksheet.Cells[19 + i, "I"] = dt7.Rows[i]["外购价一"].ToString();
                worksheet.Cells[19 + i, "J"] = dt7.Rows[i]["管理费一"].ToString();
                worksheet.Cells[19 + i, "K"] = dt7.Rows[i]["小计一"].ToString();
                worksheet.Cells[19 + i, "L"] = dt7.Rows[i]["类型二"].ToString();
                worksheet.Cells[19 + i, "M"] = dt7.Rows[i]["外购价二"].ToString();
                worksheet.Cells[19 + i, "N"] = dt7.Rows[i]["管理费二"].ToString();
                worksheet.Cells[19 + i, "O"] = dt7.Rows[i]["小计二"].ToString();

            }/*dgv7 end */
            dt8 = bc.getdt(cprint_transport.sql + " WHERE A.PFID='" + PFID + "'");/*dgv8 start*/
            for (i = 0; i < dt8.Rows.Count - 2; i++)
            {
                worksheet.Cells[16 + i, "P"] = dt8.Rows[i]["长"].ToString();
                worksheet.Cells[16 + i, "Q"] = dt8.Rows[i]["宽"].ToString();
                worksheet.Cells[16 + i, "R"] = dt8.Rows[i]["高"].ToString();
                worksheet.Cells[16 + i, "S"] = dt8.Rows[i]["总箱数"].ToString();
                worksheet.Cells[16 + i, "T"] = dt8.Rows[i]["总立方数"].ToString();
                worksheet.Cells[16 + i, "V"] = dt8.Rows[i]["运输方式"].ToString();
                worksheet.Cells[16 + i, "X"] = dt8.Rows[i]["单价"].ToString();
                worksheet.Cells[16 + i, "Y"] = dt8.Rows[i]["小计"].ToString();
      
            }/*dgv8 end */
            worksheet.Cells[2, "D"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
            worksheet.Cells[2, "H"] = dt.Rows[0]["数量"].ToString();
            worksheet.Cells[2, "L"] = dt.Rows[0]["项目号"].ToString();
            worksheet.Cells[2, "P"] = dt.Rows[0]["报价编号"].ToString();
            worksheet.Cells[2, "U"] = dt.Rows[0]["报价"].ToString();
            worksheet.Cells[2, "X"] = dt.Rows[0]["日期"].ToString();
            worksheet.get_Range(worksheet.Cells[24, "B"], worksheet.Cells[24 + dt.Rows .Count-1 , "Y"]).Font.Size = 10;
            worksheet.get_Range(worksheet.Cells[24, "B"], worksheet.Cells[24 + dt.Rows.Count - 1, "Y"]).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            for (i = 0; i < dt.Rows.Count; i++)
            {
                worksheet.Cells[24 + i, "B"] = (i + 1).ToString();
                worksheet.Cells[24 + i, "C"] = dt.Rows[i]["部品名"].ToString();
                worksheet.Cells[24 + i, "D"] = dt.Rows[i]["图纸门幅"].ToString();
                worksheet.Cells[24 + i, "E"] = dt.Rows[i]["图纸纸长"].ToString();
                worksheet.Cells[24 + i, "F"] = dt.Rows[i]["部品数"].ToString();
                worksheet.Cells[24 + i, "G"] = dt.Rows[i]["印刷选项"].ToString();
                worksheet.Cells[24 + i, "H"] = dt.Rows[i]["面纸"].ToString();
                worksheet.Cells[24 + i, "I"] = dt.Rows[i]["面纸克重"].ToString();
                worksheet.Cells[24 + i, "J"] = dt.Rows[i]["芯纸"].ToString();
                worksheet.Cells[24 + i, "K"] = dt.Rows[i]["芯纸规格"].ToString();
                worksheet.Cells[24 + i, "L"] = dt.Rows[i]["底纸"].ToString();
                worksheet.Cells[24 + i, "M"] = dt.Rows[i]["底纸克重"].ToString();
                worksheet.Cells[24 + i, "N"] = dt.Rows[i]["正面4C"].ToString();
                worksheet.Cells[24 + i, "O"] = dt.Rows[i]["正面专色"].ToString();
                worksheet.Cells[24 + i, "P"] = dt.Rows[i]["正面防晒"].ToString();
                worksheet.Cells[24 + i, "Q"] = dt.Rows[i]["双面印刷"].ToString();
                worksheet.Cells[24 + i, "R"] = dt.Rows[i]["反面4C"].ToString();
                worksheet.Cells[24 + i, "S"] = dt.Rows[i]["反面专色"].ToString();
                worksheet.Cells[24 + i, "T"] = dt.Rows[i]["反面防晒"].ToString();
                worksheet.Cells[24 + i, "U"] = dt.Rows[i]["表面加工"].ToString();
                worksheet.Cells[24 + i, "V"] = dt.Rows[i]["表面次数"].ToString();
                worksheet.Cells[24 + i, "W"] = dt.Rows[i]["裱纸工艺"].ToString();
                worksheet.Cells[24 + i, "X"] = dt.Rows[i]["裱纸次数"].ToString();
                worksheet.Cells[24 + i, "Y"] = dt.Rows[i]["模切"].ToString();
          
               
            }/*dgv1-1/2*/

            worksheet.get_Range(worksheet.Cells[24, "B"], worksheet.Cells[24 + i-1, "Y"]).Borders.LineStyle = 1;
          
         
        }
        #endregion
        #region ExcelPrint_FOR_BASEINFO_AE
        public void ExcelPrint_FOR_BASEINFO_AE(DataTable dtt, string BillName, string Printpath)
        {
            //int j;
            decimal  x1 = 0,x2=0;
            PFID = dtt.Rows[0]["报价ID"].ToString();
            dtx = bc.getdt(cprint_cost_total.sql + " WHERE A.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
             
                dtx = cprint_cost_total.RETURN_DT(dtx);
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR (dtx.Rows[4]["主件用量"].ToString(),'%')))
                {
                    x1 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[4]["主件用量"].ToString(), '%'));//辅材单价
                }
             
            }
            sqb = new StringBuilder();
            sqb.AppendFormat(" WHERE A.PROJECT_NAME='{0}' AND C.CNAME='{1}' ", "代购管理", dtt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dtt.Rows[0]["品牌"].ToString());
            DataTable dtx1 = bc.getdt(cother_cost.sql + sqb.ToString());
            if (dtx1.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dtx1.Rows[0]["客户比例"].ToString(), '%')))
                {
                    x2 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx1.Rows[0]["客户比例"].ToString(), '%'));//代购管理
                }
                //MessageBox.Show(sqb.ToString());
            }
        
            dt = bc.getdt(sql + " WHERE A.PFID='" + PFID + "'");
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);
            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
        
            dt2 = bc.getdt(cprint_die_cutting.sql + " WHERE A.PFID='" + PFID + "'");/*dgv2 start*/
            decimal d11 = 0;
            for (i = 0; i < dt2.Rows.Count - 1; i++)
            {
              
                worksheet.Cells[4 + i, "B"] = dt2.Rows[i]["项目"].ToString();
                worksheet.Cells[4 + i, "C"] = dt2.Rows[i]["刀模长米"].ToString();
                if (dt2.Rows[i]["元米"].ToString() != "")
                {
                    d11 = decimal.Parse(dt2.Rows[i]["元米"].ToString());
                    worksheet.Cells[4 + i, "D"] = d11 *(1+x1/100);
                }
                worksheet.Cells[4 + i, "E"] = dt2.Rows[i]["圆孔个数"].ToString();
                if (dt2.Rows[i]["元个"].ToString() != "")
                {
                    d11 = decimal.Parse(dt2.Rows[i]["元个"].ToString());
                    worksheet.Cells[4 + i, "F"] = d11 * (1 + x1 / 100);
                }
                if (dt2.Rows[i]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt2.Rows[i]["小计"].ToString());
                    worksheet.Cells[4 + i, "G"] = d11 * (1 + x1 / 100);
                }
              
            }
            if (dt2.Rows.Count > 0)
            {
                if (dt2.Rows[2]["圆孔个数"].ToString() != "")
                {
                    d11 = decimal.Parse(dt2.Rows[2]["圆孔个数"].ToString());

                    worksheet.Cells[6, "E"] = d11 * (1 + x1 / 100);
                }

                if (dt2.Rows[2]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt2.Rows[2]["小计"].ToString());

                    worksheet.Cells[6, "G"] = d11 * (1 + x1 / 100);
                }
            }
        
           /*dgv2 end */
          
            dt3 = bc.getdt(cprint_portray.sql + " WHERE A.PFID='" + PFID + "'");/*dgv3 start*/
            for (i = 0; i < dt3.Rows.Count - 1; i++)
            {
                worksheet.Cells[8 + i, "B"] = dt3.Rows[i]["写真类型"].ToString();
                worksheet.Cells[8 + i, "C"] = dt3.Rows[i]["长"].ToString();
                worksheet.Cells[8 + i, "D"] = dt3.Rows[i]["宽"].ToString();
                worksheet.Cells[8 + i, "E"] = dt3.Rows[i]["总数量"].ToString();
     

                if (dt3.Rows[i]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt3.Rows[i]["单价"].ToString());

                    worksheet.Cells[8 + i, "F"] = d11 * (1 + x1 / 100);
                }
                if (dt3.Rows[i]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt3.Rows[i]["小计"].ToString());

                    worksheet.Cells[8 + i, "G"] = d11 * (1 + x1 / 100);
                }
            }
            if (dt3.Rows.Count > 0)
            {
                if (dt3.Rows[dt3.Rows.Count - 1]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt3.Rows[dt3.Rows.Count - 1]["单价"].ToString());
                    worksheet.Cells[17, "F"] = d11 * (1 + x1 / 100);

                }
            }
    
            /*dgv3 end */
            dt4 = RETURN_PARTS_AUXILIAR_DT(PFID, dtt);/*dgv4 start*/
            for (i = 0; i < dt4.Rows.Count - 1; i++)
            {
                worksheet.Cells[4 + i, "I"] = dt4.Rows[i]["配件名"].ToString();
                worksheet.Cells[4 + i, "J"] = dt4.Rows[i]["用量"].ToString();
                worksheet.Cells[4 + i, "L"] = dt4.Rows[i]["单位"].ToString();
                if (dt4.Rows[i]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt4.Rows[i]["单价"].ToString());
                    worksheet.Cells[4 + i, "K"] = d11 * (1 + x1 / 100);
                }
                if (dt4.Rows[i]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt4.Rows[i]["小计"].ToString());
                    worksheet.Cells[4 + i, "M"] = d11 * (1 + x1 / 100);
                }
                worksheet.Cells[4 + i, "N"] = dt4.Rows[i]["备注"].ToString();
            }
            if (dt4.Rows.Count > 0)
            {
                if (dt4.Rows[dt4.Rows.Count - 1]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt4.Rows[dt4.Rows.Count - 1]["单价"].ToString());
                    worksheet.Cells[17, "L"] = d11 * (1 + x1 / 100);

                }
            }
            /*dgv4 end */
            dt5 = RETURN_PACK_MATERIAL_DT(PFID, dtt);/*dgv5 start*/
            for (i = 0; i < dt5.Rows.Count - 1; i++)
            {
                worksheet.Cells[4 + i, "Q"] = dt5.Rows[i]["项目"].ToString();
                worksheet.Cells[4 + i, "R"] = dt5.Rows[i]["数量"].ToString();
                worksheet.Cells[4 + i, "S"] = dt5.Rows[i]["长"].ToString();
                worksheet.Cells[4 + i, "T"] = dt5.Rows[i]["宽"].ToString();
                worksheet.Cells[4 + i, "U"] = dt5.Rows[i]["高"].ToString();
                worksheet.Cells[4 + i, "V"] = dt5.Rows[i]["箱形"].ToString();
                worksheet.Cells[4 + i, "W"] = dt5.Rows[i]["材质"].ToString();
          
                if (dt5.Rows[i]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt5.Rows[i]["单价"].ToString());
                    worksheet.Cells[4 + i, "X"] = d11 * (1 + x1 / 100);
                }
                if (dt5.Rows[i]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt5.Rows[i]["小计"].ToString());
                    worksheet.Cells[4 + i, "Y"] = d11 * (1 + x1 / 100);
                }
            }
            if (dt5.Rows.Count > 0)
            {
                if (dt5.Rows[dt5.Rows.Count - 1]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt5.Rows[dt5.Rows.Count - 1]["单价"].ToString());
                    worksheet.Cells[14, "X"] = d11 * (1 + x1 / 100);

                }
            }
            /*dgv5 end */
            dt6 = RETURN_ARTIFICIAL_DT(PFID, dtt);/*dgv6 start*/
            for (i = 0; i < dt6.Rows.Count; i++)
            {
                worksheet.Cells[19 + i, "C"] = dt6.Rows[i]["项目"].ToString();

                worksheet.Cells[19 + i, "E"] = dt6.Rows[i]["数量"].ToString();
     
                if (dt6.Rows[i]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt6.Rows[i]["单价"].ToString());
                    worksheet.Cells[19 + i, "D"] = d11 * (1 + x1 / 100);
                }
                if (dt6.Rows[i]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt6.Rows[i]["小计"].ToString());
                    worksheet.Cells[19 + i, "F"] = d11 * (1 + x1 / 100);
                }
                if (dt6.Rows[i]["元套"].ToString() != "")
                {
                    d11 = decimal.Parse(dt6.Rows[i]["元套"].ToString());
                    worksheet.Cells[19 + i, "G"] = (d11 * (1 + x1 / 100));
                }
            }
   
            /*dgv6 end */
            worksheet.get_Range(worksheet.Cells[19, "I"], worksheet.Cells[20, "O"]).NumberFormat = "0.00";
            dt7 =bc.getdt(cprint_purchase .sql +" WHERE A.PFID='"+PFID +"'");/*dgv7 start*/
            for (i = 0; i < dt7.Rows.Count -1; i++)
            {
                worksheet.Cells[19 + i, "H"] = dt7.Rows[i]["类型一"].ToString();

                if (!string.IsNullOrEmpty(dt7.Rows[i]["外购价一"].ToString()))
                {
                    worksheet.Cells[19 + i, "I"] = (decimal.Parse(dt7.Rows[i]["小计一"].ToString()) /(1 + x2 / 100));//16/01/13
                }
                worksheet.Cells[19 + i, "L"] = dt7.Rows[i]["类型二"].ToString();
                if (!string.IsNullOrEmpty(dt7.Rows[i]["外购价二"].ToString()))
                {
                    worksheet.Cells[19 + i, "M"] = (decimal.Parse(dt7.Rows[i]["小计二"].ToString()) / (1 + x2 / 100));//16/01/13
                }

                if (dt7.Rows[i]["管理费一"].ToString() != "")
                {
                    d11 = decimal.Parse(dt7.Rows[i]["管理费一"].ToString());
                    worksheet.Cells[19 + i, "J"] = decimal.Parse(dt7.Rows[i]["小计一"].ToString()) / (1 + x2 / 100)*x2/100;//16/01/13
                }
                if (dt7.Rows[i]["小计一"].ToString() != "")
                {
                    d11 = decimal.Parse(dt7.Rows[i]["小计一"].ToString());
                    worksheet.Cells[19 + i, "K"] = d11;
                }
                if (dt7.Rows[i]["管理费二"].ToString() != "")
                {
                    d11 = decimal.Parse(dt7.Rows[i]["管理费二"].ToString());
                    worksheet.Cells[19 + i, "N"] = decimal.Parse(dt7.Rows[i]["小计二"].ToString()) / (1 + x2 / 100) * x2 / 100;//16/01/13
                }
                if (dt7.Rows[i]["小计二"].ToString() != "")
                {
                    d11 = decimal.Parse(dt7.Rows[i]["小计二"].ToString());
                    worksheet.Cells[19 + i, "O"] = d11;
                }

            }
            if (dt7.Rows.Count > 0)
            {
                if (dt7.Rows[1]["外购价二"].ToString() != "")
                {
                    d11 = decimal.Parse(dt7.Rows[1]["外购价二"].ToString());
                    worksheet.Cells[20, "N"] = (decimal.Parse(dt7.Rows[i]["小计二"].ToString()) / (1 + x2 / 100)).ToString ("0.00");//16/01/13

                }
            }
   
            /*dgv7 end */
            dt8 = bc.getdt(cprint_transport.sql + " WHERE A.PFID='" + PFID + "'");/*dgv8 start*/
            for (i = 0; i < dt8.Rows.Count; i++)
            {
                worksheet.Cells[16 + i, "P"] = dt8.Rows[i]["长"].ToString();
                worksheet.Cells[16 + i, "Q"] = dt8.Rows[i]["宽"].ToString();
                worksheet.Cells[16 + i, "R"] = dt8.Rows[i]["高"].ToString();
                worksheet.Cells[16 + i, "S"] = dt8.Rows[i]["总箱数"].ToString();
                worksheet.Cells[16 + i, "T"] = dt8.Rows[i]["总立方数"].ToString();
                worksheet.Cells[16 + i, "V"] = dt8.Rows[i]["运输方式"].ToString();
            

                if (dt8.Rows[i]["单价"].ToString() != "")
                {
                    d11 = decimal.Parse(dt8.Rows[i]["单价"].ToString());
                    worksheet.Cells[16 + i, "X"] = d11 * (1 + x1 / 100);
                }
                if (dt8.Rows[i]["小计"].ToString() != "")
                {
                    d11 = decimal.Parse(dt8.Rows[i]["小计"].ToString());
                    worksheet.Cells[16 + i, "Y"] = d11 * (1 + x1 / 100);
                }
            }/*dgv8 end */
            worksheet.Cells[2, "D"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
            worksheet.Cells[2, "H"] = dt.Rows[0]["数量"].ToString();
            worksheet.Cells[2, "L"] = dt.Rows[0]["项目号"].ToString();
            worksheet.Cells[2, "P"] = dt.Rows[0]["报价编号"].ToString();
            worksheet.Cells[2, "U"] = dt.Rows[0]["报价"].ToString();
            worksheet.Cells[2, "X"] = dt.Rows[0]["日期"].ToString();
            worksheet.get_Range(worksheet.Cells[24, "B"], worksheet.Cells[24 + dt.Rows .Count -1, "Y"]).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            worksheet.get_Range(worksheet.Cells[24, "B"], worksheet.Cells[24 + dt.Rows .Count -1, "Y"]).Font.Size = 10;
            for (i = 0; i < dt.Rows.Count; i++)
            {
                worksheet.Cells[24 + i, "B"] = (i + 1).ToString();
                worksheet.Cells[24 + i, "C"] = dt.Rows[i]["部品名"].ToString();
                worksheet.Cells[24 + i, "D"] = dt.Rows[i]["图纸门幅"].ToString();
                worksheet.Cells[24 + i, "E"] = dt.Rows[i]["图纸纸长"].ToString();
                worksheet.Cells[24 + i, "F"] = dt.Rows[i]["部品数"].ToString();
                worksheet.Cells[24 + i, "G"] = dt.Rows[i]["印刷选项"].ToString();
                worksheet.Cells[24 + i, "H"] = dt.Rows[i]["面纸"].ToString();
                worksheet.Cells[24 + i, "I"] = dt.Rows[i]["面纸克重"].ToString();
                worksheet.Cells[24 + i, "J"] = dt.Rows[i]["芯纸"].ToString();
                worksheet.Cells[24 + i, "K"] = dt.Rows[i]["芯纸规格"].ToString();
                worksheet.Cells[24 + i, "L"] = dt.Rows[i]["底纸"].ToString();
                worksheet.Cells[24 + i, "M"] = dt.Rows[i]["底纸克重"].ToString();
                worksheet.Cells[24 + i, "N"] = dt.Rows[i]["正面4C"].ToString();
                worksheet.Cells[24 + i, "O"] = dt.Rows[i]["正面专色"].ToString();
                worksheet.Cells[24 + i, "P"] = dt.Rows[i]["正面防晒"].ToString();
                worksheet.Cells[24 + i, "Q"] = dt.Rows[i]["双面印刷"].ToString();
                worksheet.Cells[24 + i, "R"] = dt.Rows[i]["反面4C"].ToString();
                worksheet.Cells[24 + i, "S"] = dt.Rows[i]["反面专色"].ToString();
                worksheet.Cells[24 + i, "T"] = dt.Rows[i]["反面防晒"].ToString();
                worksheet.Cells[24 + i, "U"] = dt.Rows[i]["表面加工"].ToString();
                worksheet.Cells[24 + i, "V"] = dt.Rows[i]["表面次数"].ToString();
                worksheet.Cells[24 + i, "W"] = dt.Rows[i]["裱纸工艺"].ToString();
                worksheet.Cells[24 + i, "X"] = dt.Rows[i]["裱纸次数"].ToString();
                worksheet.Cells[24 + i, "Y"] = dt.Rows[i]["模切"].ToString();
             

            }/*dgv1-1/2*/
            worksheet.get_Range(worksheet.Cells[24, "B"], worksheet.Cells[24 + i-1, "Y"]).Borders.LineStyle = 1;
           
        }
        #endregion
        #region ExcelPrint_FOR_NUCLEAR_PRICE
        public void ExcelPrint_FOR_NUCLEAR_PRICE(DataTable dt, string BillName, string Printpath)
        {
         
            PFID = dt.Rows[0]["报价ID"].ToString();
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);
            
            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[2, "D"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
                worksheet.Cells[2, "J"] = dt.Rows[0]["数量"].ToString();
                worksheet.Cells[2, "P"] = dt.Rows[0]["项目号"].ToString();
                worksheet.Cells[2, "W"] = dt.Rows[0]["报价编号"].ToString();
                worksheet.Cells[2, "AC"] = dt.Rows[0]["报价"].ToString();
                worksheet.Cells[2, "AF"] = dt.Rows[0]["日期"].ToString();
            }
            DataTable dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.PFID='" +PFID + "'");
            if (dtx.Rows.Count > 0)
            {
              
                dt7= cprint_cost_total.RETURN_DT(dtx);
            }
            if (dt7.Rows.Count > 0)
            {
               
                for (i = 0; i < dt7.Rows.Count; i++)
                {
                   
          
                    if (i ==dt7.Rows.Count - 1)
                    {
                        if (dt7.Rows[i]["项目"].ToString() == "无外购采购比" && (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dt7.Rows[i]["元套"].ToString(), '%'))))
                       {
                        worksheet.Cells[i + 5, "D"] = decimal.Parse(bc.RETURN_UNTIL_CHAR(dt7.Rows[i]["元套"].ToString(), '%')) / 100;
                     
                          }
                        if (dt7.Rows[i]["项目"].ToString() == "无外购采购比" && (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dt7.Rows[i]["主件用量"].ToString(), '%'))))
                       {

                        worksheet.Cells[i + 5, "F"] = decimal.Parse(bc.RETURN_UNTIL_CHAR(dt7.Rows[i]["主件用量"].ToString(), '%')) / 100;
                       }

                    }
                    else if (dt7.Rows[i]["项目"].ToString() == "未税" || dt7.Rows[i]["项目"].ToString() == "含税")
                    {
                        worksheet.Cells[i + 5, "B"] = dt7.Rows[i]["项目"].ToString();
                        worksheet.Cells[i + 5, "C"] = dt7.Rows[i]["元套"].ToString();
                        worksheet.Cells[i + 5, "E"] = dt7.Rows[i]["批量小计"].ToString();
                    }

                    else
                    {
                        worksheet.Cells[i + 5, "D"] = dt7.Rows[i]["元套"].ToString();
                        worksheet.Cells[i + 5, "C"] = dt7.Rows[i]["项目"].ToString();
                        worksheet.Cells[i + 5, "F"] = dt7.Rows[i]["主件用量"].ToString();
                        worksheet.Cells[i + 5, "E"] = dt7.Rows[i]["批量小计"].ToString();

                    }
                }
            }

            dt1 = RETURN_PAPER_TOTAL(dt);/*PAPER _START*/
       
            if (dt1.Rows.Count > 0)
            {
                for (i = 0; i < dt1.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "G"] = dt1.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "H"] = dt1.Rows[i]["面纸"].ToString();
                    worksheet.Cells[i + 6, "I"] = dt1.Rows[i]["面纸克重"].ToString();
                    worksheet.Cells[i + 6, "J"] = dt1.Rows[i]["面纸单价"].ToString();
                    worksheet.Cells[i + 6, "K"] = dt1.Rows[i]["面纸单个用量"].ToString();
               
                        worksheet.Cells[i + 6, "L"] = dt1.Rows[i]["面纸小计"].ToString();
                    
                    worksheet.get_Range(worksheet.Cells[i + 6, "I"], worksheet.Cells[i + 6, "L"]).HorizontalAlignment = 
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "J"], worksheet.Cells[i + 6, "J"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "K"]).NumberFormat = "0.000";
              
                    if (i == dt1.Rows.Count - 2 || i == dt1.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "I"], worksheet.Cells[i + 6, "L"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "L"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "L"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "L"]).NumberFormat = "0";

                    }
                    if (i == dt1.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "L"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "G"], worksheet.Cells[6+ i, "L"]).Borders.LineStyle = 1;
                }
               
            }/*PAPER _END*/
            dt2 = RETURN_PAPER_CORE_TOTAL(dt);/*PAPER_CORE _START*/
            if (dt2.Rows.Count > 0)
            {
                for (i = 0; i < dt2.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "M"] = dt2.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "N"] = dt2.Rows[i]["芯纸"].ToString();
                    worksheet.Cells[i + 6, "O"] = dt2.Rows[i]["芯纸规格"].ToString();
                    worksheet.Cells[i + 6, "P"] = dt2.Rows[i]["芯纸单价"].ToString();
                    worksheet.Cells[i + 6, "Q"] = dt2.Rows[i]["芯纸单个用量"].ToString();
                    worksheet.Cells[i + 6, "R"] = dt2.Rows[i]["芯纸小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "R"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "P"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "Q"]).NumberFormat = "0.000";
                   
                    if (i == dt2.Rows.Count - 2 || i == dt2.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "O"], worksheet.Cells[i + 6, "R"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "R"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "R"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "R"], worksheet.Cells[i + 6, "R"]).NumberFormat = "0";
                    }
                    if (i == dt2.Rows.Count - 1)
                    {

                        worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "R"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "M"], worksheet.Cells[6 + i, "R"]).Borders.LineStyle = 1;
                }
            }/*PAPER_CORE _END*/
            dt3 = RETURN_PRINTING_TOTAL(dt);/*PRINTING _START*/
            if (dt3.Rows.Count > 0)
            {
                for (i = 0; i < dt3.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "S"] = dt3.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "T"] = dt3.Rows[i]["机器型号"].ToString();
                    worksheet.Cells[i + 6, "U"] = dt3.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "V"] = dt3.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
              
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).NumberFormat = "0";

                    if (i == dt3.Rows.Count - 2 || i == dt3.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "T"], worksheet.Cells[i + 6, "V"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "V"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "V"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "U"]).NumberFormat = "0";
                    }
                    if (i == dt3.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "V"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "S"], worksheet.Cells[6 + i, "V"]).Borders.LineStyle = 1;
                }
            }/*PRINTING _END*/
            dt4 = RETURN_SURFACE_MACHINING(dt);/*SURFACE_MACHINING _START*/
            if (dt4.Rows.Count > 0)
            {
                for (i = 0; i < dt4.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "W"] = dt4.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "X"] = dt4.Rows[i]["表面加工"].ToString();
                    worksheet.Cells[i + 6, "Y"] = dt4.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "Z"] = dt4.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "0";

                    if (i == dt4.Rows.Count - 2 || i == dt4.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "X"], worksheet.Cells[i + 6, "Z"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Z"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Y"]).NumberFormat = "0";
                    }

                    if ( i == dt4.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "W"], worksheet.Cells[6 + i, "Z"]).Borders.LineStyle = 1;
                }
            }/*SURFACE_MACHINING _END*/
            dt5 = RETURN_LAMINATING_PROCESS(dt);/*LAMINATING_PROCESS _START*/
            if (dt5.Rows.Count > 0)
            {
                for (i = 0; i < dt5.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "AA"] = dt5.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "AB"] = dt5.Rows[i]["裱纸工艺"].ToString();
                    worksheet.Cells[i + 6, "AC"] = dt5.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "AD"] = dt5.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "AD"], worksheet.Cells[i + 6, "AD"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                   
                    worksheet.get_Range(worksheet.Cells[i + 6, "AD"], worksheet.Cells[i + 6, "AD"]).NumberFormat = "0";

                    if (i == dt5.Rows.Count - 2 || i == dt5.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AB"], worksheet.Cells[i + 6, "AD"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AD"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AD"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AC"]).NumberFormat = "0";
                    }
                    if ( i == dt5.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AD"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "AA"], worksheet.Cells[6 + i, "AD"]).Borders.LineStyle = 1;
                }
            }/*LAMINATING_PROCESS _END*/
            dt6 = RETURN_DIE_CUTTING(dt);/*DIE_CUTTING _START*/
            if (dt6.Rows.Count > 0)
            {
                for (i = 0; i < dt6.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "AE"] = dt6.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "AF"] = dt6.Rows[i]["机器型号"].ToString();
                    worksheet.Cells[i + 6, "AG"] = dt6.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "AH"] = dt6.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "AH"], worksheet.Cells[i + 6, "AH"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                 
                    worksheet.get_Range(worksheet.Cells[i + 6, "AH"], worksheet.Cells[i + 6, "AH"]).NumberFormat = "0";

                    if (i == dt6.Rows.Count - 2 || i == dt6.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AF"], worksheet.Cells[i + 6, "AH"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AH"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AH"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AG"]).NumberFormat = "0";
                    }
                    if ( i == dt6.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AH"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "AE"], worksheet.Cells[6 + i, "AH"]).Borders.LineStyle = 1;
                }
            }/*DIE_CUTTING _END*/
            worksheet.get_Range(worksheet.Cells[32, "B"], worksheet.Cells[32 + dt.Rows .Count -1, "F"]).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            worksheet.get_Range(worksheet.Cells[32, "G"], worksheet.Cells[32 + dt.Rows .Count -1, "AH"]).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
            worksheet.get_Range(worksheet.Cells[32, "B"], worksheet.Cells[32 + dt.Rows .Count -1, "AH"]).Font.Size = 10;
            for (i = 0; i < dt.Rows.Count; i++)
            {
                if (i != dt.Rows.Count - 2 && i != dt.Rows.Count - 1)
                {
                    worksheet.get_Range(worksheet.Cells[32 + i, "G"], worksheet.Cells[32 + i, "G"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[32 + i, "I"], worksheet.Cells[32 + i, "I"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[32 + i, "P"], worksheet.Cells[32 + i, "P"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[32 + i, "W"], worksheet.Cells[32 + i, "W"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[32 + i, "N"], worksheet.Cells[32 + i, "N"]).NumberFormat = "0.000";
                    worksheet.get_Range(worksheet.Cells[32 + i, "U"], worksheet.Cells[32 + i, "U"]).NumberFormat = "0.000";
                    worksheet.get_Range(worksheet.Cells[32 + i, "Y"], worksheet.Cells[32 + i, "Y"]).NumberFormat = "0.000";
                }
                worksheet.Cells[32 + i, "B"] = dt.Rows[i]["部品名"].ToString();
                worksheet.Cells[32 + i, "C"] = dt.Rows[i]["加工门幅"].ToString();
                worksheet.Cells[32 + i, "D"] = dt.Rows[i]["加工长度"].ToString();
                worksheet.Cells[32 + i, "E"] = dt.Rows[i]["部品总数"].ToString();
                worksheet.Cells[32 + i, "F"] = dt.Rows[i]["机器型号"].ToString();
                worksheet.Cells[32 + i, "G"] = dt.Rows[i]["部品单价"].ToString();
                worksheet.Cells[32 + i, "H"] = dt.Rows[i]["部品总价"].ToString();
                worksheet.Cells[32 + i, "I"] = dt.Rows[i]["面纸单价"].ToString();
                worksheet.Cells[32 + i, "J"] = dt.Rows[i]["面纸用量"].ToString();
                worksheet.Cells[32 + i, "K"] = dt.Rows[i]["面纸门幅"].ToString();
                worksheet.Cells[32 + i, "L"] = dt.Rows[i]["面纸纸长"].ToString();
                worksheet.Cells[32 + i, "M"] = dt.Rows[i]["面纸可用"].ToString();
           
                worksheet.Cells[32 + i, "N"] = dt.Rows[i]["面纸单个用量"].ToString();
           
                worksheet.Cells[32 + i, "O"] = dt.Rows[i]["面纸小计"].ToString();
                worksheet.Cells[32 + i, "P"] = dt.Rows[i]["芯纸单价"].ToString();
                worksheet.Cells[32 + i, "Q"] = dt.Rows[i]["芯纸用量"].ToString();
                worksheet.Cells[32 + i, "R"] = dt.Rows[i]["芯纸门幅"].ToString();
                worksheet.Cells[32 + i, "S"] = dt.Rows[i]["芯纸纸长"].ToString();
                worksheet.Cells[32 + i, "T"] = dt.Rows[i]["芯纸可用"].ToString();
                worksheet.Cells[32 + i, "U"] = dt.Rows[i]["芯纸单个用量"].ToString();
          
                worksheet.Cells[32 + i, "V"] = dt.Rows[i]["芯纸小计"].ToString();
                worksheet.Cells[32 + i, "W"] = dt.Rows[i]["底纸单价"].ToString();
                worksheet.Cells[32 + i, "X"] = dt.Rows[i]["底纸用量"].ToString();
                worksheet.Cells[32 + i, "Y"] = dt.Rows[i]["底纸单个用量"].ToString();
           
                worksheet.Cells[32 + i, "Z"] = dt.Rows[i]["底纸小计"].ToString();
                worksheet.Cells[32 + i, "AA"] = dt.Rows[i]["正反CTP合计"].ToString();
                worksheet.Cells[32 + i, "AB"] = dt.Rows[i]["正反印工合计"].ToString();
                worksheet.Cells[32 + i, "AC"] = dt.Rows[i]["表面处理单价"].ToString();
                worksheet.Cells[32 + i, "AD"] = dt.Rows[i]["表面加工小计"].ToString();
                worksheet.Cells[32 + i, "AE"] = dt.Rows[i]["裱工单价"].ToString();
                worksheet.Cells[32 + i, "AF"] = dt.Rows[i]["裱工小计"].ToString();
                worksheet.Cells[32 + i, "AG"] = dt.Rows[i]["刀模小计"].ToString();
                worksheet.Cells[32 + i, "AH"] = dt.Rows[i]["模切小计"].ToString();
                if (i == dt.Rows.Count - 2 || i == dt.Rows.Count - 1)
                {
                    worksheet.get_Range(worksheet.Cells[32 + i, "F"], worksheet.Cells[i + 32, "AH"]).Font.Bold = true;
                    worksheet.get_Range(worksheet.Cells[32 + i, "L"], worksheet.Cells[i + 32, "M"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[32 + i, "G"], worksheet.Cells[i + 32, "H"]).MergeCells = true;

                    worksheet.get_Range(worksheet.Cells[32 + i, "N"], worksheet.Cells[i + 32, "O"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[32 + i, "U"], worksheet.Cells[i + 32, "V"]).MergeCells = true;

                    worksheet.get_Range(worksheet.Cells[32 + i, "Y"], worksheet.Cells[i + 32, "Z"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[32 + i, "AA"], worksheet.Cells[i + 32, "AB"]).MergeCells = true;


                    worksheet.get_Range(worksheet.Cells[32 + i, "AC"], worksheet.Cells[i + 32, "AD"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[32 + i, "AE"], worksheet.Cells[i + 32, "AF"]).MergeCells = true;

                    worksheet.get_Range(worksheet.Cells[32 + i, "AG"], worksheet.Cells[i + 32, "AH"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[i + 32, "F"], worksheet.Cells[i + 32, "AH"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;

                }
            

            }
            worksheet.get_Range(worksheet.Cells[32, "B"], worksheet.Cells[32 + i-1, "AH"]).Borders.LineStyle = 1;
            worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[65536, 256]).Columns.AutoFit();
        }
        #endregion
        #region ExcelPrint_FOR_NUCLEAR_PURCHASE
        public void ExcelPrint_FOR_NUCLEAR_PURCHASE(DataTable dt, string BillName, string Printpath)
        {
            PFID = dt.Rows[0]["报价ID"].ToString();
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);

            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[2, "D"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
                worksheet.Cells[2, "J"] = dt.Rows[0]["数量"].ToString();
                worksheet.Cells[2, "P"] = dt.Rows[0]["项目号"].ToString();
                worksheet.Cells[2, "W"] = dt.Rows[0]["报价编号"].ToString();
                worksheet.Cells[2, "AC"] = dt.Rows[0]["报价"].ToString();
                worksheet.Cells[2, "AF"] = dt.Rows[0]["日期"].ToString();
            }
            DataTable dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
                dt7 = cprint_cost_total.RETURN_DT(dtx);
            }
            d1 = 0;
            d2 = 0;
            if (dt7.Rows.Count > 0)
            {
                dt7 = bc.GET_DT_TO_DV_TO_DT(dt7, "", "项目 NOT IN ('管理','利润','代购费','未税','含税','无外购采购比','成本合计')");
                for (i = 0; i < dt7.Rows.Count; i++)
                {
                  
                    worksheet.Cells[i + 5, "D"] = dt7.Rows[i]["元套"].ToString();
                    worksheet.Cells[i + 5, "C"] = dt7.Rows[i]["项目"].ToString();
                    worksheet.Cells[i + 5, "E"] = dt7.Rows[i]["批量小计"].ToString();
                    if (!string.IsNullOrEmpty(dt7.Rows[i]["元套"].ToString()))
                    {
                        d1 = d1 + decimal.Parse(dt7.Rows[i]["元套"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt7.Rows[i]["批量小计"].ToString()))
                    {
                        d2 = d2 + decimal.Parse(dt7.Rows[i]["批量小计"].ToString());
                    }
                }
                d10 = 0;
                string v1 = bc.getOnlyString("SELECT CUSTOMER_PERCENT FROM OTHER_COST WHERE PROJECT_NAME='税金'");
                if (!string.IsNullOrEmpty(v1))
                {
                    d10 = decimal.Parse(v1) / 100;
                }
                worksheet.Cells[18, "C"] = d1;
                worksheet.Cells[18, "E"] = d2;
                worksheet.Cells[19, "C"] = d1*(1 + d10);
                worksheet.Cells[19, "E"] = d2*(1 + d10);
            }
        
            dt1 = RETURN_PAPER_TOTAL(dt);/*PAPER _START*/
            if (dt1.Rows.Count > 0)
            {
                for (i = 0; i < dt1.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "G"] = dt1.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "H"] = dt1.Rows[i]["面纸"].ToString();
                    worksheet.Cells[i + 6, "I"] = dt1.Rows[i]["面纸克重"].ToString();
                    worksheet.Cells[i + 6, "J"] = dt1.Rows[i]["面纸单价"].ToString();
                    worksheet.Cells[i + 6, "K"] = dt1.Rows[i]["面纸单个用量"].ToString();
                    worksheet.Cells[i + 6, "L"] = dt1.Rows[i]["面纸小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "I"], worksheet.Cells[i + 6, "L"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "J"], worksheet.Cells[i + 6, "J"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "K"]).NumberFormat = "0.000";
               
                    if (i == dt1.Rows.Count - 2 || i == dt1.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "I"], worksheet.Cells[i + 6, "L"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "L"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "L"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "L"]).NumberFormat = "0";
                    }
                    if (i == dt1.Rows.Count - 1)
                    {
                   
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "L"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "G"], worksheet.Cells[6 + i, "L"]).Borders.LineStyle = 1;
                }
            }/*PAPER _END*/
         
            dt2 = RETURN_PAPER_CORE_TOTAL(dt);/*PAPER_CORE _START*/
            if (dt2.Rows.Count > 0)
            {
                for (i = 0; i < dt2.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "M"] = dt2.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "N"] = dt2.Rows[i]["芯纸"].ToString();
                    worksheet.Cells[i + 6, "O"] = dt2.Rows[i]["芯纸规格"].ToString();
                    worksheet.Cells[i + 6, "P"] = dt2.Rows[i]["芯纸单价"].ToString();
                    worksheet.Cells[i + 6, "Q"] = dt2.Rows[i]["芯纸单个用量"].ToString();
                    worksheet.Cells[i + 6, "R"] = dt2.Rows[i]["芯纸小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "R"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "P"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "Q"]).NumberFormat = "0.000";
                  
                    if (i == dt2.Rows.Count - 2 || i == dt2.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "O"], worksheet.Cells[i + 6, "R"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "R"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "R"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "R"], worksheet.Cells[i + 6, "R"]).NumberFormat = "0";
                    }
                    if (i == dt2.Rows.Count - 1)
                    {
                      
                        worksheet.get_Range(worksheet.Cells[i + 6, "Q"], worksheet.Cells[i + 6, "R"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "M"], worksheet.Cells[6 + i, "R"]).Borders.LineStyle = 1;
                }
            }/*PAPER_CORE _END*/
            //MessageBox.Show("OK");
            dt3 = RETURN_PRINTING_TOTAL(dt);/*PRINTING _START*/
            if (dt3.Rows.Count > 0)
            {
                for (i = 0; i < dt3.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "S"] = dt3.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "T"] = dt3.Rows[i]["机器型号"].ToString();
                    worksheet.Cells[i + 6, "U"] = dt3.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "V"] = dt3.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
          
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).NumberFormat = "0";

                    if (i == dt3.Rows.Count - 2 || i == dt3.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "T"], worksheet.Cells[i + 6, "V"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "V"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "V"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "U"]).NumberFormat = "0";
                    }
                    if (i == dt3.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "V"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "S"], worksheet.Cells[6 + i, "V"]).Borders.LineStyle = 1;
                }
            }/*PRINTING _END*/
            //MessageBox.Show("OK2");
            dt4 = RETURN_SURFACE_MACHINING(dt);/*SURFACE_MACHINING _START*/
            if (dt4.Rows.Count > 0)
            {
                for (i = 0; i < dt4.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "W"] = dt4.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "X"] = dt4.Rows[i]["表面加工"].ToString();
                    worksheet.Cells[i + 6, "Y"] = dt4.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "Z"] = dt4.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                 
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "0";

                    if (i == dt4.Rows.Count - 2 || i == dt4.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "X"], worksheet.Cells[i + 6, "Z"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Z"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Y"]).NumberFormat = "0";
                    }
                    if ( i == dt4.Rows.Count - 1)
                    {
             
                        worksheet.get_Range(worksheet.Cells[i + 6, "Y"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "W"], worksheet.Cells[6 + i, "Z"]).Borders.LineStyle = 1;
                }
            }/*SURFACE_MACHINING _END*/
            dt5 = RETURN_LAMINATING_PROCESS(dt);/*LAMINATING_PROCESS _START*/
            if (dt5.Rows.Count > 0)
            {
                for (i = 0; i < dt5.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "AA"] = dt5.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "AB"] = dt5.Rows[i]["裱纸工艺"].ToString();
                    worksheet.Cells[i + 6, "AC"] = dt5.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "AD"] = dt5.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "AD"], worksheet.Cells[i + 6, "AD"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                
                    worksheet.get_Range(worksheet.Cells[i + 6, "AD"], worksheet.Cells[i + 6, "AD"]).NumberFormat = "0";

                    if (i == dt5.Rows.Count - 2 || i == dt5.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AB"], worksheet.Cells[i + 6, "AD"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AD"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AD"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AC"]).NumberFormat = "0";

                    }
                    if ( i == dt5.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AC"], worksheet.Cells[i + 6, "AD"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "AA"], worksheet.Cells[6 + i, "AD"]).Borders.LineStyle = 1;
                }
            }/*LAMINATING_PROCESS _END*/
            dt6 = RETURN_DIE_CUTTING(dt);/*DIE_CUTTING _START*/
            if (dt6.Rows.Count > 0)
            {
                for (i = 0; i < dt6.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "AE"] = dt6.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "AF"] = dt6.Rows[i]["机器型号"].ToString();
                    worksheet.Cells[i + 6, "AG"] = dt6.Rows[i]["几款"].ToString();
                    worksheet.Cells[i + 6, "AH"] = dt6.Rows[i]["小计"].ToString();
                    worksheet.get_Range(worksheet.Cells[i + 6, "AH"], worksheet.Cells[i + 6, "AH"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
               
                    worksheet.get_Range(worksheet.Cells[i + 6, "AH"], worksheet.Cells[i + 6, "AH"]).NumberFormat = "0";

                    if (i == dt6.Rows.Count - 2 || i == dt6.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AF"], worksheet.Cells[i + 6, "AH"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AH"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AH"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AG"]).NumberFormat = "0";
                    }
                    if (i == dt6.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AG"], worksheet.Cells[i + 6, "AG"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "AE"], worksheet.Cells[6 + i, "AH"]).Borders.LineStyle = 1;
                }
            }/*DIE_CUTTING _END*/
            worksheet.get_Range(worksheet.Cells[27, "B"], worksheet.Cells[27 + dt.Rows .Count -1, "F"]).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            worksheet.get_Range(worksheet.Cells[27, "G"], worksheet.Cells[27 + dt.Rows .Count -1, "AH"]).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
            worksheet.get_Range(worksheet.Cells[27, "B"], worksheet.Cells[27 + dt.Rows .Count -1, "AH"]).Font.Size = 10;
            for (i = 0; i < dt.Rows.Count; i++)
            {
                if (i != dt.Rows.Count - 2 && i != dt.Rows.Count - 1)
                {
                    worksheet.get_Range(worksheet.Cells[27 + i, "G"], worksheet.Cells[27 + i, "G"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[27 + i, "I"], worksheet.Cells[27 + i, "I"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[27 + i, "P"], worksheet.Cells[27 + i, "P"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[27 + i, "W"], worksheet.Cells[27 + i, "W"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[27 + i, "N"], worksheet.Cells[27 + i, "N"]).NumberFormat = "0.000";
                    worksheet.get_Range(worksheet.Cells[27 + i, "U"], worksheet.Cells[27 + i, "U"]).NumberFormat = "0.000";
                    worksheet.get_Range(worksheet.Cells[27 + i, "Y"], worksheet.Cells[27 + i, "Y"]).NumberFormat = "0.000";
                }
                worksheet.Cells[27 + i, "B"] = dt.Rows[i]["部品名"].ToString();
                worksheet.Cells[27 + i, "C"] = dt.Rows[i]["加工门幅"].ToString();
                worksheet.Cells[27 + i, "D"] = dt.Rows[i]["加工长度"].ToString();
                worksheet.Cells[27 + i, "E"] = dt.Rows[i]["部品总数"].ToString();
                worksheet.Cells[27 + i, "F"] = dt.Rows[i]["机器型号"].ToString();
                worksheet.Cells[27 + i, "G"] = dt.Rows[i]["部品单价"].ToString();
                worksheet.Cells[27 + i, "H"] = dt.Rows[i]["部品总价"].ToString();
                worksheet.Cells[27 + i, "I"] = dt.Rows[i]["面纸单价"].ToString();
                worksheet.Cells[27 + i, "J"] = dt.Rows[i]["面纸用量"].ToString();
                worksheet.Cells[27 + i, "K"] = dt.Rows[i]["面纸门幅"].ToString();
                worksheet.Cells[27 + i, "L"] = dt.Rows[i]["面纸纸长"].ToString();
                worksheet.Cells[27 + i, "M"] = dt.Rows[i]["面纸可用"].ToString();
                worksheet.Cells[27 + i, "N"] = dt.Rows[i]["面纸单个用量"].ToString();
                worksheet.Cells[27 + i, "O"] = dt.Rows[i]["面纸小计"].ToString();
                worksheet.Cells[27 + i, "P"] = dt.Rows[i]["芯纸单价"].ToString();
                worksheet.Cells[27 + i, "Q"] = dt.Rows[i]["芯纸用量"].ToString();
                worksheet.Cells[27 + i, "R"] = dt.Rows[i]["芯纸门幅"].ToString();
                worksheet.Cells[27 + i, "S"] = dt.Rows[i]["芯纸纸长"].ToString();
                worksheet.Cells[27 + i, "T"] = dt.Rows[i]["芯纸可用"].ToString();
                worksheet.Cells[27 + i, "U"] = dt.Rows[i]["芯纸单个用量"].ToString();
                worksheet.Cells[27 + i, "V"] = dt.Rows[i]["芯纸小计"].ToString();
                worksheet.Cells[27 + i, "W"] = dt.Rows[i]["底纸单价"].ToString();
                worksheet.Cells[27 + i, "X"] = dt.Rows[i]["底纸用量"].ToString();
                worksheet.Cells[27 + i, "Y"] = dt.Rows[i]["底纸单个用量"].ToString();
                worksheet.Cells[27 + i, "Z"] = dt.Rows[i]["底纸小计"].ToString();
                worksheet.Cells[27 + i, "AA"] = dt.Rows[i]["正反CTP合计"].ToString();
                worksheet.Cells[27 + i, "AB"] = dt.Rows[i]["正反印工合计"].ToString();
                worksheet.Cells[27 + i, "AC"] = dt.Rows[i]["表面处理单价"].ToString();
                worksheet.Cells[27 + i, "AD"] = dt.Rows[i]["表面加工小计"].ToString();
                worksheet.Cells[27 + i, "AE"] = dt.Rows[i]["裱工单价"].ToString();
                worksheet.Cells[27 + i, "AF"] = dt.Rows[i]["裱工小计"].ToString();
                worksheet.Cells[27 + i, "AG"] = dt.Rows[i]["刀模小计"].ToString();
                worksheet.Cells[27 + i, "AH"] = dt.Rows[i]["模切小计"].ToString();
                if (i == dt.Rows.Count - 2 || i == dt.Rows.Count - 1)
                {
                    worksheet.get_Range(worksheet.Cells[27 + i, "F"], worksheet.Cells[i + 27, "AH"]).Font.Bold = true;
                    worksheet.get_Range(worksheet.Cells[27 + i, "L"], worksheet.Cells[i + 27, "M"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[27 + i, "G"], worksheet.Cells[i + 27, "H"]).MergeCells = true;

                    worksheet.get_Range(worksheet.Cells[27 + i, "N"], worksheet.Cells[i + 27, "O"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[27 + i, "U"], worksheet.Cells[i + 27, "V"]).MergeCells = true;

                    worksheet.get_Range(worksheet.Cells[27 + i, "Y"], worksheet.Cells[i + 27, "Z"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[27 + i, "AA"], worksheet.Cells[i + 27, "AB"]).MergeCells = true;


                    worksheet.get_Range(worksheet.Cells[27 + i, "AC"], worksheet.Cells[i + 27, "AD"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[27 + i, "AE"], worksheet.Cells[i + 27, "AF"]).MergeCells = true;

                    worksheet.get_Range(worksheet.Cells[27 + i, "AG"], worksheet.Cells[i + 27, "AH"]).MergeCells = true;
                    worksheet.get_Range(worksheet.Cells[i + 27, "F"], worksheet.Cells[i + 27, "AH"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;

                }
            

            }
            worksheet.get_Range(worksheet.Cells[27, "B"], worksheet.Cells[27 + i-1, "AH"]).Borders.LineStyle = 1;
            worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[65536, 256]).Columns.AutoFit();
        }
        #endregion
        #region ExcelPrint_FOR_MAIN_DETAIL
        public void ExcelPrint_FOR_MAIN_DETAIL(DataTable dt, string BillName, string Printpath)
        {
            PFID = dt.Rows[0]["报价ID"].ToString();
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);

            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
            decimal x1 = 0;
            decimal x2 = 0;
            dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.PFID='" + PFID + "'");
            if (dtx.Rows.Count > 0)
            {
                dtx = cprint_cost_total.RETURN_DT(dtx);
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dtx.Rows[0]["主件用量"].ToString(), '%')))//主件用量
                {
                    x2 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[0]["主件用量"].ToString(), '%'));
                }
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dtx.Rows[2]["主件用量"].ToString(), '%')))//主件单价
                {
                    x1 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[2]["主件用量"].ToString(), '%'));
                }
            }
            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[2, "D"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
                worksheet.Cells[2, "I"] = dt.Rows[0]["数量"].ToString();
                worksheet.Cells[2, "N"] = dt.Rows[0]["项目号"].ToString();
                worksheet.Cells[2, "T"] = dt.Rows[0]["报价编号"].ToString();
                worksheet.Cells[2, "Y"] = dt.Rows[0]["报价"].ToString();
                worksheet.Cells[2, "AB"] = dt.Rows[0]["日期"].ToString();
            }
      
            dt1 = RETURN_PAPER_TOTAL(dt);/*PAPER _START*/
            d1 = 0;
            d2 = 0;
            d3 = 0;
       
     
            if (dt1.Rows.Count > 0)
            {
                for (i = 0; i < dt1.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "B"] = dt1.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "C"] = dt1.Rows[i]["面纸"].ToString();
                    worksheet.Cells[i + 6, "D"] = dt1.Rows[i]["面纸克重"].ToString();
                    if (!string.IsNullOrEmpty(dt1.Rows[i]["面纸小计"].ToString()))
                    {
                        d3 = decimal.Parse(dt1.Rows[i]["面纸小计"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt1.Rows[i]["面纸单价"].ToString()) &&
                        bc.yesno(dt1.Rows[i]["面纸单价"].ToString()) != 0 && i != dt1.Rows.Count - 1 && i != dt1.Rows.Count - 2)
                    {
                        d1 = decimal.Parse(dt1.Rows[i]["面纸单价"].ToString());
                        worksheet.Cells[i + 6, "E"] = (d1 * (1 + x1 / 100)).ToString("0.00");
                        if (!string.IsNullOrEmpty(dt1.Rows[i]["面纸单个用量"].ToString()))
                        {
                            d2 = decimal.Parse(dt1.Rows[i]["面纸单个用量"].ToString());
                        }

                        worksheet.Cells[i + 6, "F"] = (d2 * (1 + x2 / 100)).ToString("0.000");
                        worksheet.Cells[i + 6, "G"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0");
                    }
                    else
                    {
                       
                   
                        worksheet.Cells[i + 6, "E"] = dt1.Rows[i]["面纸单价"].ToString();
                        worksheet.Cells[i + 6, "G"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0.00");
                    }
                
                    worksheet.get_Range(worksheet.Cells[i + 6, "C"], worksheet.Cells[i + 6, "C"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlLeft;
                    worksheet.get_Range(worksheet.Cells[i + 6, "E"], worksheet.Cells[i + 6, "F"]).HorizontalAlignment =
                      Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "E"], worksheet.Cells[i + 6, "E"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[i + 6, "F"], worksheet.Cells[i + 6, "F"]).NumberFormat = "0.000";
                  
                    worksheet.get_Range(worksheet.Cells[i + 6, "B"], worksheet.Cells[i + 6, "G"]).Font.Size = 10;
                    worksheet.get_Range(worksheet.Cells[i + 6, "B"], worksheet.Cells[i + 6, "G"]).Font.Name = "微软雅黑";

                    if (i == dt1.Rows.Count - 2 || i == dt1.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "C"], worksheet.Cells[i + 6, "G"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "F"], worksheet.Cells[i + 6, "G"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "F"], worksheet.Cells[i + 6, "G"]).HorizontalAlignment =
                  Microsoft.Office.Interop.Excel.Constants.xlRight;
                        worksheet.get_Range(worksheet.Cells[i + 6, "B"], worksheet.Cells[i + 6, "G"]).NumberFormat = "￥0";
            
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "G"], worksheet.Cells[i + 6, "G"]).NumberFormat = "0";
                    }
                    if (i == dt1.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "F"], worksheet.Cells[i + 6, "G"]).NumberFormat = "0.00";
                      
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "B"], worksheet.Cells[6 + i, "G"]).Borders.LineStyle = 1;
                }
        
            }/*PAPER _END*/
            d1 = 0;
            d2 = 0;
            d3 = 0;
            dt2 = RETURN_PAPER_CORE_TOTAL(dt);/*PAPER_CORE _START*/
            if (dt2.Rows.Count > 0)
            {
                for (i = 0; i < dt2.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "H"] = dt2.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "I"] = dt2.Rows[i]["芯纸"].ToString();
                    worksheet.Cells[i + 6, "J"] = dt2.Rows[i]["芯纸规格"].ToString();
                    if (!string.IsNullOrEmpty(dt2.Rows[i]["芯纸小计"].ToString()))
                    {
                        d3 = decimal.Parse(dt2.Rows[i]["芯纸小计"].ToString());
                    }
                    if (!string.IsNullOrEmpty(dt2.Rows[i]["芯纸单价"].ToString()) && bc.yesno(dt2.Rows[i]["芯纸单价"].ToString()) != 0
                        && i != dt2.Rows.Count - 1 && i != dt2.Rows.Count - 2)
                    {
                        d1 = decimal.Parse(dt2.Rows[i]["芯纸单价"].ToString());
                        worksheet.Cells[i + 6, "K"] = (d1 * (1 + x1 / 100)).ToString("0.00");
                        if (!string.IsNullOrEmpty(dt2.Rows[i]["芯纸单个用量"].ToString()))
                        {
                            d2 = decimal.Parse(dt2.Rows[i]["芯纸单个用量"].ToString());
                        }
                        //MessageBox.Show(dt2.Rows[i]["芯纸"].ToString() + "," + dt2.Rows[i]["芯纸规格"].ToString()+"1");
                        worksheet.Cells[i + 6, "L"] = (d2 * (1 + x2 / 100)).ToString("0.000");
                        worksheet.Cells[i + 6, "M"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0");
                    }
                    else
                    {
                        //MessageBox.Show(dt2.Rows[i]["芯纸"].ToString() + "," + dt2.Rows[i]["芯纸规格"].ToString() + "2");
                        worksheet.Cells[i + 6, "K"] = dt2.Rows[i]["芯纸单价"].ToString();
                        worksheet.Cells[i + 6, "M"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0.00");
                    }
              

              
                    worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "R"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "H"], worksheet.Cells[i + 6, "H"]).HorizontalAlignment =
                          Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "L"]).NumberFormat = "0.000";
              
                    worksheet.get_Range(worksheet.Cells[i + 6, "H"], worksheet.Cells[i + 6, "M"]).Font.Size = 10;
                    worksheet.get_Range(worksheet.Cells[i + 6, "H"], worksheet.Cells[i + 6, "M"]).Font.Name = "微软雅黑";
                    if (i == dt2.Rows.Count - 2 || i == dt2.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "J"], worksheet.Cells[i + 6, "M"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "M"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "J"], worksheet.Cells[i + 6, "M"]).Font.Size = 10;
                        worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "M"]).Font.Name = "微软雅黑";
                        worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "M"]).NumberFormat = "￥0";

                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "M"], worksheet.Cells[i + 6, "M"]).NumberFormat = "0";
                        worksheet.get_Range(worksheet.Cells[i + 6, "K"], worksheet.Cells[i + 6, "K"]).NumberFormat = "0.00";
                    }
               
                    if (i == dt2.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "L"], worksheet.Cells[i + 6, "M"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "H"], worksheet.Cells[6 + i, "M"]).Borders.LineStyle = 1;
                }
            }/*PAPER_CORE _END*/
            d3 = 0;
            dt3 = RETURN_PRINTING_TOTAL(dt);/*PRINTING _START*/
            if (dt3.Rows.Count > 0)
            {
                for (i = 0; i < dt3.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "N"] = dt3.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "O"] = dt3.Rows[i]["机器型号"].ToString();
                    worksheet.Cells[i + 6, "P"] = dt3.Rows[i]["几款"].ToString();
                    if (!string.IsNullOrEmpty(dt3.Rows[i]["小计"].ToString()))
                    {
                        d3 = decimal.Parse(dt3.Rows[i]["小计"].ToString());
                    }
                    worksheet.Cells[i + 6, "Q"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0.00");
                    worksheet.get_Range(worksheet.Cells[i + 6, "N"], worksheet.Cells[i + 6, "N"]).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.Constants.xlRight;
                 
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).NumberFormat = "0.00";
                    worksheet.get_Range(worksheet.Cells[i + 6, "N"], worksheet.Cells[i + 6, "Q"]).Font.Size = 10;
                    worksheet.get_Range(worksheet.Cells[i + 6, "N"], worksheet.Cells[i + 6, "Q"]).Font.Name = "微软雅黑";
                    if (i == dt3.Rows.Count - 2 || i == dt3.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "O"], worksheet.Cells[i + 6, "Q"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "Q"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "O"], worksheet.Cells[i + 6, "Q"]).Font.Size = 10;
                        worksheet.get_Range(worksheet.Cells[i + 6, "O"], worksheet.Cells[i + 6, "Q"]).Font.Name = "微软雅黑";
                        worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "Q"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "Q"]).NumberFormat = "0";
                    }
                    if ( i == dt3.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "P"], worksheet.Cells[i + 6, "Q"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "N"], worksheet.Cells[6 + i, "Q"]).Borders.LineStyle = 1;
                }
            }/*PRINTING _END*/
            d3 = 0;
            dt4 = RETURN_SURFACE_MACHINING(dt);/*SURFACE_MACHINING _START*/
            if (dt4.Rows.Count > 0)
            {
                for (i = 0; i < dt4.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "R"] = dt4.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "S"] = dt4.Rows[i]["表面加工"].ToString();
                    worksheet.Cells[i + 6, "T"] = dt4.Rows[i]["几款"].ToString();
                    if (!string.IsNullOrEmpty(dt4.Rows[i]["小计"].ToString()))
                    {
                        d3 = decimal.Parse(dt4.Rows[i]["小计"].ToString());
                    }
                    worksheet.Cells[i + 6, "U"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0.00");
                    worksheet.get_Range(worksheet.Cells[i + 6, "R"], worksheet.Cells[i + 6, "R"]).HorizontalAlignment =
                      Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).NumberFormat = "0.00";

                    if (i == dt4.Rows.Count - 2 || i == dt4.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "S"], worksheet.Cells[i + 6, "U"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "T"], worksheet.Cells[i + 6, "U"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "T"], worksheet.Cells[i + 6, "U"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "U"], worksheet.Cells[i + 6, "U"]).NumberFormat = "0";

                    }
                    if (i==dt4.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "S"], worksheet.Cells[i + 6, "U"]).Font.Bold = true;
                        //worksheet.get_Range(worksheet.Cells[i + 6, "T"], worksheet.Cells[i + 6, "U"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "T"], worksheet.Cells[i + 6, "U"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "R"], worksheet.Cells[6 + i, "U"]).Borders.LineStyle = 1;
                }
            }/*SURFACE_MACHINING _END*/
            d3 = 0;
            dt5 = RETURN_LAMINATING_PROCESS(dt);/*LAMINATING_PROCESS _START*/
            if (dt5.Rows.Count > 0)
            {
                for (i = 0; i < dt5.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "V"] = dt5.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "W"] = dt5.Rows[i]["裱纸工艺"].ToString();
                    worksheet.Cells[i + 6, "X"] = dt5.Rows[i]["几款"].ToString();
                    if (!string.IsNullOrEmpty(dt5.Rows[i]["小计"].ToString()))
                    {
                        d3 = decimal.Parse(dt5.Rows[i]["小计"].ToString());
                    }
                    worksheet.Cells[i + 6, "Y"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0.00");
                    worksheet.get_Range(worksheet.Cells[i + 6, "AD"], worksheet.Cells[i + 6, "AD"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                   
                    worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "V"]).HorizontalAlignment =
                                   Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    if (i == dt5.Rows.Count - 2 || i == dt5.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "W"], worksheet.Cells[i + 6, "Y"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "X"], worksheet.Cells[i + 6, "Y"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "W"], worksheet.Cells[i + 6, "Y"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "V"], worksheet.Cells[i + 6, "Y"]).NumberFormat = "0";

                    }
                    if (i == dt5.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "W"], worksheet.Cells[i + 6, "Y"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "V"], worksheet.Cells[6 + i, "Y"]).Borders.LineStyle = 1;
                }
            }/*LAMINATING_PROCESS _END*/
            d3 = 0;
            dt6 = RETURN_DIE_CUTTING(dt);/*DIE_CUTTING _START*/
            if (dt6.Rows.Count > 0)
            {
                for (i = 0; i < dt6.Rows.Count; i++)
                {
                    worksheet.Cells[i + 6, "Z"] = dt6.Rows[i]["序号"].ToString();/*dgv1-1/i+6*/
                    worksheet.Cells[i + 6, "AA"] = dt6.Rows[i]["机器型号"].ToString();
                    worksheet.Cells[i + 6, "AB"] = dt6.Rows[i]["几款"].ToString();
                    if (!string.IsNullOrEmpty(dt6.Rows[i]["小计"].ToString()))
                    {
                        d3 = decimal.Parse(dt6.Rows[i]["小计"].ToString());
                    }
                    worksheet.Cells[i + 6, "AC"] = (d3 * (1 + x1 / 100) * (1 + x2 / 100)).ToString("0.00");
                    worksheet.get_Range(worksheet.Cells[i + 6, "AH"], worksheet.Cells[i + 6, "AH"]).HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.Constants.xlRight;
                    worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "Z"]).HorizontalAlignment =
                     Microsoft.Office.Interop.Excel.Constants.xlCenter;

                    if (i == dt6.Rows.Count - 2 || i == dt6.Rows.Count - 1)
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "AA"], worksheet.Cells[i + 6, "AC"]).Font.Bold = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AB"], worksheet.Cells[i + 6, "AC"]).MergeCells = true;
                        worksheet.get_Range(worksheet.Cells[i + 6, "AB"], worksheet.Cells[i + 6, "AC"]).NumberFormat = "￥0";
                    }
                    else
                    {
                        worksheet.get_Range(worksheet.Cells[i + 6, "Z"], worksheet.Cells[i + 6, "AC"]).NumberFormat = "0";

                    }
                    if ( i == dt6.Rows.Count - 1)
                    {

                        worksheet.get_Range(worksheet.Cells[i + 6, "AB"], worksheet.Cells[i + 6, "AC"]).NumberFormat = "0.00";
                    }
                    worksheet.get_Range(worksheet.Cells[6 + i, "Z"], worksheet.Cells[6 + i, "AC"]).Borders.LineStyle = 1;
                }
            }/*DIE_CUTTING _END*/

            worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[65536, 256]).Columns.AutoFit();
        }
        #endregion
        #region ExcelPrint_OFFER_FOR_AE_1
        public void ExcelPrint_OFFER_FOR_AE_1(DataTable dt, string BillName, string Printpath)
        {
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);

            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;

            if (dt.Rows.Count > 0)
            {
                worksheet.Cells[5,"H"] = dt.Rows[0]["客户"].ToString();
                worksheet.Cells[9,"H"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
                worksheet.Cells[9,"Y"] = dt.Rows[0]["数量"].ToString();
                worksheet.Cells[13,"P"] = dt.Rows[0]["数量"].ToString();
                worksheet.Cells[10,"H"] = dt.Rows[0]["项目号"].ToString();
                worksheet.Cells[10,"Y"] = dt.Rows[0]["报价编号"].ToString();
                worksheet.Cells[6,"Y"] = dt.Rows[0]["报价"].ToString();
                //worksheet.Cells[36, "S"] = dt.Rows[0]["日期"].ToString();
              
            }
     
            decimal x1 = 0;
            OFFER_ID = dt.Rows[0]["报价编号"].ToString();
            dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.OFFER_ID='" + OFFER_ID + "'");
            dt1 = bc.getdt(cother_cost.sql);
            if (dtx.Rows.Count > 0)
            {
                dtx = cprint_cost_total.RETURN_DT(dtx);
                if (!string.IsNullOrEmpty(dtx.Rows[18]["元套"].ToString()))
                {
                    x1 = decimal.Parse(dtx.Rows[18]["元套"].ToString());
                }
           
            }

            worksheet.Cells[13, "S"] = x1;
        }
        #endregion
        #region ExcelPrint_OFFER_FOR_AE_2
        public void ExcelPrint_OFFER_FOR_AE_2(DataTable dt, string BillName, string Printpath)
        {
            PFID = dt.Rows[0]["报价ID"].ToString();
            SaveFileDialog sfdg = new SaveFileDialog();
            //sfdg.DefaultExt = @"D:\xls";
            sfdg.Filter = "Excel(*.xls)|*.xls";
            sfdg.RestoreDirectory = true;
            sfdg.FileName = Printpath;
            sfdg.CreatePrompt = true;
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            workbook = application.Workbooks._Open(sfdg.FileName, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing);

            worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            application.Visible = true;
            application.ExtendList = false;
            application.DisplayAlerts = false;
            application.AlertBeforeOverwriting = false;
        
            if (dt.Rows.Count > 0)
            {   
                worksheet.Cells[4, "F"] = dt.Rows[0]["客户"].ToString();
                worksheet.Cells[8, "F"] = dt.Rows[0]["项目名称"].ToString();/*dgv1-1/2*/
                worksheet.Cells[8, "P"] = dt.Rows[0]["数量"].ToString();
                worksheet.Cells[9, "F"] = dt.Rows[0]["项目号"].ToString();
                worksheet.Cells[9, "P"] = dt.Rows[0]["报价编号"].ToString();
                worksheet.Cells[5, "P"] = dt.Rows[0]["报价"].ToString();
                worksheet.Cells[36, "S"] = dt.Rows[0]["日期"].ToString();
            }
            decimal x1 = 0;
            decimal x2 = 0;
            decimal x3 = 0;
            decimal x4 = 0;
            dtx = bc.getdt(cprint_cost_total.sql + " WHERE C.PFID='" + PFID + "'");
            sqb = new StringBuilder();
            sqb.AppendFormat(cother_cost.sql);
            sqb.AppendFormat(" WHERE C.CNAME='{0}'", dt.Rows[0]["客户"].ToString());
            sqb.AppendFormat(" AND A.BRAND='{0}'", dt.Rows[0]["品牌"].ToString());
            dt1 = bc.getdt(sqb.ToString());
           
            if (dtx.Rows.Count > 0)
            {
                dtx = cprint_cost_total.RETURN_DT(dtx);
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dtx.Rows[2]["主件用量"].ToString(), '%')))//主件单价
                {
                    x1 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[2]["主件用量"].ToString(), '%'));
                }
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dtx.Rows[0]["主件用量"].ToString(), '%')))//主件用量
                {
                    x2 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[0]["主件用量"].ToString(), '%'));
                }
                if (!string.IsNullOrEmpty(bc.RETURN_UNTIL_CHAR(dtx.Rows[4]["主件用量"].ToString(), '%')))//辅材单价
                {
                    x3 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx.Rows[4]["主件用量"].ToString(), '%'));
                }
        
                worksheet.Cells[25, "P"] = dtx.Rows[6]["主件用量"].ToString();
                worksheet.Cells[26, "P"] = dtx.Rows[8]["主件用量"].ToString();
              
            }
            DataTable dtx2 = bc.GET_DT_TO_DV_TO_DT(dt1, "", "项目='代购管理'");
            if (dtx2.Rows.Count > 0)
            {
                worksheet.Cells[28, "P"] = dtx2.Rows[0]["客户比例"].ToString();
                if (!string.IsNullOrEmpty(dtx2.Rows[0]["客户比例"].ToString()))//代购管理比例
                {
                    x4 = decimal.Parse(bc.RETURN_UNTIL_CHAR(dtx2.Rows[0]["客户比例"].ToString(), '%')) / 100;
                }
          
            }
            if (dtx.Rows.Count > 0)
            {
                for (i = 0; i < 6; i++)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["元套"].ToString()))
                    {
                        worksheet.Cells[12 + i, "I"] = decimal.Parse(dtx.Rows[i]["元套"].ToString()) * (1 + x1 / 100) * (1 + x2 / 100);
                    }
                }
                for (i = 6; i < 12; i++)
                {
                    if (!string.IsNullOrEmpty(dtx.Rows[i]["元套"].ToString()))
                    {
                        worksheet.Cells[12 + i, "I"] = decimal.Parse(dtx.Rows[i]["元套"].ToString()) * (1 + x3 / 100);
                    }
                }
           
        
                if (dtx.Rows[15]["元套"].ToString() == "" && dtx.Rows[16]["元套"].ToString() == "")
                {
                }
                else
                {
                    sqb = new StringBuilder();
                    sqb.AppendFormat("代购件={0}, ", dtx.Rows[15]["元套"].ToString());
                    sqb.AppendFormat("代购管理={0}, ", dtx.Rows[16]["元套"].ToString());
                    //MessageBox.Show(sqb.ToString());
                    worksheet.Cells[27, "I"] = (decimal.Parse(dtx.Rows[15]["元套"].ToString()) + decimal.Parse(dtx.Rows[16]["元套"].ToString())) /
                        (1 + x4);//代购件
                }
               //代购件管理费不用写，用EXCEL公式自动得到结果
             
             
            }
            worksheet.get_Range(worksheet.Cells[12, 5], worksheet.Cells[65536, 256]).Columns.AutoFit();
        }
        #endregion
        #region save_print_total
        public void save_print_total(DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            basec.getcoms("DELETE PRINT_TOTAL WHERE PFID='" + PFID   + "'");
            SQlcommandE_print_total(dt);
            IFExecution_SUCCESS = true;
        }
        #endregion
        #region SQlcommandE_print_total
        protected void SQlcommandE_print_total(DataTable dt)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            foreach (DataRow dr in dt.Rows)
            {
                SqlConnection sqlcon = bc.getcon();
                SqlCommand sqlcom = new SqlCommand(setsqlei, sqlcon);
                sqlcon.Open();
                sqlcom.Parameters.Add("PTID", SqlDbType.VarChar, 20).Value = bc.numYMD_NEW(12, 4, "0001", "print_total", "ptid", "PT");
                PFID = bc.getOnlyString("SELECT PFID FROM PRINTING_OFFER_MST WHERE OFFER_ID='" + dt.Rows [0]["报价编号"].ToString() + "'");
                sqlcom.Parameters.Add("PFID", SqlDbType.VarChar, 20).Value = PFID;
                sqlcom.Parameters.Add("C1", SqlDbType.VarChar, 20).Value = dr["序号"].ToString();
                sqlcom.Parameters.Add("C9", SqlDbType.VarChar, 20).Value = dr["加工门幅"].ToString();
                sqlcom.Parameters.Add("C10", SqlDbType.VarChar, 20).Value = dr["加工长度"].ToString();
                sqlcom.Parameters.Add("C11", SqlDbType.VarChar, 20).Value = dr["部品总数"].ToString();
                sqlcom.Parameters.Add("C12", SqlDbType.VarChar, 20).Value = dr["机器型号"].ToString();
                sqlcom.Parameters.Add("C13", SqlDbType.VarChar, 20).Value = dr["部品单价"].ToString();
                sqlcom.Parameters.Add("C14", SqlDbType.VarChar, 20).Value = dr["部品总价"].ToString();
                sqlcom.Parameters.Add("C15", SqlDbType.VarChar, 20).Value = dr["面纸单价"].ToString();
                sqlcom.Parameters.Add("C16", SqlDbType.VarChar, 20).Value = dr["面纸用量"].ToString();
                sqlcom.Parameters.Add("C17", SqlDbType.VarChar, 20).Value = dr["面纸内耗"].ToString();
                sqlcom.Parameters.Add("C18", SqlDbType.VarChar, 20).Value = dr["面纸下单"].ToString();
                sqlcom.Parameters.Add("C19", SqlDbType.VarChar, 20).Value = dr["面纸外耗"].ToString();
                sqlcom.Parameters.Add("C20", SqlDbType.VarChar, 20).Value = dr["面纸门幅"].ToString();
                sqlcom.Parameters.Add("C21", SqlDbType.VarChar, 20).Value = dr["面纸纸长"].ToString();
                sqlcom.Parameters.Add("C22", SqlDbType.VarChar, 20).Value = dr["面纸可用"].ToString();
                if (!string.IsNullOrEmpty(dr["面纸单个用量"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C23", SqlDbType.VarChar, 20).Value = dr["面纸单个用量"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C23", SqlDbType.VarChar, 20).Value = DBNull.Value;
             
                }
        
                if (!string.IsNullOrEmpty(dr["面纸小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C24", SqlDbType.VarChar, 20).Value = dr["面纸小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C24", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("C25", SqlDbType.VarChar, 20).Value = dr["芯纸单价"].ToString();
                sqlcom.Parameters.Add("C26", SqlDbType.VarChar, 20).Value = dr["芯纸内耗"].ToString();
                sqlcom.Parameters.Add("C27", SqlDbType.VarChar, 20).Value = dr["芯纸用量"].ToString();
                sqlcom.Parameters.Add("C28", SqlDbType.VarChar, 20).Value = dr["芯纸门幅"].ToString();
                sqlcom.Parameters.Add("C29", SqlDbType.VarChar, 20).Value = dr["芯纸纸长"].ToString();
                sqlcom.Parameters.Add("C30", SqlDbType.VarChar, 20).Value = dr["芯纸可用"].ToString();
                if (!string.IsNullOrEmpty(dr["芯纸单个用量"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C31", SqlDbType.VarChar, 20).Value = dr["芯纸单个用量"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C31", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
             
                if (!string.IsNullOrEmpty(dr["芯纸小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C32", SqlDbType.VarChar, 20).Value = dr["芯纸小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C32", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("C33", SqlDbType.VarChar, 20).Value = dr["底纸单价"].ToString();
                sqlcom.Parameters.Add("C34", SqlDbType.VarChar, 20).Value = dr["底纸用量"].ToString();
                sqlcom.Parameters.Add("C35", SqlDbType.VarChar, 20).Value = dr["底纸内耗"].ToString();
                sqlcom.Parameters.Add("C36", SqlDbType.VarChar, 20).Value = dr["底纸下单"].ToString();
                sqlcom.Parameters.Add("C37", SqlDbType.VarChar, 20).Value = dr["底纸外耗"].ToString();
                if (!string.IsNullOrEmpty(dr["底纸单个用量"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C38", SqlDbType.VarChar, 20).Value = dr["底纸单个用量"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C38", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
          
                if (!string.IsNullOrEmpty(dr["底纸小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C39", SqlDbType.VarChar, 20).Value = dr["底纸小计"].ToString();//此处需要DECIMAL类型，不然程序会报错
                }
                else
                {
                    sqlcom.Parameters.Add("C39", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("C40", SqlDbType.VarChar, 20).Value = dr["印工单色单价"].ToString();
                sqlcom.Parameters.Add("C41", SqlDbType.VarChar, 20).Value = dr["超出单色单张价"].ToString();
                sqlcom.Parameters.Add("C42", SqlDbType.VarChar, 20).Value = dr["CTP单张价"].ToString();
                sqlcom.Parameters.Add("C43", SqlDbType.VarChar, 20).Value = dr["正面色数共计"].ToString();
                sqlcom.Parameters.Add("C44", SqlDbType.VarChar, 20).Value = dr["正面CTP张数"].ToString();
                sqlcom.Parameters.Add("C45", SqlDbType.VarChar, 20).Value = dr["正面纸张损耗"].ToString();
                sqlcom.Parameters.Add("C46", SqlDbType.VarChar, 20).Value = dr["正面防晒合计"].ToString();
                sqlcom.Parameters.Add("C47", SqlDbType.VarChar, 20).Value = dr["正面CTP价计"].ToString();
                sqlcom.Parameters.Add("C48", SqlDbType.VarChar, 20).Value = dr["正面印工合计"].ToString();
                sqlcom.Parameters.Add("C49", SqlDbType.VarChar, 20).Value = dr["反面色数共计"].ToString();
                sqlcom.Parameters.Add("C50", SqlDbType.VarChar, 20).Value = dr["反面CTP张数"].ToString();
                sqlcom.Parameters.Add("C51", SqlDbType.VarChar, 20).Value = dr["反面纸张损耗"].ToString();
                sqlcom.Parameters.Add("C52", SqlDbType.VarChar, 20).Value = dr["反面防晒合计"].ToString();
                sqlcom.Parameters.Add("C53", SqlDbType.VarChar, 20).Value = dr["反面CTP价计"].ToString();
                sqlcom.Parameters.Add("C54", SqlDbType.VarChar, 20).Value = dr["反面印工合计"].ToString();
                if (!string.IsNullOrEmpty(dr["正反CTP合计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C55", SqlDbType.VarChar, 20).Value = dr["正反CTP合计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C55", SqlDbType.VarChar, 20).Value = DBNull.Value;//此处需要DECIMAL类型，不然程序会报错
                }
                if (!string.IsNullOrEmpty(dr["正反印工合计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C56", SqlDbType.VarChar, 20).Value = dr["正反印工合计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C56", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("C57", SqlDbType.VarChar, 20).Value = dr["表面处理单价"].ToString();
                sqlcom.Parameters.Add("C58", SqlDbType.VarChar, 20).Value = dr["无印刷表面处理损耗"].ToString();
                sqlcom.Parameters.Add("C59", SqlDbType.VarChar, 20).Value = dr["表面处理用量"].ToString();
                if (!string.IsNullOrEmpty(dr["表面加工小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C60", SqlDbType.VarChar, 20).Value = dr["表面加工小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C60", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
                sqlcom.Parameters.Add("C61", SqlDbType.VarChar, 20).Value = dr["裱工单价"].ToString();
                sqlcom.Parameters.Add("C62", SqlDbType.VarChar, 20).Value = dr["裱工用量"].ToString();
                if (!string.IsNullOrEmpty(dr["裱工小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C63", SqlDbType.VarChar, 20).Value = dr["裱工小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C63", SqlDbType.VarChar, 20).Value = DBNull.Value;

                }
                if (!string.IsNullOrEmpty(dr["刀模小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C64", SqlDbType.VarChar, 20).Value = dr["刀模小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C64", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
         
                if (!string.IsNullOrEmpty(dr["模切小计"].ToString())) //此为导出的xxx报价纸品估计计算表FOR_核价表等报表用到面纸小计SUM求和需要用DECIMAL类型
                {
                    sqlcom.Parameters.Add("C65", SqlDbType.VarChar, 20).Value = dr["模切小计"].ToString();
                }
                else
                {
                    sqlcom.Parameters.Add("C65", SqlDbType.VarChar, 20).Value = DBNull.Value;
                }
           
                sqlcom.Parameters.Add("MakerID", SqlDbType.VarChar, 20).Value = MAKERID;
                sqlcom.Parameters.Add("Date", SqlDbType.VarChar, 20).Value = varDate;
                sqlcom.Parameters.Add("YEAR", SqlDbType.VarChar, 20).Value = year;
                sqlcom.Parameters.Add("MONTH", SqlDbType.VarChar, 20).Value = month;
                sqlcom.Parameters.Add("DAY", SqlDbType.VarChar, 20).Value = day;
                sqlcom.ExecuteNonQuery();
                sqlcon.Close();
            }

        }
        #endregion
        
    } 
}
