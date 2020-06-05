using AutoUpdaterDotNET;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WFA_ExcelValidate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            button2.Enabled = true;
            button3.Enabled = true;
            AutoUpdater.Start("https://raw.githubusercontent.com/Linlijian/WFA-ExcelValidate/master/WFA-ExcelValidate/AutoUpdater.xml");
        }

        string path;
        string path2;
        List<crr_t_mst_insurance> CRR_T_MST_INSURANCE_MODEL = new List<crr_t_mst_insurance>();
        List<CRR_T_SEC_USRGROUP> CRR_T_SEC_USRGROUP_MODEL = new List<CRR_T_SEC_USRGROUP>();
        private void button1_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Excel Files(.xlsx)|*.xlsx";
            this.openFileDialog1.Title = "Select an excel file";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileStream stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelReader.AsDataSet();
                path = openFileDialog1.FileName;


                while (excelReader.Read())
                {
                    CRR_T_MST_INSURANCE_MODEL.Add(new crr_t_mst_insurance
                    {
                        MS_INS_ID = excelReader.GetString(1),
                        MS_INS_CODE = excelReader.GetString(2),
                        MS_INS_NAME_THAI = excelReader.GetString(3),
                        MS_INS_NAME_ENG = excelReader.GetString(4),
                        MS_INS_INITIALS = excelReader.GetString(5),
                        MS_INS_COM_TYPE = excelReader.GetString(6),
                        MS_INS_CER_DIR = excelReader.GetString(7),
                        MS_INS_CER_FILE = excelReader.GetString(8),
                        MS_INS_CONTRACT_TO = excelReader.GetString(9),
                        MS_INS_REMARK = excelReader.GetString(10),
                        MS_INS_COMPANY_BRAND = excelReader.GetString(11),
                        STATUS = excelReader.GetString(12),
                        ACTIVE = excelReader.GetString(13),
                        CREATED_USER = excelReader.GetString(14),
                        CREATED_DT = excelReader.GetString(15),
                        UPDATE_USER = excelReader.GetString(16),
                        UPDATE_DT = excelReader.GetString(17),
                        MS_INS_COMPANY_LOGO = excelReader.GetString(18),
                        MS_INS_CONTRACT_COM = excelReader.GetString(19),
                        MS_INS_TEL_COM = excelReader.GetString(20),
                        MS_INS_LICENSE_NO = excelReader.GetString(21),
                        MS_INS_MAILSERVER = excelReader.GetString(22),
                        MS_INS_ADDR1 = excelReader.GetString(23),
                        MS_INS_ADDR2 = excelReader.GetString(24),
                        MS_ESTABLISHMENT = excelReader.GetString(25),
                        MS_ORIGIN = excelReader.GetString(26)
                    });
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Excel Files(.xlsx)|*.xlsx";
            this.openFileDialog1.Title = "Select an excel file";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileStream stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelReader.AsDataSet();
                path = openFileDialog1.FileName;


                while (excelReader.Read())
                {
                    CRR_T_SEC_USRGROUP_MODEL.Add(new CRR_T_SEC_USRGROUP
                    {
                        COM_CODE = excelReader.GetString(1),
                        USG_TYPE = excelReader.GetString(2),
                        MS_INS_ID = excelReader.GetString(3),
                        USG_ID = excelReader.GetString(4),
                        USG_CODE = excelReader.GetString(5),
                        USG_NAME_TH = excelReader.GetString(6),
                        USG_NAME_EN = excelReader.GetString(7),
                        USG_STATUS = excelReader.GetString(8),
                        USG_LEVEL = excelReader.GetString(9),
                        SID = excelReader.GetString(10),
                        CREATED_USER = excelReader.GetString(11),
                        CREATED_DT = excelReader.GetString(12),
                        UPDATE_USER = excelReader.GetString(13),
                        UPDATE_DT = excelReader.GetString(14),
                        IS_DUMMY_DEPT = excelReader.GetString(15),
                        STATUS = excelReader.GetString(16),
                        ACTIVE = excelReader.GetString(17),
                        WARNING_DT = excelReader.GetString(18),
                        EXPIRY_DT = excelReader.GetString(19),
                        USG_INFO = excelReader.GetString(20),
                        EFFECTIVE_DT = excelReader.GetString(21),
                        MS_INS_COM_TYPE = excelReader.GetString(22)

                    });
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            string ussaasasg = (from m in CRR_T_SEC_USRGROUP_MODEL
                                where (m.USG_STATUS == "Y") && (m.USG_TYPE == "2") && (m.MS_INS_ID == CRR_T_MST_INSURANCE_MODEL[2].MS_INS_ID)
                                select m.USG_ID).ToString();
            string a = ussaasasg;

            string sss = (from m in CRR_T_SEC_USRGROUP_MODEL
                          where (m.MS_INS_ID == CRR_T_MST_INSURANCE_MODEL[2].MS_INS_ID)
                          select m.USG_ID).ToString();
            string asss = sss;

            var popopo = CRR_T_MST_INSURANCE_MODEL.Where(w => w.MS_INS_ID == "139").ToList();
            List<string> insert = new List<string>();
            List<string> USERAPP = new List<string>();
            try
            {
                foreach (crr_t_mst_insurance model in CRR_T_MST_INSURANCE_MODEL)
                {
                    var ugs2 = CRR_T_SEC_USRGROUP_MODEL.Where(w => w.MS_INS_ID == model.MS_INS_ID).ToList();
                    if (ugs2 != null)
                    {
                        string gensql = "INSERT INTO CRR_T_SEC_USER(";
                        gensql += "COM_CODE,";
                        gensql += "USER_TYPE,";
                        gensql += "MS_INS_ID,";
                        gensql += "USER_ID,";
                        gensql += "ID_CARD_NO,";
                        gensql += "USER_FNAME_TH,";
                        gensql += "USER_LNAME_TH,";
                        gensql += "USER_FNAME_EN,";
                        gensql += "USER_LNAME_EN,";
                        gensql += "POSITION,";
                        gensql += "SID,";
                        gensql += "USG_ID,";
                        gensql += "USER_SPEC_ID,";
                        gensql += "USER_PWD,";
                        gensql += "USER_EFF_DATE,";
                        gensql += "USER_EXP_DATE,";
                        gensql += "PWD_EXP_DATE,";
                        gensql += "WNING_USER_DATE,";
                        gensql += "WNING_PWD_DATE,";
                        gensql += "END_ACT_DATE,";
                        gensql += "TELEPHONE,";
                        gensql += "FAX,";
                        gensql += "EMAIL,";
                        gensql += "REASON,";
                        gensql += "REMARK,";
                        gensql += "PROXY_FILE,";
                        gensql += "USER_STATUS,";
                        gensql += "IS_FCP,";
                        gensql += "IS_NCE,";
                        gensql += "IS_DISABLED,";
                        gensql += "LAST_LOGIN_DATE,";
                        gensql += "CREATED_USER,";
                        gensql += "CREATED_DT,";
                        gensql += "UPDATE_USER,";
                        gensql += "UPDATE_DT,";
                        gensql += "CARD_REF_TYPE,";
                        gensql += "IS_APPROVE,";
                        gensql += "APPROVED_DT,";
                        gensql += "EMPLOYEE_ID,";
                        gensql += "ROLE_LI_1,";
                        gensql += "ROLE_LI_2,";
                        gensql += "ROLE_LI_3,";
                        gensql += "ROLE_LI_4,";
                        gensql += "ROLE_LI_5,";
                        gensql += "ROLE_LI_6,";
                        gensql += "ROLE_LI_7,";
                        gensql += "ROLE_LI_8,";
                        gensql += "ROLE_NL_1,";
                        gensql += "ROLE_NL_2,";
                        gensql += "ROLE_NL_3,";
                        gensql += "ROLE_NL_4,";
                        gensql += "ROLE_NL_5,";
                        gensql += "ROLE_NL_6,";
                        gensql += "ROLE_NL_71,";
                        gensql += "ROLE_NL_72,";
                        gensql += "ROLE_NL_73,";
                        gensql += "ROLE_NL_81,";
                        gensql += "ROLE_NL_82,";
                        gensql += "ROLE_NL_9,";
                        gensql += "APPROVE_STATUS,";
                        gensql += "APPROVE_DT,";
                        gensql += "APRROVE_BY,";
                        gensql += "REASON_TYPE_ID,";
                        gensql += "REASON_REJECT,";
                        gensql += "LAST_LOCK_DATE,";
                        gensql += "MS_INS_COM_TYPE,";
                        gensql += "CANCLE_STATUS,";
                        gensql += "CANCLE_DATE,";
                        gensql += "CANCEL_NOTE,";
                        gensql += "CANCEL_BY,";
                        gensql += "TITLE_ID,";
                        gensql += "ACTIVE,";
                        gensql += "STATUS)VALUES(";

                        gensql += "'OIC' " + ',';
                        gensql += "2" + ',';
                        gensql += "" + "\'" + model.MS_INS_ID + "\'" + ',';
                        gensql += "'TEST_" + model.MS_INS_NAME_ENG.Split(' ').First() + "@" + model.MS_INS_MAILSERVER + "'" + ',';
                        gensql += "'P1uzQMcRvsKgpmp0f8Vaaw=='" + ',';
                        gensql += "'" + model.MS_INS_NAME_ENG + "'" + ',';
                        gensql += "'" + model.MS_INS_NAME_ENG + "'" + ',';
                        gensql += "'" + model.MS_INS_NAME_ENG + "'" + ',';
                        gensql += "'" + model.MS_INS_NAME_ENG + "'" + ',';
                        gensql += "'TEST'" + ',';
                        gensql += "''" + ',';


                        txtValidateX.Text += model.MS_INS_ID + "__";

                        gensql += "'" + ugs2[0].USG_ID + "'" + ',';

                        gensql += "''" + ',';
                        gensql += "'9MhKmP3caWjI2qzFPdEIrw==" + "'" + ',';
                        gensql += "TO_DATE('2020/04/01', 'yyyy/mm/dd')" + ',';
                        gensql += "TO_DATE('2020/09/01', 'yyyy/mm/dd')" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "'999999999" + "'" + ',';
                        gensql += "''" + ',';
                        gensql += "'TEST_" + model.MS_INS_NAME_ENG.Split(' ').First() + "@" + model.MS_INS_MAILSERVER + "'" + ',';
                        gensql += "'20022007426'" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "'Y'" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "'N'" + ',';
                        gensql += "''" + ',';
                        gensql += "'TEST_" + model.MS_INS_NAME_ENG.Split(' ').First() + "@" + model.MS_INS_MAILSERVER + "'" + ',';
                        gensql += "''" + ',';
                        gensql += "'vsmadmin'" + ',';
                        gensql += "TO_DATE('2020/04/01 8:30:25', 'YYYY/MM/DD HH:MI:SS')" + ',';
                        gensql += "'2'" + ',';
                        gensql += "'Y'" + ',';
                        gensql += "TO_DATE('2020/04/01 8:30:25', 'YYYY/MM/DD HH:MI:SS')" + ',';
                        gensql += "''" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'N'" + ',';
                        gensql += "'50040003'" + ',';
                        gensql += "TO_DATE('2020/04/01 8:30:25', 'YYYY/MM/DD HH:MI:SS')" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "'" + model.MS_INS_COM_TYPE + "'" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "''" + ',';
                        gensql += "'102'" + ',';
                        gensql += "''" + ',';
                        gensql += "'');";

                        insert.Add(gensql);

                        string sqlUSERAPP = "INSERT INTO CRR_T_SEC_USERAPP(";
                        sqlUSERAPP += "USER_ID,";
                        sqlUSERAPP += "USG_ID,";
                        sqlUSERAPP += "MS_INS_COM_TYPE,";
                        sqlUSERAPP += "COM_CODE,";
                        sqlUSERAPP += "CREATED_USER,";
                        sqlUSERAPP += "CREATED_DT,";
                        sqlUSERAPP += "UPDATE_USER,";
                        sqlUSERAPP += "UPDATE_DT)VALUES(";
                        sqlUSERAPP += "'TEST_" + model.MS_INS_NAME_ENG.Split(' ').First() + "@" + model.MS_INS_MAILSERVER + "'" + ',';
                        sqlUSERAPP += "'" + ugs2[0].USG_ID + "'" + ',';
                        sqlUSERAPP += "'" + model.MS_INS_COM_TYPE + "'" + ',';
                        sqlUSERAPP += "'OIC'" + ',';
                        sqlUSERAPP += "'TEST_" + model.MS_INS_NAME_ENG.Split(' ').First() + "@" + model.MS_INS_MAILSERVER + "'" + ',';
                        sqlUSERAPP += "TO_DATE('2020/04/01 8:30:25', 'YYYY/MM/DD HH:MI:SS')" + ',';
                        sqlUSERAPP += "'vsmadmin'" + ',';
                        sqlUSERAPP += "TO_DATE('2020/04/01 8:30:25', 'YYYY/MM/DD HH:MI:SS'))" + ';';

                        USERAPP.Add(sqlUSERAPP);
                    }

                }
            }
            catch { }

            foreach (string aaa in insert)
            {
                txtReportCode.Text += aaa + "                   ";
            }
            foreach (string aaaa in USERAPP)
            {
                txtValidateY.Text += aaaa + "                   ";
            }

        }
    }
}



//string sql = @"INSERT INTO CRR_T_SEC_USER 
//                                    (COM_CODE,
//                                    USER_TYPE,
//                                    MS_INS_ID,
//                                    USER_ID,
//                                    ID_CARD_NO,
//                                    USER_FNAME_TH,
//                                    USER_LNAME_TH,
//                                    USER_FNAME_EN,
//                                    USER_LNAME_EN,
//                                    POSITION,
//                                    SID,
//                                    USG_ID,
//                                    USER_SPEC_ID,
//                                    USER_PWD,
//                                    USER_EFF_DATE,
//                                    USER_EXP_DATE,
//                                    PWD_EXP_DATE,
//                                    WNING_USER_DATE,
//                                    WNING_PWD_DATE,
//                                    END_ACT_DATE,
//                                    TELEPHONE,
//                                    FAX,
//                                    EMAIL,
//                                    REASON,
//                                    REMARK,
//                                    PROXY_FILE,
//                                    USER_STATUS,
//                                    IS_FCP,
//                                    IS_NCE,
//                                    IS_DISABLED,
//                                    LAST_LOGIN_DATE,
//                                    CREATED_USER,
//                                    CREATED_DT,
//                                    UPDATE_USER,
//                                    UPDATE_DT,
//                                    CARD_REF_TYPE,
//                                    IS_APPROVE,
//                                    APPROVED_DT,
//                                    EMPLOYEE_ID,
//                                    ROLE_LI_1,
//                                    ROLE_LI_2,
//                                    ROLE_LI_3,
//                                    ROLE_LI_4,
//                                    ROLE_LI_5,
//                                    ROLE_LI_6,
//                                    ROLE_LI_7,
//                                    ROLE_LI_8,
//                                    ROLE_NL_1,
//                                    ROLE_NL_2,
//                                    ROLE_NL_3,
//                                    ROLE_NL_4,
//                                    ROLE_NL_5,
//                                    ROLE_NL_6,
//                                    ROLE_NL_71,
//                                    ROLE_NL_72,
//                                    ROLE_NL_73,
//                                    ROLE_NL_81,
//                                    ROLE_NL_82,
//                                    ROLE_NL_9,
//                                    APPROVE_STATUS,
//                                    APPROVE_DT,
//                                    APRROVE_BY,
//                                    REASON_TYPE_ID,
//                                    REASON_REJECT,
//                                    LAST_LOCK_DATE,
//                                    MS_INS_COM_TYPE,
//                                    CANCLE_STATUS,
//                                    CANCLE_DATE,
//                                    CANCEL_NOTE,
//                                    CANCEL_BY,
//                                    TITLE_ID,
//                                    ACTIVE,
//                                    STATUS
//                                    )
//                                    VALUES 
//                                        (COM_CODE={1},
//USER_TYPE={2},
//MS_INS_ID={3},
//USER_ID={4},
//ID_CARD_NO={5},
//USER_FNAME_TH={6},
//USER_LNAME_TH={7},
//USER_FNAME_EN={8},
//USER_LNAME_EN={9},
//POSITION={10},
//SID={11},
//USG_ID={12},
//USER_SPEC_ID={13},
//USER_PWD={14},
//USER_EFF_DATE={15},
//USER_EXP_DATE={16},
//PWD_EXP_DATE={17},
//WNING_USER_DATE={18},
//WNING_PWD_DATE={19},
//END_ACT_DATE={20},
//TELEPHONE={21},
//FAX={22},
//EMAIL={23},
//REASON={24},
//REMARK={25},
//PROXY_FILE={26},
//USER_STATUS={27},
//IS_FCP={28},
//IS_NCE={29},
//IS_DISABLED={30},
//LAST_LOGIN_DATE={31},
//CREATED_USER={32},
//CREATED_DT={33},
//UPDATE_USER={34},
//UPDATE_DT={35},
//CARD_REF_TYPE={36},
//IS_APPROVE={37},
//APPROVED_DT={38},
//EMPLOYEE_ID={39},
//ROLE_LI_1={40},
//ROLE_LI_2={41},
//ROLE_LI_3={42},
//ROLE_LI_4={43},
//ROLE_LI_5={44},
//ROLE_LI_6={45},
//ROLE_LI_7={46},
//ROLE_LI_8={47},
//ROLE_NL_1={48},
//ROLE_NL_2={49},
//ROLE_NL_3={50},
//ROLE_NL_4={51},
//ROLE_NL_5={52},
//ROLE_NL_6={53},
//ROLE_NL_71={54},
//ROLE_NL_72={55},
//ROLE_NL_73={56},
//ROLE_NL_81={57},
//ROLE_NL_82={58},
//ROLE_NL_9={59},
//APPROVE_STATUS={60},
//APPROVE_DT={61},
//APRROVE_BY={62},
//REASON_TYPE_ID={63},
//REASON_REJECT={64},
//LAST_LOCK_DATE={65},
//MS_INS_COM_TYPE={66},
//CANCLE_STATUS={67},
//CANCLE_DATE={68},
//CANCEL_NOTE={69},
//CANCEL_BY={70},
//TITLE_ID={71},
//ACTIVE={72},
//STATUS={73})";