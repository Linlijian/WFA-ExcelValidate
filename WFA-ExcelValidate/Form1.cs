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
            AutoUpdater.Start("https://raw.githubusercontent.com/Linlijian/WFA-ExcelValidate/master/WFA-ExcelValidate/AutoUpdater.xml");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Excel Files(.xlsx)|*.xlsx";
            this.openFileDialog1.Title = "Select an excel file";

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                try
                {
                    //string path = @"D:\Users\x\Downloads\Compressed\DataGridView_Import_Excel\Sample Excel\validate_XML_life_quarterly_unlock_2-0-0.xlsx";
                    FileStream stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    DataSet result = excelReader.AsDataSet();

                    //string REPORT_CODE = "ช0200";
                    //string VALIDATE_X = "1.0.0.0.0,1.1.0.0.0,1.2.0.0.0,1.3.0.0.0,2.0.0.0.0";
                    //string VALIDATE_Y = "XB,XC,XD,XE,XF,XG,XH,XI,XJ";

                    var area = txtDiff.Text.Replace("\r\n", ",");

                    string REPORT_CODE = txtReportCode.Text;
                    string VALIDATE_X = txtValidateX.Text;
                    string VALIDATE_Y = txtValidateY.Text = area;



                    var people = new List<DemoModel>();
                    while (excelReader.Read())
                    {
                        people.Add(new DemoModel
                        {
                            COM_GROUP_ID = excelReader.GetString(1),
                            RPT_TYPE_ID = excelReader.GetString(2),
                            VALIDATE_ID = excelReader.GetString(3),
                            VALIDATE_CODE = excelReader.GetString(4),
                            REPORT_CODE = excelReader.GetString(5),
                            VALIDATE_X = excelReader.GetString(6),
                            VALIDATE_Y = excelReader.GetString(7),
                            DATA_TYPE = excelReader.GetString(8),
                            IS_NOTNULL = excelReader.GetString(9),
                            REGULAR_EXPRESSIONS = excelReader.GetString(10),
                            START_EXCEL_ROW = excelReader.GetString(11),
                            ACTIVE = excelReader.GetString(12),
                            REMARK = excelReader.GetString(13)
                        });
                    }

                    var report_model = people.Where(c => c.REPORT_CODE == REPORT_CODE).ToList<DemoModel>();
                    List<DemoModel> output_report = new List<DemoModel>();
                    var item_x = VALIDATE_X.Split(',');
                    var item_y = VALIDATE_Y.Split(',');
                    foreach (var valid_x in item_x)
                    {
                        var row_x = report_model.Where(c => c.VALIDATE_X == valid_x).ToList<DemoModel>();
                        if (row_x.Count() > 0)
                        {
                            foreach (var valid_y in item_y)
                            {
                                var row_y = row_x.Where(c => c.VALIDATE_X == valid_x && c.VALIDATE_Y == valid_y).ToList();
                                if (row_y.Count() > 1)
                                {
                                    bool first = true;
                                    foreach (var update in row_y)
                                    {
                                        if (first)
                                        {
                                            update.ACTIVE = "Y";
                                            output_report.Add(update);
                                            first = false;
                                        }
                                        else
                                        {
                                            update.ACTIVE = "N";
                                            update.REMARK = "DUP";
                                            output_report.Add(update);
                                        }
                                    }
                                }
                                else if (row_y.Count() == 1)
                                {
                                    row_y[0].ACTIVE = "Y";
                                    output_report.Add(row_y[0]);
                                }
                                else if (row_y.Count() == 0)
                                {
                                    row_y.Add(new DemoModel
                                    {
                                        COM_GROUP_ID = "",
                                        RPT_TYPE_ID = "",
                                        VALIDATE_ID = "",
                                        REPORT_CODE = REPORT_CODE,
                                        VALIDATE_X = valid_x,
                                        VALIDATE_Y = valid_y,
                                        DATA_TYPE = "",
                                        IS_NOTNULL = "",
                                        REGULAR_EXPRESSIONS = "",
                                        START_EXCEL_ROW = "",
                                        ACTIVE = "N",
                                        REMARK = "MISS 'VALIDATE_Y' IN TEMPLATE"
                                    });
                                    output_report.Add(row_y[0]);
                                }
                            }
                        }
                        else
                        {

                            foreach (var valid_y in item_y)
                            {
                                List<DemoModel> output = new List<DemoModel>();
                                //old
                                output.Add(new DemoModel
                                {
                                    COM_GROUP_ID = "",
                                    RPT_TYPE_ID = "",
                                    VALIDATE_ID = "",
                                    REPORT_CODE = REPORT_CODE,
                                    VALIDATE_X = valid_x,
                                    VALIDATE_Y = valid_y,
                                    DATA_TYPE = "",
                                    IS_NOTNULL = "",
                                    REGULAR_EXPRESSIONS = "",
                                    START_EXCEL_ROW = "",
                                    ACTIVE = "N",
                                    REMARK = "MISS 'VALIDATE_X' IN TEMPLATE"
                                });
                                output_report.Add(output[0]);

                                //new
                                output = new List<DemoModel>();
                                output.Add(new DemoModel
                                {
                                    COM_GROUP_ID = "",
                                    RPT_TYPE_ID = "",
                                    VALIDATE_ID = "",
                                    REPORT_CODE = REPORT_CODE,
                                    VALIDATE_X = "INPUTE VALIDATE_X " + valid_x,
                                    VALIDATE_Y = valid_y,
                                    DATA_TYPE = "",
                                    IS_NOTNULL = "",
                                    REGULAR_EXPRESSIONS = "",
                                    START_EXCEL_ROW = "",
                                    ACTIVE = "N",
                                    REMARK = "MISS 'VALIDATE_X' IN TEMPLATE"
                                });
                                output_report.Add(output[0]);
                            }
                        }

                    }

                    this.resultGrid.DataSource = output_report;
                }
                catch { MessageBox.Show("KEY!"); }               
            }
        }
    }
}
