using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace _8寸灯具参数指令解析
{
    public partial class Form1 : Form
    {
        //设置全局变量

        byte[] byteData;                      //声明字节数组，存储文件数据
        int Number = 0;                       //定义变量测试次数
        
        //Excel文件保存
        string str_fileName;                                                  //定义变量Excel文件名
        Microsoft.Office.Interop.Excel.Application ExcelApp;                  //声明Excel应用程序
        Workbook ExcelDoc;                                                    //声明工作簿
        Worksheet ExcelSheet;                                                 //声明工作表

        //8寸灯具各项参数存储集合
        ArrayList RMS1 = new ArrayList();
        ArrayList Val2 = new ArrayList();
        ArrayList Val3 = new ArrayList();
        ArrayList RMS = new ArrayList();
        ArrayList Current_Ratio1 = new ArrayList();
        ArrayList Current_Ratio3 = new ArrayList();
        ArrayList RES_IA = new ArrayList();
        ArrayList RES_IIA = new ArrayList();
        ArrayList SNS_IA = new ArrayList();
        ArrayList SNS_IIA = new ArrayList();
        ArrayList LED_F1 = new ArrayList();
        ArrayList T = new ArrayList();
        ArrayList Second = new ArrayList();


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                FileStream aFile = new FileStream(InputFileName.Text, FileMode.Open, FileAccess.Read);
                       
                try
                {
                    Number = Convert.ToInt32(TestNumber.Text);
                }
                catch
                {
                    MessageBox.Show("请输入阿拉伯数字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                byteData = new byte[96 * Number];
                aFile.Read(byteData, 0, 96*Number);

                DataAnalysis(byteData);

                EightInchesLampParametersCreatExcel();

            }
            catch
            {

            }
        }

        void DataAnalysis(byte[] byteDataAnalysis)
        {
            for(int i=0;i<Number;i++)
            {
                if(byteDataAnalysis[96*i]=='0'&&byteDataAnalysis[96*i+1]=='2')
                {
                    RMS1.Add(DataMakeUp(byteDataAnalysis[96 * i + 15], byteDataAnalysis[96 * i + 16]) * 1100);
                    Val2.Add(DataMakeUp(byteDataAnalysis[96 * i + 18], byteDataAnalysis[96 * i + 19]) * 20);
                    Val3.Add(DataMakeUp(byteDataAnalysis[96 * i + 21], byteDataAnalysis[96 * i + 22]));
                    RMS.Add(DataMakeUp(byteDataAnalysis[96 * i + 24], byteDataAnalysis[96 * i + 25]) * 4);
                    Current_Ratio1.Add(DataMakeUp(byteDataAnalysis[96 * i + 27], byteDataAnalysis[96 * i + 28])/10);
                    Current_Ratio3.Add(DataMakeUp(byteDataAnalysis[96 * i + 30], byteDataAnalysis[96 * i + 31])/10);
                    RES_IA.Add(DataMakeUp(byteDataAnalysis[96 * i + 33], byteDataAnalysis[96 * i + 34]) * 124);
                    RES_IIA.Add(DataMakeUp(byteDataAnalysis[96 * i + 36], byteDataAnalysis[96 * i + 37]) * 124);
                    SNS_IA.Add(DataMakeUp(byteDataAnalysis[96 * i + 39], byteDataAnalysis[96 * i + 40]) * 16);
                    SNS_IIA.Add(DataMakeUp(byteDataAnalysis[96 * i + 42], byteDataAnalysis[96 * i + 43]) * 16);
                    LED_F1.Add(DataMakeUp(byteDataAnalysis[96 * i + 45], byteDataAnalysis[96 * i + 46]));
                    T.Add(DataMakeUp(byteDataAnalysis[96 * i + 48], byteDataAnalysis[96 * i + 49]));

                    int SecondResult = 0;
                    for (int j=0;j<4;j++)
                    {
                        int SecondOrigin = (Int32)DataMakeUp(byteDataAnalysis[96 * i + (17 + j) * 3], byteDataAnalysis[96 * i + (17 + j) * 3 + 1]);                       
                        SecondResult |= SecondOrigin;
                        if(j<3)
                        {
                            SecondResult <<= 8;
                        }                                                                      
                    }
                    Second.Add(SecondResult);                   
                }
                else
                {
                    MessageBox.Show("指令错误！请确认导入文件数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        byte DataMakeUp(byte RawData1,byte RawData2)
        {
            byte DataAfter1 = DataTypeConversion(RawData1);
            byte DataAfter2 = DataTypeConversion(RawData2);
            
            DataAfter1 <<= 4;
            DataAfter1 |= DataAfter2;

            return DataAfter1;
        }

        byte DataTypeConversion(byte RawData0)
        {
            byte Result = 0x00;

            switch (RawData0)
            {
                case 0x30: Result = 0x00; break;
                case 0x31: Result = 0x01; break;
                case 0x32: Result = 0x02; break;
                case 0x33: Result = 0x03; break;
                case 0x34: Result = 0x04; break;
                case 0x35: Result = 0x05; break;
                case 0x36: Result = 0x06; break;
                case 0x37: Result = 0x07; break;
                case 0x38: Result = 0x08; break;
                case 0x39: Result = 0x09; break;
                case 0x41: Result = 0x0A; break;
                case 0x42: Result = 0x0B; break;
                case 0x43: Result = 0x0C; break;
                case 0x44: Result = 0x0D; break;
                case 0x45: Result = 0x0E; break;
                case 0x46: Result = 0x0F; break;               
            }

            return Result;

        }

            

        void EightInchesLampParametersCreatExcel()
        {
            //创建excel模板
            str_fileName = "d:\\8寸灯具参数解析" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
            ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
            ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

            //设置Excel列名           
            ExcelSheet.Cells[1, 1] = "8寸灯具参数解析";
            ExcelSheet.Cells[2, 1] = "序号";
            ExcelSheet.Cells[2, 2] = "RMS1";
            ExcelSheet.Cells[2, 3] = "Val2";
            ExcelSheet.Cells[2, 4] = "Val3";
            ExcelSheet.Cells[2, 5] = "RMS";
            ExcelSheet.Cells[2, 6] = "Current_Ratio1";
            ExcelSheet.Cells[2, 7] = "Current_Ratio3";
            ExcelSheet.Cells[2, 8] = "RES_IA";
            ExcelSheet.Cells[2, 9] = "RES_IIA";
            ExcelSheet.Cells[2, 10] = "SNS_IA";
            ExcelSheet.Cells[2, 11] = "SNS_IIA";
            ExcelSheet.Cells[2, 12] = "LED_F1";
            ExcelSheet.Cells[2, 13] = "T";
            ExcelSheet.Cells[2, 14] = "Second";

            for(int i=0;i<RMS1.Count;i++)
            {
                ExcelSheet.Cells[3+i, 1] = (i + 1).ToString();
                ExcelSheet.Cells[3+i, 2] = RMS1[i].ToString();
                ExcelSheet.Cells[3+i, 3] = Val2[i].ToString();
                ExcelSheet.Cells[3+i, 4] = Val3[i].ToString();
                ExcelSheet.Cells[3+i, 5] = RMS[i].ToString();
                ExcelSheet.Cells[3+i, 6] = Current_Ratio1[i].ToString();
                ExcelSheet.Cells[3+i, 7] = Current_Ratio3[i].ToString();
                ExcelSheet.Cells[3+i, 8] = RES_IA[i].ToString();
                ExcelSheet.Cells[3+i, 9] = RES_IIA[i].ToString();
                ExcelSheet.Cells[3+i, 10] = SNS_IA[i].ToString();
                ExcelSheet.Cells[3+i, 11] = SNS_IIA[i].ToString();
                ExcelSheet.Cells[3+i, 12] = LED_F1[i].ToString();
                ExcelSheet.Cells[3+i, 13] = T[i].ToString();
                ExcelSheet.Cells[3+i, 14] = Second[i].ToString();
            }

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序

        }


    }
}
