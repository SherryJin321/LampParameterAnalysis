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
        ArrayList InputData = new ArrayList();     //定义集合，存储整个文件数据
        int intCode = 0;                           //定义变量，存储一次读取文件返回的数据
        
        //Excel文件保存
        string str_fileName;                                                  //定义变量Excel文件名
        Microsoft.Office.Interop.Excel.Application ExcelApp;                  //声明Excel应用程序
        Workbook ExcelDoc;                                                    //声明工作簿
        Worksheet ExcelSheet;                                                 //声明工作表

        //8寸灯具各项参数存储集合
        ArrayList RMS1_8inches = new ArrayList();
        ArrayList Val2_8inches = new ArrayList();
        ArrayList Val3_8inches = new ArrayList();
        ArrayList RMS_8inches = new ArrayList();
        ArrayList Current_Ratio1_8inches = new ArrayList();
        ArrayList Current_Ratio3_8inches = new ArrayList();
        ArrayList RES_IA_8inches = new ArrayList();
        ArrayList RES_IIA_8inches = new ArrayList();
        ArrayList SNS_IA_8inches = new ArrayList();
        ArrayList SNS_IIA_8inches = new ArrayList();
        ArrayList LED_F1_8inches = new ArrayList();
        ArrayList T_8inches = new ArrayList();
        ArrayList Second_8inches = new ArrayList();

        //12寸灯具各项参数存储集合
        ArrayList RMS1_12inches = new ArrayList();
        ArrayList RMS2_12inches = new ArrayList();
        ArrayList Val2_12inches = new ArrayList();
        ArrayList Val3_12inches = new ArrayList();
        ArrayList RMSMID1_12inches = new ArrayList();
        ArrayList RMSMID2_12inches = new ArrayList();
        ArrayList RMS_12inches = new ArrayList();
        ArrayList Current_Ratio1_12inches = new ArrayList();
        ArrayList Current_Ratio2_12inches = new ArrayList();
        ArrayList Current_Ratio3_12inches = new ArrayList();
        ArrayList Current_Ratio4_12inches = new ArrayList();
        ArrayList RES_IA_12inches = new ArrayList();
        ArrayList RES_IB_12inches = new ArrayList();
        ArrayList RES_IIA_12inches = new ArrayList();
        ArrayList RES_IIB_12inches = new ArrayList();
        ArrayList SNS_IA_12inches = new ArrayList();
        ArrayList SNS_IB_12inches = new ArrayList();
        ArrayList SNS_IIA_12inches = new ArrayList();
        ArrayList SNS_IIB_12inches = new ArrayList();
        ArrayList LED_F1_12inches = new ArrayList();
        ArrayList T_12inches = new ArrayList();
        ArrayList Second_12inches = new ArrayList();


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                FileStream aFile = new FileStream(InputFileName.Text, FileMode.Open);
                StreamReader sr = new StreamReader(aFile);

                //读取整个文件数据
                intCode = sr.Read();
                while(intCode!=-1)
                {
                    InputData.Add(intCode);
                    intCode = sr.Read();
                }
                sr.Close();
                               
                if(radioButton1.Checked==true)
                {
                    EightInchesDataAnalysis(InputData);                
                    EightInchesLampParametersCreatExcel();
                }
                else
                {                                        
                    TwelveInchesDataAnalysis(InputData);
                    TwelveInchesLampParametersCreatExcel();
                }
            }
            catch
            {
                MessageBox.Show("未找到文件！请确认文件路径是否正确输入", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
}

        /// <summary>
        /// 函数功能：8寸灯具，解析指令，将各个参数值保存在各自的集合中
        /// </summary>
        /// <param name="byteDataAnalysis"></param>
        void EightInchesDataAnalysis(ArrayList byteDataAnalysis)
        {
            for (int i = 0; i < byteDataAnalysis.Count/96; i++)
            {
                if ((int)byteDataAnalysis[96 * i] == 48 && (int)byteDataAnalysis[96 * i + 1] == 50)
                {
                    RMS1_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 15], (int)byteDataAnalysis[96 * i + 16]) * 1100);
                    Val2_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 18], (int)byteDataAnalysis[96 * i + 19]) * 20);
                    Val3_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 21], (int)byteDataAnalysis[96 * i + 22]));
                    RMS_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 24], (int)byteDataAnalysis[96 * i + 25]) * 4);
                    Current_Ratio1_8inches.Add((float)(DataMakeUp((int)byteDataAnalysis[96 * i + 27], (int)byteDataAnalysis[96 * i + 28]) / 10.0));
                    Current_Ratio3_8inches.Add((float)(DataMakeUp((int)byteDataAnalysis[96 * i + 30], (int)byteDataAnalysis[96 * i + 31]) / 10.0));
                    RES_IA_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 33], (int)byteDataAnalysis[96 * i + 34]) * 124);
                    RES_IIA_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 36], (int)byteDataAnalysis[96 * i + 37]) * 124);
                    SNS_IA_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 39], (int)byteDataAnalysis[96 * i + 40]) * 16);
                    SNS_IIA_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 42], (int)byteDataAnalysis[96 * i + 43]) * 16);
                    LED_F1_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 45], (int)byteDataAnalysis[96 * i + 46]));
                    T_8inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 48], (int)byteDataAnalysis[96 * i + 49]));

                    int SecondResult = 0;
                    for (int j = 0; j < 4; j++)
                    {
                        int SecondOrigin = (Int32)DataMakeUp((int)byteDataAnalysis[96 * i + (17 + j) * 3], (int)byteDataAnalysis[96 * i + (17 + j) * 3 + 1]);
                        SecondResult |= SecondOrigin;
                        if (j < 3)
                        {
                            SecondResult <<= 8;
                        }
                    }
                    Second_8inches.Add(SecondResult);
                }
                else
                {
                    MessageBox.Show("文件格式错误！请确认文件内容", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        /// <summary>
        /// 函数功能：12寸灯具，解析指令，将各个参数值保存在各自的集合中
        /// </summary>
        /// <param name="byteDataAnalysis"></param>
        void TwelveInchesDataAnalysis(ArrayList byteDataAnalysis)
        {
            for (int i = 0; i < byteDataAnalysis.Count / 96; i++)
            {
                if ((int)byteDataAnalysis[96 * i] == 48 && (int)byteDataAnalysis[96 * i + 1] == 50)
                {
                    RMS1_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 15], (int)byteDataAnalysis[96 * i + 16]) * 500);
                    RMS2_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 18], (int)byteDataAnalysis[96 * i + 19]) * 500);
                    Val2_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 21], (int)byteDataAnalysis[96 * i + 22]) * 20);
                    Val3_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 24], (int)byteDataAnalysis[96 * i + 25]));
                    RMSMID1_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 27], (int)byteDataAnalysis[96 * i + 28]) * 16);
                    RMSMID2_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 30], (int)byteDataAnalysis[96 * i + 31]) * 16);
                    RMS_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 33], (int)byteDataAnalysis[96 * i + 34]) * 4);
                    Current_Ratio1_12inches.Add((float)(DataMakeUp((int)byteDataAnalysis[96 * i + 36], (int)byteDataAnalysis[96 * i + 37]) / 10.0));
                    Current_Ratio2_12inches.Add((float)(DataMakeUp((int)byteDataAnalysis[96 * i + 39], (int)byteDataAnalysis[96 * i + 40]) / 10.0));
                    Current_Ratio3_12inches.Add((float)(DataMakeUp((int)byteDataAnalysis[96 * i + 42], (int)byteDataAnalysis[96 * i + 43]) / 10.0));
                    Current_Ratio4_12inches.Add((float)(DataMakeUp((int)byteDataAnalysis[96 * i + 45], (int)byteDataAnalysis[96 * i + 46]) / 10.0));
                    RES_IA_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 48], (int)byteDataAnalysis[96 * i + 49]) * 124);
                    RES_IB_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 51], (int)byteDataAnalysis[96 * i + 52]) * 124);
                    RES_IIA_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 54], (int)byteDataAnalysis[96 * i + 55]) * 124);
                    RES_IIB_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 57], (int)byteDataAnalysis[96 * i + 58]) * 124);
                    SNS_IA_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 60], (int)byteDataAnalysis[96 * i + 61]) * 16);
                    SNS_IB_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 63], (int)byteDataAnalysis[96 * i + 64]) * 16);
                    SNS_IIA_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 66], (int)byteDataAnalysis[96 * i + 67]) * 16);
                    SNS_IIB_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 69], (int)byteDataAnalysis[96 * i + 70]) * 16);
                    LED_F1_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 72], (int)byteDataAnalysis[96 * i + 73]));
                    T_12inches.Add(DataMakeUp((int)byteDataAnalysis[96 * i + 75], (int)byteDataAnalysis[96 * i + 76]));

                    int SecondResult = 0;
                    for (int j = 0; j < 4; j++)
                    {
                        int SecondOrigin = (Int32)DataMakeUp((int)byteDataAnalysis[96 * i + (26 + j) * 3], (int)byteDataAnalysis[96 * i + (26 + j) * 3 + 1]);
                        SecondResult |= SecondOrigin;
                        if (j < 3)
                        {
                            SecondResult <<= 8;
                        }
                    }
                    Second_12inches.Add(SecondResult);
                }
                else
                {
                    MessageBox.Show("文件格式错误！请确认文件内容", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        /// <summary>
        /// 函数功能：将两个十六进制数组合成一个字节变量
        /// </summary>
        /// <param name="RawData1"></param>
        /// <param name="RawData2"></param>
        /// <returns></returns>
        byte DataMakeUp(int RawData1,int RawData2)
        {
            byte DataAfter1 = DataTypeConversion(RawData1);
            byte DataAfter2 = DataTypeConversion(RawData2);
            
            DataAfter1 <<= 4;
            DataAfter1 |= DataAfter2;

            return DataAfter1;
        }

        /// <summary>
        /// 函数功能：导入存储为ASCII表十进制数的int类型变量，导出对应字符的十六进制数
        /// </summary>
        /// <param name="RawData0"></param>
        /// <returns></returns>
        byte DataTypeConversion(int RawData0)
        {
            byte Result = 0x00;

            switch (RawData0)
            {
                case 48: Result = 0x00; break;
                case 49: Result = 0x01; break;
                case 50: Result = 0x02; break;
                case 51: Result = 0x03; break;
                case 52: Result = 0x04; break;
                case 53: Result = 0x05; break;
                case 54: Result = 0x06; break;
                case 55: Result = 0x07; break;
                case 56: Result = 0x08; break;
                case 57: Result = 0x09; break;
                case 65: Result = 0x0A; break;
                case 66: Result = 0x0B; break;
                case 67: Result = 0x0C; break;
                case 68: Result = 0x0D; break;
                case 69: Result = 0x0E; break;
                case 70: Result = 0x0F; break;               
            }

            return Result;
        }
            
        /// <summary>
        /// 函数功能：8寸灯具参数解析，创建Excel文件
        /// </summary>
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

            //输出各个参数值
            for(int i=0;i< RMS1_8inches.Count;i++)
            {
                ExcelSheet.Cells[3+i, 1] = (i + 1).ToString();
                ExcelSheet.Cells[3+i, 2] = RMS1_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 3] = Val2_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 4] = Val3_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 5] = RMS_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 6] = Current_Ratio1_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 7] = Current_Ratio3_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 8] = RES_IA_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 9] = RES_IIA_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 10] = SNS_IA_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 11] = SNS_IIA_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 12] = LED_F1_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 13] = T_8inches[i].ToString();
                ExcelSheet.Cells[3+i, 14] = ((int)Second_8inches[i] / 3600).ToString() + ":" + (((int)Second_8inches[i] % 3600) / 60).ToString() + ":" + (((int)Second_8inches[i] % 3600) % 60).ToString();         
            }

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序    

            //清空集合内容
            RMS1_8inches.Clear();
            Val2_8inches.Clear();
            Val3_8inches.Clear();
            RMS_8inches.Clear();
            Current_Ratio1_8inches.Clear();
            Current_Ratio3_8inches.Clear();
            RES_IA_8inches.Clear();
            RES_IIA_8inches.Clear();
            SNS_IA_8inches.Clear();
            SNS_IIA_8inches.Clear();
            LED_F1_8inches.Clear();
            T_8inches.Clear();
            Second_8inches.Clear();
            InputData.Clear();

        }

        /// <summary>
        /// 12寸灯具参数解析，创建Excel文件
        /// </summary>
        void TwelveInchesLampParametersCreatExcel()
        {
            //创建excel模板
            str_fileName = "d:\\12寸灯具参数解析" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";    //文件保存路径及名称
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();                          //创建Excel应用程序 ExcelApp
            ExcelDoc = ExcelApp.Workbooks.Add(Type.Missing);                                      //在应用程序ExcelApp下，创建工作簿ExcelDoc
            ExcelSheet = ExcelDoc.Worksheets.Add(Type.Missing);                                   //在工作簿ExcelDoc下，创建工作表ExcelSheet

            //设置Excel列名           
            ExcelSheet.Cells[1, 1] = "12寸灯具参数解析";
            ExcelSheet.Cells[2, 1] = "序号";
            ExcelSheet.Cells[2, 2] = "RMS1";
            ExcelSheet.Cells[2, 3] = "RMS2";
            ExcelSheet.Cells[2, 4] = "Val2";
            ExcelSheet.Cells[2, 5] = "Val3";
            ExcelSheet.Cells[2, 6] = "RMSMID1";
            ExcelSheet.Cells[2, 7] = "RMSMID2";
            ExcelSheet.Cells[2, 8] = "RMS";
            ExcelSheet.Cells[2, 9] = "Current_Ratio1";
            ExcelSheet.Cells[2, 10] = "Current_Ratio2";
            ExcelSheet.Cells[2, 11] = "Current_Ratio3";
            ExcelSheet.Cells[2, 12] = "Current_Ratio4";
            ExcelSheet.Cells[2, 13] = "RES_IA";
            ExcelSheet.Cells[2, 14] = "RES_IB";
            ExcelSheet.Cells[2, 15] = "RES_IIA";
            ExcelSheet.Cells[2, 16] = "RES_IIB";
            ExcelSheet.Cells[2, 17] = "SNS_IA";
            ExcelSheet.Cells[2, 18] = "SNS_IB";
            ExcelSheet.Cells[2, 19] = "SNS_IIA";
            ExcelSheet.Cells[2, 20] = "SNS_IIB";
            ExcelSheet.Cells[2, 21] = "LED_F1";
            ExcelSheet.Cells[2, 22] = "T";
            ExcelSheet.Cells[2, 23] = "Second";

            //输出各个参数值
            for (int i = 0; i < RMS1_12inches.Count; i++)
            {
                ExcelSheet.Cells[3 + i, 1] = (i + 1).ToString();
                ExcelSheet.Cells[3 + i, 2] = RMS1_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 3] = RMS2_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 4] = Val2_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 5] = Val3_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 6] = RMSMID1_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 7] = RMSMID2_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 8] = RMS_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 9] = Current_Ratio1_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 10] = Current_Ratio2_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 11] = Current_Ratio3_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 12] = Current_Ratio4_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 13] = RES_IA_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 14] = RES_IB_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 15] = RES_IIA_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 16] = RES_IIB_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 17] = SNS_IA_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 18] = SNS_IB_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 19] = SNS_IIA_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 20] = SNS_IIB_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 21] = LED_F1_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 22] = T_12inches[i].ToString();
                ExcelSheet.Cells[3 + i, 23] = ((int)Second_12inches[i]/3600).ToString()+ ":"+ (((int)Second_12inches[i] % 3600) / 60).ToString() + ":"+ (((int)Second_12inches[i] %3600)%60).ToString();
            }

            ExcelSheet.SaveAs(str_fileName);                                                      //保存Excel工作表
            ExcelDoc.Close(Type.Missing, str_fileName, Type.Missing);                             //关闭Excel工作簿
            ExcelApp.Quit();                                                                      //退出Excel应用程序            

            //清空集合内容
            RMS1_12inches.Clear();
            RMS2_12inches.Clear();
            Val2_12inches.Clear();
            Val3_12inches.Clear();
            RMSMID1_12inches.Clear();
            RMSMID2_12inches.Clear();
            RMS_12inches.Clear();
            Current_Ratio1_12inches.Clear();
            Current_Ratio2_12inches.Clear();
            Current_Ratio3_12inches.Clear();
            Current_Ratio4_12inches.Clear();
            RES_IA_12inches.Clear();
            RES_IB_12inches.Clear();
            RES_IIA_12inches.Clear();
            RES_IIB_12inches.Clear();
            SNS_IA_12inches.Clear();
            SNS_IB_12inches.Clear();
            SNS_IIA_12inches.Clear();
            SNS_IIB_12inches.Clear();
            LED_F1_12inches.Clear();
            T_12inches.Clear();
            Second_12inches.Clear();
            InputData.Clear();
        }
    }
}
