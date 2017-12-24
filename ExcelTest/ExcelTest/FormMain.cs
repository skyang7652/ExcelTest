using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading;
using System.IO;

namespace ExcelTest
{
    public partial class FormMain : Form
    {
        public string ProjectName = "工程名稱";
        public string WriteDate = "填報日期";                    //填報日期
        public string StartDate = "開工日期";                    //開工日期
        public string EndDate = "預定完工日期";                      //完工日期
        public string AllDate = "契約工期";                         //契約工期
                                                                // public string nCumulativeDate = ;                  //累計工期
                                                                // public string nSurplusDate;                     //剩餘
        public string ScheduleProcessToday = "預定進度";          //預定進度(本日)
                                                              // public string ScheduleProcessCumulative;     //預定進度(累計)

        public string nProjectChangeNum = "契約變更次數";                //契約更變次數
        public string nProjectDalayNum = "工期展延天數";                 //契約展延天數

        public string fActiveProcessToday = "實際進度";            //實際進度(本日)
                                                               // public string fActiveProcessCumulative;       //實際進度(累計)

        public string nOriMoney = "契約\n金額";                        //契約金額(原契約)
                                                                   // public string nChangeMoney;                     //契約金額(更變契約)

        public String ProjectItemStart = "第一號明細表";
        public String ProjectItemEnd = "第二號明細表";

        public String MaterialItemStart = "供給材料名稱";
        public String MaterialItemEnd = "二、監督依照設計圖說施工";

        public String TestItemStart = "應做試體及試驗";
        public String TestItemEnd = "機械車量名稱";

        ItemCode itemCode = new ItemCode();

        List<ProjectLog> projectLogs = new List<ProjectLog>();
        Object oMissing = System.Reflection.Missing.Value;
        IniRead fileIni;
        Excel.Application oXL;
        Excel.Workbook oWB;
        Excel.Worksheet oSheet;
        double ProjectCount = 0;
        double ProcessTemp = 0.0f;
        string sLoadFilePath = null;
        string sSaveFilePath = null;
        int i;
        public FormMain()
        {
            InitializeComponent();
            initialCode();
        }
        
        private void initialCode()
        {
            itemCode.sProjectName = tbProjectName.Text;
            itemCode.sWriteDate = tbWriteDate.Text;
            itemCode.sStartDate = tbStartDate.Text;
            itemCode.sEndDate = tbEndDate.Text;
           // itemCode.sAllDate = tbAllDate.Text;
            itemCode.sChangeNum = tbChangeNum.Text;
            itemCode.sDelayDay = tbDelayDay.Text;
            itemCode.sChangeMoney = tbChangeMoney.Text;
            itemCode.sOriMoney = tbOriMoney.Text;
        
            itemCode.sActiveProcess = tbActiveProcess.Text;
            itemCode.sScheduleProcess = tbScheduleProcess.Text;
            itemCode.nProjectItemStart = (int)nudProjectItemStart.Value;
            itemCode.nProjectItemEnd =(int) nudProjectItemEnd.Value;
            itemCode.sProjectItem = tbProjectItem.Text;
            itemCode.sProjectUnit = tbUnit.Text;
            itemCode.sProjectNum = tbNumber.Text;
            itemCode.sProjectNumToday = tbNumToday.Text;
            itemCode.sProjectNumCumulative = tbNumCumulative.Text;


            for (int i = 0; i < 10; i++)
            {
                System.Windows.Forms.TextBox tbBox = new System.Windows.Forms.TextBox();
                tbBox = (System.Windows.Forms.TextBox)this.Controls.Find("tbRemark" + (i + 1), true)[0];
                itemCode.sRemark[i] = tbBox.Text;

                tbBox = (System.Windows.Forms.TextBox)this.Controls.Find("tbSummary" + (i + 1), true)[0];
                itemCode.sSummary[i] = tbBox.Text;
            }

            itemCode.s140CylinderAll = tb140CylinderAll.Text;
            itemCode.s140CylinderCum = tb140CylinderCum.Text;
            itemCode.s210CylinderAll = tb210CylinderAll.Text;
            itemCode.s210CylinderCum = tb210CylinderCum.Text;

            itemCode.s140DrillingAll = tb140DrillingAll.Text;
            itemCode.s140DrillingCum = tb140DrillingCum.Text;
            itemCode.s210DrillingAll = tb210DrillingAll.Text;
            itemCode.s210DrillingCum = tb210DrillingCum.Text;

            itemCode.sMaterialName = tbMaterialName.Text;
            itemCode.sMaterialNum = tbMaterialNum.Text;
            itemCode.sMaterialUnit = tbMaterialUnit.Text;
            itemCode.sMaterialNumToday = tbMaterialNumToday.Text;
            itemCode.sMaterialNumCumulative = tbMaterialNumCumulative.Text;
            itemCode.sMaterialNumNoUse = tbMaterialNumNoUse.Text;








        }
       
        private void buttonConvert_Click(object sender, EventArgs e)
        {

            projectLogs.Clear();

            itemCode.sProjectName = tbProjectName.Text;
            itemCode.sStartDate = tbStartDate.Text;
            itemCode.sEndDate = tbEndDate.Text;
            itemCode.sAllDate = tbAllDate.Text;
           OpenFileDialog loaddFileDialog = new OpenFileDialog();
            loaddFileDialog.Title = "選擇讀入監工日誌檔案";
            loaddFileDialog.InitialDirectory = ".\\";
            loaddFileDialog.Filter = "xlsx files(*.*)|*.xlsx|xls files (*.*)|*.xls";
            if (loaddFileDialog.ShowDialog() == DialogResult.OK)
            {
                sLoadFilePath = loaddFileDialog.FileName;
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "xlsx files(*.*)|*.xlsx|xls files (*.*)|*.xls";
                saveFileDialog.Title = "選擇匯出施工日誌檔案";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    sSaveFilePath = saveFileDialog.FileName;
                    Thread thread = new Thread(new ThreadStart(startAnalysis)); //模擬進度條
                    thread.IsBackground = true;
                    thread.Start();
                }
            }

        }
        private delegate void ProgressBarShow(int i,int Num ,string status);
        private void ShowPro(int value,int Num,string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new ProgressBarShow(ShowPro), value, Num, status);
            }
            else
            {
                float temp = (float)value / (float)Num * 100;
                this.progressBar1.Maximum = Num;
                this.progressBar1.Value = value;
                this.labelStatus.Text = string.Format("狀態：{0}.....({1} %)", status, temp);
            }
        }

        private void startAnalysis()
        {
            readExcel(sLoadFilePath);
            writeExcel(sSaveFilePath);
            Thread.CurrentThread.Abort();
        }

        private void writeExcel(string saveFile)
        {
            

            oXL = new Excel.Application();
           
            oXL.Visible = false;
            oXL.UserControl = false;
            oWB = oXL.Workbooks.Add(oMissing);
            String fontName = "標楷體";
            if(CBNot.Checked)
            {
                Thread.Sleep(10000);
            }

            int strRight = (int)Microsoft.Office.Core.XlHAlign.xlHAlignRight;
            int strCenter = (int)Microsoft.Office.Core.XlHAlign.xlHAlignCenter;
            int strLeft = (int)Microsoft.Office.Core.XlHAlign.xlHAlignLeft;
            string sTemp;

            //目前工作表的數量
            int currentNumOfSheet = oWB.Worksheets.Count;
            //補足工作表
            for (; currentNumOfSheet < projectLogs.Count; currentNumOfSheet++)
                oWB.Worksheets.Add(oMissing, oMissing, oMissing, oMissing);

            this.ShowPro(0, projectLogs.Count, "轉檔中");
            for (int i = 1; i <= projectLogs.Count; i++)
            {                
                ProjectLog temp = new ProjectLog();
                temp = projectLogs[i - 1];
                oSheet = (Excel.Worksheet)oWB.Sheets[i];
                oSheet.Name = temp.sSheetName;
                
                Excel.PageSetup pageSetup = oSheet.PageSetup;
                //FitToPagesTall property設為false，則以FitToPagesWide來縮放Worksheet
                oXL.PrintCommunication = false;
                pageSetup.FitToPagesTall = false;
                pageSetup.FitToPagesWide = 1;
                oXL.PrintCommunication = true;
                // oSheet.Cells[1, 1] = temp.sProjectName;

                // 第一列

                ((Excel.Range)oSheet.Columns["A:O", System.Type.Missing]).ColumnWidth = 4.63;
                ((Excel.Range)oSheet.Columns["P:U", System.Type.Missing]).ColumnWidth = 2.25;
                ((Excel.Range)oSheet.Rows["1:1", System.Type.Missing]).Cells.RowHeight = 40;

                writeData(temp.sProjectName, "A1", "N1", strRight, fontName, 16,false,false,false,false,false);               
                writeData("施工日誌表", "O1", "U1", strCenter, fontName, 16,false,false,false,false, true);
              

                // 第二列
                ((Excel.Range)oSheet.Rows["2:2", System.Type.Missing]).RowHeight = 16.5;

                writeData("本日天氣:", "A2", "B2", strRight, fontName, 10, false,false,false,false, true);
                writeData("上午:", "C2", "C2", strCenter, fontName, 10, false,false,false,false, true);
                writeData("晴", "D2", "D2", strLeft, fontName, 10, false,false,false,false, true);
                writeData("下午:", "E2", "E2", strCenter, fontName, 10, false,false,false,false, true);
                writeData("晴", "F2", "F2", strLeft, fontName, 10, false,false,false,false, true);
             
                
                writeData("填報日期:", "O2", "Q2", strCenter, fontName, 10, false,false,false,false, true);
                writeData(temp.sWriteDate, "R2", "U2", strCenter, fontName, 10, false,false,false,false, true);

                writeData(temp.sWriteDate, "R2", "U2", strCenter, fontName, 10, false, false, false, false, true);

                // 第三列
                writeData("契約工期", "A3", "B3", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.nAllDate.ToString(), "C3", "D3", strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計工期", "E3", "F3", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.nCumulativeDate.ToString(), "G3", "H3", strCenter, fontName, 10, true, true, true, true, true);
                writeData("天", "I3", "I3", strCenter, fontName, 10, true, true, true, true, true);
                writeData("剩餘", "J3", "K3", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.nSurplusDate.ToString(), "L3", "M3", strCenter, fontName, 10, true, true, true, true, true);
                writeData("天", "N3", "N3", strCenter, fontName, 10, true, true, true, true, true);
                writeData("開工日期", "O3", "Q3", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.sStartDate, "R3", "U3", strCenter, fontName, 10, true, true, true, true, true);


                //第四列
                writeData("契約變更次數", "A4", "D4", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.nProjectChangeNum.ToString(), "E4", "F4", strCenter, fontName, 10, true, true, true, true, true);
                writeData("次", "G4", "G4", strCenter, fontName, 10, true, true, true, true, true);

                writeData("工程展延天數", "H4", "J4", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.nProjectDalayNum.ToString(), "K4", "L4", strCenter, fontName, 10, true, true, true, true, true);
                writeData("次", "M4", "M4", strCenter, fontName, 10, true, true, true, true, true);

                writeData("預定完工日期", "N4", "Q4", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.sEndDate, "R4", "U4", strCenter, fontName, 10, true, true, true, true, true);

                //第五、六列

                writeData("預定進度\n（％）", "A5", "B6", strCenter, fontName, 10, true, true, true, true, true);
                writeData("本日", "C5", "C5", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.fScheduleProcessToday.ToString() + "%", "D5", "F5", strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計", "C6", "C6", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.fScheduleProcessToday.ToString() + "%", "D6", "F6", strCenter, fontName, 10, true, true, true, true, true);

                writeData("實際進度\n（％）", "G5", "H6", strCenter, fontName, 10, true, true, true, true, true);
                writeData("本日", "I5", "I5", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.fActiveProcessToday.ToString() + "%", "J5", "L5", strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計", "I6", "I6", strCenter, fontName, 10, true, true, true, true, true);
                writeData(temp.fActiveProcessCumulative.ToString() + "%", "J6", "L6", strCenter, fontName, 10, true, true, true, true, true);

                
                writeData("契約金額", "M5", "N6", strCenter, fontName, 10, true, true, true, true, true);
                writeData("原契約：", "O5", "O5", strCenter, fontName, 8, true, true, true, false, true);
                writeData(temp.nOriMoney.ToString() + "元整", "P5", "U5", strCenter, fontName, 10, true, true, true, true, true);
                writeData("變更後契約：", "O6", "O6", strCenter, fontName, 5, true, true, true, false, true);

                if(temp.nChangeMoney == 0)
                {
                  writeData("", "P6", "U6", strCenter, fontName, 10, true, true, true, true, true);
                }
                else
                {
                    writeData(temp.nChangeMoney.ToString() + "元整", "P6", "U6", strCenter, fontName, 10, true, true, true, true, true);
                }

                //第七列

                sTemp = "一、依施工計畫書執行按圖施工概況(含約定之重要施工項目及完成數量等):";
                writeData(sTemp, "A7", "U7", strCenter, fontName, 10, true, true, true, true, true);

                //第八列
                writeData("施工項目", "A8", "G8", strCenter, fontName, 10, true, true, true, true, true);
                writeData("單位", "H8", "H8", strCenter, fontName, 10, true, true, true, true, true);
                writeData("契約數量", "I8", "K8", strCenter, fontName, 10, true, true, true, true, true);
                writeData("本日完成數量", "L8", "N8", strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計完成數量", "O8", "S8", strCenter, fontName, 10, true, true, true, true, true);
                writeData("備註", "T8", "U8", strCenter, fontName, 10, true, true, true, true, true);

                //第九列

                int Count = temp.ProjectItem.Count;
                string sMun = null;
                int runCount;
                if (Count  <  36 - 8)
                {
                    runCount = 36 - 8;

                }
                else
                {
                    runCount = Count;
                }
                for(int j = 0; j < runCount; j++)
                {
                    sMun = (j + 8 + 1).ToString();
                    if (j < Count)
                    {
                        writeData(temp.ProjectItem[j].sConstructionProject, "A" + sMun, "G" + sMun, strLeft, fontName, 10, true, true, true, true, false);
                        writeData(temp.ProjectItem[j].sUint, "H" + sMun, "H" + sMun, strCenter, fontName, 10, true, true, true, true, true);
                        writeData(temp.ProjectItem[j].dMumber.ToString(), "I" + sMun, "K" + sMun, strRight, fontName, 10, true, true, true, true, true);
                        writeData(temp.ProjectItem[j].dNumberToday.ToString(), "L" + sMun, "N" + sMun, strRight, fontName, 10, true, true, true, true, true);
                        writeData(temp.ProjectItem[j].dNumberCumulative.ToString(), "O" + sMun, "S" + sMun, strRight, fontName, 10, true, true, true, true, true);
                        writeData("", "T" + sMun, "U" + sMun, strCenter, fontName, 10, true, true, true, true, true);
                    }
                    else
                    {
                        writeData("", "A" + sMun, "G" + sMun, strLeft, fontName, 10, true, true, true, true, true);
                        writeData("", "H" + sMun, "H" + sMun, strCenter, fontName, 10, true, true, true, true, true);
                        writeData("", "I" + sMun, "K" + sMun, strRight, fontName, 10, true, true, true, true, true);
                        writeData("", "L" + sMun, "N" + sMun, strRight, fontName, 10, true, true, true, true, true);
                        writeData("", "O" + sMun, "S" + sMun, strRight, fontName, 10, true, true, true, true, true);
                        writeData("", "T" + sMun, "U" + sMun, strCenter, fontName, 10, true, true, true, true, true);

                    }
                }
                runCount = runCount + 8 + 1;
                //第 37 列
                writeData("材料名稱", "A" + runCount, "G" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("單位", "H" + runCount, "H" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("數量", "I" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("本日使用數量", "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計使用數量", "M" + runCount, "O" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("在庫數量", "P" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);

                Count = temp.MaterialItem.Count;
                runCount++;

                for (int j = 0; j < Count; j++)
                {
                    sMun = (j + runCount + 1).ToString();
                    writeData(temp.MaterialItem[j].sMetrialName, "A" + runCount, "G" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                    writeData(temp.MaterialItem[j].sUint, "H" + runCount, "H" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                    writeData(temp.MaterialItem[j].dMumber.ToString(), "I" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                    writeData(temp.MaterialItem[j].dUseNumToday.ToString(), "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                    writeData(temp.MaterialItem[j].dUseNumCumulative.ToString(), "M" + runCount, "O" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                    writeData(temp.MaterialItem[j].dNoUseNum.ToString(), "P" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);                 
                }
                runCount = runCount + Count;

                writeData("", "A" + runCount, "G" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "H" + runCount, "H" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "I" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "M" + runCount, "O" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "P" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);

                runCount++;

                sTemp = "三、工地人員及機具管理(含約定之出工人數及機具使用情形及數量)";
                writeData(sTemp, "A" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);

                runCount++;

                writeData("工別", "A" + runCount, "C" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("本日人數", "D" + runCount, "F" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計人數", "G" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("機具名稱", "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("本日使用數量", "M" + runCount, "O" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("累計使用數量", "P" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);

                runCount++;

                writeData("", "A" + runCount, "C" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "D" + runCount, "F" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "G" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("挖土機", "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "M" + runCount, "O" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("", "P" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);

                runCount++;

                writeData("應做試體組數", "A" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("應做鑽心組數", "J" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, true, true, true);

                runCount++;


                writeData("140kgf/cm2", "A" + runCount, "C" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("共", "D" + runCount, "D" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n140CylinderAll.ToString(), "E" + runCount, "E" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "F" + runCount, "F" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                writeData("第", "G" + runCount, "G" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n140CylinderCum.ToString(), "H" + runCount, "H" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "I" + runCount, "I" + runCount, strLeft, fontName, 10, true, true, false, true, true);


                writeData("140kgf/cm2", "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("共", "M" + runCount, "M" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n140DrillingAll.ToString(), "N" + runCount, "N" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "O" + runCount, "O" + runCount, strRight, fontName, 10, true, true, false, true, true);

                writeData("第", "P" + runCount, "Q" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n140DrillingCum.ToString(), "R" + runCount, "S" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "T" + runCount, "U" + runCount, strRight, fontName, 10, true, true, false, true, true);


                runCount++;
               
                writeData("210kg/cm2", "A" + runCount, "C" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("共", "D" + runCount, "D" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n210CylinderAll.ToString(), "E" + runCount, "E" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "F" + runCount, "F" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                writeData("第", "G" + runCount, "G" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n210CylinderCum.ToString(), "H" + runCount, "H" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "I" + runCount, "I" + runCount, strLeft, fontName, 10, true, true, false, true, true);


                writeData("210kg/cm2", "J" + runCount, "L" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("共", "M" + runCount, "M" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n210DrillingAll.ToString(), "N" + runCount, "N" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "O" + runCount, "O" + runCount, strRight, fontName, 10, true, true, false, true, true);

                writeData("第", "P" + runCount, "Q" + runCount, strRight, fontName, 10, true, true, true, false, true);
                writeData(temp.n210DrillingCum.ToString(), "R" + runCount, "S" + runCount, strCenter, fontName, 10, true, true, false, false, true);
                writeData("組", "T" + runCount, "U" + runCount, strRight, fontName, 10, true, true, false, true, true);

                runCount++;

                writeData("施工概況", "A" + runCount, "I" + runCount, strCenter, fontName, 10, true, true, true, true, true);
                writeData("記事", "J" + runCount, "U" + runCount, strCenter, fontName, 10, true, true, false, true, true);

                runCount++;

                writeData(temp.Summary, "A47", "I56" , strLeft, fontName, 10, true, true, true, true, true);
                writeData(temp.Remark, "J47", "U56" , strLeft, fontName, 10, true, true, true, true, true);

                Excel.Range TempRange = oSheet.get_Range("A47", "I56");
                TempRange.VerticalAlignment = Microsoft.Office.Core.XlVAlign.xlVAlignTop;
                TempRange = oSheet.get_Range("J47", "U56");
                TempRange.VerticalAlignment = Microsoft.Office.Core.XlVAlign.xlVAlignTop;

                runCount = runCount + 9 + 1;
                sTemp = "四、";
                writeData(sTemp, "A" + runCount, "A" + runCount, strCenter, fontName, 10, true, true, true, false, true);
                sTemp = "本日施工項目是否有需依「營造業專業工程特定施工項目應置之技術士總類、比率或人數標準表」規定應設置技術士之專業工程: □有□無 (此項如勾選”有”，則應填寫後附「公共工程施工日誌之技術士簽章表」)";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                ((Excel.Range)oSheet.Rows[runCount+ ":" + runCount, System.Type.Missing]).Cells.RowHeight = 50;
                runCount++;
                sTemp = "五、";
                writeData(sTemp, "A" + runCount, "A" + runCount, strCenter, fontName, 10, true, true, true, false, true);
                sTemp = "工地職業安全衛生事項之督導、公共環境與安全之維護及其他工地行政事務：";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);
                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 20;
                runCount++;
                sTemp = "六、";
                writeData(sTemp, "A" + runCount, "A" + runCount, strCenter, fontName, 10, true, true, true, false, true);
                sTemp = "施工取樣試驗紀錄：";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);
                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 20;
                runCount++;
                sTemp = "七、";
                writeData(sTemp, "A" + runCount, "A" + runCount, strCenter, fontName, 10, true, true, true, false, true);
                sTemp = "通知協力廠商辦理事項：";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);
                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 20;
                runCount++;
                sTemp = "八、";
                writeData(sTemp, "A" + runCount, "A" + runCount, strCenter, fontName, 10, true, true, true, false, true);
                sTemp = "重要事項紀錄：";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);
                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 20;
                runCount++;
                sTemp = "簽章表：【工地負責人】（註3）：";
                writeData(sTemp, "A" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, true, true, true);

                runCount++;
                sTemp = "註:1.";
                writeData(sTemp, "A" + runCount, "A" + runCount, strRight, fontName, 10, true, true, true, false, true);
                sTemp = "依營造業法第32條第1項第2款規定，工地主任應按日填報施工日誌。";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                runCount++;
                sTemp = "2.";
                writeData(sTemp, "A" + runCount, "A" + runCount, strRight, fontName, 10, true, true, true, false, true);
                sTemp = "本施工日誌格式僅供參考，惟原則應包含上開欄位，各機關亦得依工程性質及契約約定事項自行增訂之。";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                runCount++;
                sTemp = "3.";
                writeData(sTemp, "A" + runCount, "A" + runCount, strRight, fontName, 10, true, true, true, false, true);
                sTemp = "本工程依營造業法第30條規定須置工地主任者，由工地主任簽章；依上開規定免置工地主任者，則由營造業法第32條第2項所定之人員簽章。廠商非屬營造業者，由工地負責人簽章。";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 33;
                runCount++;
                sTemp = "4.";
                writeData(sTemp, "A" + runCount, "A" + runCount, strRight, fontName, 10, true, true, true, false, true);
                sTemp = "契約工期如有修正，應填修正後之契約工期，含展延工期及不計工期天數；如有依契約變更設計，預定進度及實際進度應填變更設計後計算之進度。";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);
                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 33;

                runCount++;
                sTemp = "5.";
                writeData(sTemp, "A" + runCount, "A" + runCount, strRight, fontName, 10, true, true, true, false, true);
                sTemp = "上開重要事項記錄包含（1）主辦機關及監造單位指示 （2）工地遇緊急異常狀況之通報處理情形（3）本日是否由專任工程人員督察按圖施工、解決施工技術問題等。";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);
                ((Excel.Range)oSheet.Rows[runCount + ":" + runCount, System.Type.Missing]).Cells.RowHeight = 33;
                runCount++;
                sTemp = "6.";
                writeData(sTemp, "A" + runCount, "A" + runCount, strRight, fontName, 10, true, true, true, false, true);
                sTemp = "公共工程屬建築物者，請依內政部99年2月5日台內營字第0990800804號令頒之「建築物施工日誌」填寫。";
                writeData(sTemp, "B" + runCount, "U" + runCount, strLeft, fontName, 10, true, true, false, true, true);

                this.ShowPro(i, projectLogs.Count, "轉檔中");
            }
            oWB.SaveAs(saveFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Excel.XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            this.ShowPro(projectLogs.Count, projectLogs.Count, "完成");
            oWB.Close();
            oWB = null;

            oXL.Quit();
            oXL = null;

        }

        private void writeData(String str,String Range1, String Range2, int XlH,String fontName,int Size,bool bTop, bool bBottom,bool bLeft,bool bRight,bool isAutoFit)
        {
            oSheet.get_Range(Range1, Range2).Merge(false); //設定儲存格合併   
            Excel.Range excelRange = oSheet.get_Range(Range1, Range2);

           
            if(isAutoFit)
            {
                excelRange.WrapText = true;//自動換行
                excelRange.ShrinkToFit = false;//字形縮小適合欄框
            }
           else
            {
                excelRange.WrapText = false;
                excelRange.ShrinkToFit = true;//字形縮小適合欄框
            }
            excelRange.EntireRow.AutoFit();

         /*   if(str != null)
            {
                int rate = str.Length / 36;
                excelRange.RowHeight = 16.5 * (rate + 1);
            }*/
            

            if (bTop)
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; 
            }
            else
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            }

            if (bBottom)
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            else
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            }

            if (bLeft)
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            else
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            }

            if (bRight)
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            else
            {
                excelRange.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            }

            excelRange.Font.Size = Size;
            excelRange.HorizontalAlignment = XlH;
            if(str == "0")
            {
                str = "";
            }
            excelRange.Cells[1, 1] = str;
            excelRange.Font.Name = fontName;

        }

        private void readExcel(string loadFile)
        {
            try
            {
                //設置程序運行語言
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                //創建Application
                Excel.Application oXL = new Excel.Application();
                //設置是否顯示警告窗體
                oXL.DisplayAlerts = false;
                //設置是否顯示Excel
                oXL.Visible = false;
                //禁止刷新屏幕
                oXL.ScreenUpdating = false;
                //根據路徑path打開
                oWB = oXL.Workbooks.Open(loadFile, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                //獲取Worksheet對象


                ProcessTemp = 0;



                int sheetCOunt = oWB.Sheets.Count;
                if (CBNot.Checked)
                {
                    Thread.Sleep(10000);
                }

               if(!cbSheetInverse.Checked)
                {
                    for (i = 1; i <= sheetCOunt; i++)
                    {
                        PaseExcel(i, sheetCOunt);
                    }

                }
               else
                {
                    for (i = sheetCOunt; i >= 1; i--)
                    {
                        PaseExcel(i, sheetCOunt);
                    }
                }


                oWB.Close(0);
                oWB = null;

                oXL.Quit();
                oXL = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);

                oWB.Close(0);
                oWB = null;

                oXL.Quit();
                oXL = null;
            }
                    
        }

        private void PaseExcel(int i, int sheetCOunt)
        {
            if(i == 13)
            {
                int x = 0;
            }
            string sItemCode;

            this.ShowPro(i, sheetCOunt, "分析中");

            ProjectLog newProjet = new ProjectLog();

            oSheet = (Worksheet)oWB.Sheets[i];

            sItemCode = itemCode.sActiveProcess;

            string temp = GetExcelString(sItemCode);

            double tmepTodyProcess = StringToDouble(temp);
            if (temp == string.Empty || (tmepTodyProcess - ProcessTemp == 0 && tmepTodyProcess != 100))
            {
                return;
            }
            newProjet.fActiveProcessToday = tmepTodyProcess - ProcessTemp;
            newProjet.fActiveProcessCumulative = tmepTodyProcess;
            ProcessTemp = tmepTodyProcess;


            newProjet.sSheetName = oSheet.Name;
            newProjet.sProjectName = cbInputProjectName.Checked ? tbInputProjectName.Text : GetExcelString(itemCode.sProjectName);

            newProjet.sWriteDate = DateProcess(GetExcelString(itemCode.sWriteDate));


            newProjet.sStartDate = DateProcess(cbInputStartDate.Checked ? tbInputStartDate.Text : GetExcelString(itemCode.sStartDate));

            newProjet.sEndDate = DateProcess(cbInputEndDate.Checked ? tbInputEndDate.Text : GetExcelString(itemCode.sEndDate));

            newProjet.nAllDate = Convert.ToInt32(cbInputEndDate.Checked ? tbInputAllDate.Text : GetExcelString(itemCode.sAllDate));

            DateTime WD = Convert.ToDateTime(newProjet.sWriteDate);
            DateTime SD = Convert.ToDateTime(newProjet.sStartDate);

            TimeSpan ts1 = new TimeSpan(WD.Ticks);
            TimeSpan ts2 = new TimeSpan(SD.Ticks);
            TimeSpan ts = ts1.Subtract(ts2).Duration();

            newProjet.nCumulativeDate = Convert.ToInt32(ts.Days) + 1;
            newProjet.nSurplusDate = newProjet.nAllDate - newProjet.nCumulativeDate;

            newProjet.nProjectChangeNum = StringToInt(GetExcelString(itemCode.sChangeNum));
            newProjet.nProjectDalayNum = StringToInt(GetExcelString(itemCode.sDelayDay));
            newProjet.nOriMoney = StringToInt(cbInputMoney.Checked ? tbInputMoney.Text : GetExcelString(itemCode.sOriMoney));
            newProjet.nChangeMoney = StringToInt(GetExcelString(itemCode.sChangeMoney));

            newProjet.ProjectItem = GetProjectItem();

            newProjet.MaterialItem = GetMaterialItem();

            newProjet.n140CylinderAll = (int)StringToDouble(GetExcelString(itemCode.s140CylinderAll));
            newProjet.n140CylinderCum = (int)StringToDouble(GetExcelString(itemCode.s140CylinderCum));
            newProjet.n140DrillingAll = (int)StringToDouble(GetExcelString(itemCode.s140DrillingAll));
            newProjet.n140DrillingCum = (int)StringToDouble(GetExcelString(itemCode.s140DrillingCum));
            newProjet.n210CylinderAll = (int)StringToDouble(GetExcelString(itemCode.s210CylinderAll));
            newProjet.n210CylinderCum = (int)StringToDouble(GetExcelString(itemCode.s210CylinderCum));
            newProjet.n210DrillingAll = (int)StringToDouble(GetExcelString(itemCode.s210DrillingAll));
            newProjet.n210DrillingCum = (int)StringToDouble(GetExcelString(itemCode.s210DrillingCum));



            for (int j = 0; j < 10; j++)
            {
                temp = GetExcelString(itemCode.sRemark[j]);
                if (temp != String.Empty)
                {
                    newProjet.Remark += temp + "\n";
                }

                temp = GetExcelString(itemCode.sSummary[j]);
                if (temp != String.Empty)
                {
                    newProjet.Summary += temp + "\n";
                }
            }

            if ((newProjet.Remark == "" || newProjet.Remark == string.Empty) && ( newProjet.Summary == "" || newProjet.Summary == string.Empty))
            {
                return;
            }
            projectLogs.Add(newProjet);


        }

        private string GetExcelString(String sRange)
        {
            string result = "";

            if(sRange == String.Empty)
            {
                return result;
            }

            Excel.Range tempRange =  oSheet.get_Range(sRange, sRange);

            result = tempRange.Cells[1, 1].Text;

            return result;
        }

        private List<TestItem> GetTestITem(String StartItem, String EndItem, Excel.Worksheet xlsWorkSheet)
        {
            List<TestItem> tempTestItem = new List<TestItem>();
            Excel.Range Fruits = xlsWorkSheet.get_Range("A1", "AH136");
            Excel.Range tempRange = Fruits.Find(StartItem, oMissing,
                                                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                oMissing, oMissing);
            Excel.Range tempEndRange = Fruits.Find(EndItem, oMissing,
                                                 Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                 Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                 oMissing, oMissing);

            if (tempRange == null || tempEndRange == null)
            {
                return tempTestItem;
            }

            int StartRow = tempRange.Cells.Row;
            int endRow = tempEndRange.Cells.Row;
            int count = endRow - StartRow - 1;

            for(int i = 1;i<= count;i++)
            {
                TestItem temp = new TestItem();

                temp.sTestName = tempRange.Cells[i + 1, 1].Text;
                temp.sUint = tempRange.Cells[i + 1, 2].Text;
                temp.dMumber = StringToDouble(tempRange.Cells[i + 1, 3].Text);
                temp.dUseNumToday = StringToDouble(tempRange.Cells[i + 1, 4].Text);
                temp.dUseNumCumulative = StringToDouble(tempRange.Cells[i + 1, 5].Text);
                tempTestItem.Add(temp);
            }

            return tempTestItem;
        }

        private List<MaterialItem> GetMaterialItem()
        {
            List<MaterialItem> tempMaterialItem = new List<MaterialItem>();
          
            MaterialItem tempItem = new MaterialItem();
                tempItem.sMetrialName = GetExcelString(itemCode.sMaterialName);
                tempItem.sUint = GetExcelString(itemCode.sMaterialUnit);
                tempItem.dMumber = StringToDouble(GetExcelString(itemCode.sMaterialNum));
                tempItem.dUseNumToday = StringToDouble(GetExcelString(itemCode.sMaterialNumToday));
                tempItem.dUseNumCumulative = StringToDouble(GetExcelString(itemCode.sMaterialNumCumulative));
                tempItem.dNoUseNum = StringToDouble(GetExcelString(itemCode.sMaterialNumNoUse));
                tempMaterialItem.Add(tempItem);

            return tempMaterialItem;
        }

        private List<ProjectItem> GetProjectItem()
        {
            List<ProjectItem> tempIProjectItem = new List<ProjectItem>();

            int StartRow = itemCode.nProjectItemStart;
            int endRow = itemCode.nProjectItemEnd;
            ProjectCount = 0;
            for (int i = StartRow; i<= endRow; i++)
            {
                ProjectItem tempItem = new ProjectItem();
                
                tempItem.sConstructionProject = GetExcelString(itemCode.sProjectItem + i);
                tempItem.sUint = GetExcelString(itemCode.sProjectUnit + i);
                tempItem.dMumber = StringToDouble(GetExcelString(itemCode.sProjectNum + i));
                tempItem.dNumberToday = StringToDouble(GetExcelString(itemCode.sProjectNumToday + i));
                tempItem.dNumberCumulative = StringToDouble(GetExcelString(itemCode.sProjectNumCumulative + i));

                ProjectCount += tempItem.dNumberToday;
                tempIProjectItem.Add(tempItem);
            }
                
            return tempIProjectItem;
        }

        private double CulSecondItem(double ProjectCount)
        {
            int StartRow = itemCode.nSecondItemStart;
            int endRow = itemCode.nSecondItemEnd;

            if (endRow - StartRow > 0)
            {
                for (int i = StartRow; i <= endRow; i++)
                {

                    ProjectCount += StringToDouble(GetExcelString(itemCode.sSecondNumToday + i));

                }
            }
            return ProjectCount;
        }

        private string DateProcess( string  date)
        {
            string str = null;
            if(date != string.Empty)
            {
                str = Regex.Replace(date, "[^0-9]", "/");
            }
           
           str = str.TrimStart('/');
           str = str.TrimEnd('/');
            string[] stemp = str.Split('/');
            if(stemp[0].Length < 4 )
            {
                stemp[0] = (Convert.ToInt32(stemp[0]) + 1911).ToString();

                str = stemp[0] + "/" + stemp[1] + "/" + stemp[2];
            }

            return str;
        }

        private int StringToInt(string Str)
        {
            int value = 0;

            if (Str != String.Empty || Str != "")
            {
                Str = Regex.Replace(Str, "[^0-9]", "");
                if (Str != String.Empty)
                {
                    value = Convert.ToInt32(Str);
                }

            }
            return value;
        }

        private double StringToDouble(string Str)
        {
            double value = 0;

            if (Str != String.Empty || Str != "")
            {
                Str = Regex.Replace(Str, "[^0-9.]", "");

                if(Str !=String.Empty)
                {
                    value = Convert.ToDouble(Str);
                }                       
            }
            return value;
        }

        private string ExcelFind(string findText, Excel.Worksheet xlsWorkSheet,int shiftRow,int shiftCol)
        {
            
            Excel.Range Fruits = xlsWorkSheet.get_Range("A1", "AH136");
            Excel.Range tempRange = Fruits.Find(findText, oMissing,
                                                 Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                 Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                 oMissing, oMissing);
            string temp = string.Empty;

            if (tempRange != null)
            {
                if (tempRange.Cells.MergeCells == true)
                {
                    int row = tempRange.MergeArea.Rows.Count - 1;
                    int col = tempRange.MergeArea.Columns.Count - 1;
                    temp = xlsWorkSheet.Cells[tempRange.MergeArea.Row + row + shiftRow, tempRange.MergeArea.Column + col + shiftCol].text;
                }
                else
                {
                    temp = tempRange.Cells[1 + shiftRow, 1 + shiftCol].text;
                }

            }

            return temp;

        }

        private void buttonSave_Click(object sender, EventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "ini files(*.*)|*.ini";
            saveFileDialog.Title = "匯出設定檔案";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileIni = new IniRead(saveFileDialog.FileName);
                updateCode();
                WriteSetting();
                ShowCode();
            }
           
        }

        private void ShowCode()
        {
             nudProjectItemEnd.Value = itemCode.nProjectItemEnd ;
             nudProjectItemStart.Value = itemCode.nProjectItemStart;
             tb140CylinderAll.Text = itemCode.s140CylinderAll;
             tb140CylinderCum.Text = itemCode.s140CylinderCum;
             tb140DrillingAll.Text = itemCode.s140DrillingAll;
             tb140DrillingCum.Text = itemCode.s140DrillingCum;
             tb210CylinderAll.Text = itemCode.s210CylinderAll;
             tb210CylinderCum.Text = itemCode.s210CylinderCum;

             tb210DrillingAll.Text = itemCode.s210DrillingAll;
             tb210DrillingCum.Text = itemCode.s210DrillingCum;
             tbActiveProcess.Text = itemCode.sActiveProcess;
             tbWriteDate.Text = itemCode.sWriteDate;
             tbAllDate.Text = itemCode.sAllDate;
             tbChangeMoney.Text = itemCode.sChangeMoney;
             tbChangeNum.Text = itemCode.sChangeNum;
             tbDelayDay.Text = itemCode.sDelayDay;
             tbEndDate.Text = itemCode.sEndDate;
             tbMaterialName.Text = itemCode.sMaterialName;
             tbMaterialNum.Text = itemCode.sMaterialNum;
             tbMaterialNumCumulative.Text = itemCode.sMaterialNumCumulative;
             tbMaterialNumNoUse.Text = itemCode.sMaterialNumNoUse;
             tbMaterialNumToday.Text = itemCode.sMaterialNumToday;
             tbMaterialUnit.Text = itemCode.sMaterialUnit;
             tbOriMoney.Text = itemCode.sOriMoney;
             tbProjectItem.Text = itemCode.sProjectItem;
             tbProjectName.Text = itemCode.sProjectName;
             tbNumber.Text = itemCode.sProjectNum;
             tbNumCumulative.Text = itemCode.sProjectNumCumulative;
             tbNumToday.Text = itemCode.sProjectNumToday;
             tbUnit.Text = itemCode.sProjectUnit;

            for(int i = 0; i < 10;i++)
            {
                System.Windows.Forms.TextBox tbBox = new System.Windows.Forms.TextBox();
                tbBox = (System.Windows.Forms.TextBox) this.Controls.Find("tbRemark" + (i + 1),true)[0];
                tbBox.Text = itemCode.sRemark[i];

                tbBox = (System.Windows.Forms.TextBox)this.Controls.Find("tbSummary" + (i + 1), true)[0];
                tbBox.Text = itemCode.sSummary[i];
            }
             tbScheduleProcess.Text = itemCode.sScheduleProcess;
             tbStartDate.Text = itemCode.sStartDate;

             tbSecondToDayNum.Text = itemCode.sSecondNumToday;
             nudSecondItemStart.Value = itemCode.nSecondItemStart;
             nudSecondItemEnd.Value = itemCode.nSecondItemEnd;
        }

        private void updateCode()
        {
            itemCode.nProjectItemEnd =(int) nudProjectItemEnd.Value;
            itemCode.nProjectItemStart = (int)nudProjectItemStart.Value;
            itemCode.s140CylinderAll = tb140CylinderAll.Text;
            itemCode.s140CylinderCum = tb140CylinderCum.Text;
            itemCode.s140DrillingAll = tb140DrillingAll.Text;
            itemCode.s140DrillingCum = tb140DrillingCum.Text;
            itemCode.s210CylinderAll = tb210CylinderAll.Text;
            itemCode.s210CylinderCum = tb210CylinderCum.Text;
            itemCode.s210DrillingAll = tb210DrillingAll.Text;
            itemCode.s210DrillingCum = tb210DrillingCum.Text;
            itemCode.sActiveProcess = tbActiveProcess.Text;
            itemCode.sAllDate = tbAllDate.Text;
            itemCode.sWriteDate = tbWriteDate.Text;
            itemCode.sChangeMoney = tbChangeMoney.Text;
            itemCode.sChangeNum = tbChangeNum.Text;
            itemCode.sDelayDay = tbDelayDay.Text;
            itemCode.sEndDate = tbEndDate.Text;
            itemCode.sMaterialName = tbMaterialName.Text;
            itemCode.sMaterialNum = tbMaterialNum.Text;
            itemCode.sMaterialNumCumulative = tbMaterialNumCumulative.Text;
            itemCode.sMaterialNumNoUse = tbMaterialNumNoUse.Text;
            itemCode.sMaterialNumToday = tbMaterialNumToday.Text;
            itemCode.sMaterialUnit = tbMaterialUnit.Text;
            itemCode.sOriMoney = tbOriMoney.Text;
            itemCode.sProjectItem = tbProjectItem.Text;
            itemCode.sProjectName = tbProjectName.Text;
            itemCode.sProjectNum = tbNumber.Text;
            itemCode.sProjectNumCumulative = tbNumCumulative.Text;
            itemCode.sProjectNumToday = tbNumToday.Text;
            itemCode.sProjectUnit = tbUnit.Text;

            for (int i = 0; i < 10; i++)
            {
                System.Windows.Forms.TextBox tbBox = new System.Windows.Forms.TextBox();
                tbBox = (System.Windows.Forms.TextBox)this.Controls.Find("tbRemark" + (i + 1), true)[0];
                itemCode.sRemark[i] = tbBox.Text;

                tbBox = (System.Windows.Forms.TextBox)this.Controls.Find("tbSummary" + (i + 1), true)[0];
                itemCode.sSummary[i] = tbBox.Text;
            }

            itemCode.sScheduleProcess = tbScheduleProcess.Text;
            itemCode.sStartDate = tbStartDate.Text;

            itemCode.nSecondItemStart = (int)nudSecondItemStart.Value;
            itemCode.nSecondItemEnd = (int)nudSecondItemEnd.Value;
            itemCode.sSecondNumToday = tbSecondToDayNum.Text;

        }

        private void WriteSetting()
        {
            fileIni.EraseSection("Rule");
            fileIni.WriteInteger("Rule", "ProjectItemEnd", itemCode.nProjectItemEnd);
            fileIni.WriteInteger("Rule", "ProjectItemStart", itemCode.nProjectItemStart);
            fileIni.WriteString("Rule", "140CylinderAll", itemCode.s140CylinderAll);
            fileIni.WriteString("Rule", "140CylinderCum", itemCode.s140CylinderCum);
            fileIni.WriteString("Rule", "140DrillingAll", itemCode.s140DrillingAll);
            fileIni.WriteString("Rule", "140DrillingCum", itemCode.s140DrillingCum);
            fileIni.WriteString("Rule", "210CylinderAll", itemCode.s210CylinderAll);
            fileIni.WriteString("Rule", "210CylinderCum", itemCode.s210CylinderCum);
            fileIni.WriteString("Rule", "210DrillingAll", itemCode.s210DrillingAll);
            fileIni.WriteString("Rule", "210DrillingCum", itemCode.s210DrillingCum);
            fileIni.WriteString("Rule", "ActiveProcess", itemCode.sActiveProcess);
            fileIni.WriteString("Rule", "AllDate", itemCode.sAllDate);
            fileIni.WriteString("Rule", "ChangeMoney", itemCode.sChangeMoney);
            fileIni.WriteString("Rule", "ChangeNum", itemCode.sChangeNum);
            fileIni.WriteString("Rule", "DelayDay", itemCode.sDelayDay);
            fileIni.WriteString("Rule", "EndDate", itemCode.sEndDate);
            fileIni.WriteString("Rule", "MaterialName", itemCode.sMaterialName);
            fileIni.WriteString("Rule", "MaterialNum", itemCode.sMaterialNum);
            fileIni.WriteString("Rule", "MaterialNumCumulative", itemCode.sMaterialNumCumulative);
            fileIni.WriteString("Rule", "MaterialNumNoUse", itemCode.sMaterialNumNoUse);
            fileIni.WriteString("Rule", "MaterialNumToday", itemCode.sMaterialNumToday);
            fileIni.WriteString("Rule", "MaterialUnit", itemCode.sMaterialUnit);
            fileIni.WriteString("Rule", "OriMoney", itemCode.sOriMoney);
            fileIni.WriteString("Rule", "ProjectItem", itemCode.sProjectItem);
            fileIni.WriteString("Rule", "ProjectName", itemCode.sProjectName);
            fileIni.WriteString("Rule", "ProjectNum", itemCode.sProjectNum);
            fileIni.WriteString("Rule", "ProjectNumCumulative", itemCode.sProjectNumCumulative);
            fileIni.WriteString("Rule", "ProjectNumToday", itemCode.sProjectNumToday);
            fileIni.WriteString("Rule", "ProjectUnit", itemCode.sProjectUnit);
            fileIni.WriteString("Rule", "ProjectName", itemCode.sProjectName);
            fileIni.WriteString("Rule", "ProjectNum", itemCode.sProjectNum);

            for (int i = 0; i < 10; i++)
            {
                fileIni.WriteString("Rule", "Remark" + i, itemCode.sRemark[i]);
                fileIni.WriteString("Rule", "Summary" + i, itemCode.sSummary[i]);
            }
     
            fileIni.WriteString("Rule", "ScheduleProcess", itemCode.sScheduleProcess);
            fileIni.WriteString("Rule", "StartDate", itemCode.sStartDate);
          
            fileIni.WriteString("Rule", "WriteDate", itemCode.sWriteDate);

            fileIni.WriteString("Rule", "SecondItemStart", itemCode.nSecondItemStart.ToString());
            fileIni.WriteString("Rule", "SecondItemEnd", itemCode.nSecondItemEnd.ToString());
            fileIni.WriteString("Rule", "SecondItemNumToday", itemCode.sSecondNumToday);

        }

        private void loadSetting()
        {

            itemCode.nProjectItemEnd = fileIni.ReadInteger("Rule", "ProjectItemEnd", 0);
            itemCode.nProjectItemStart = fileIni.ReadInteger("Rule", "ProjectItemStart", 0);
            itemCode.s140CylinderAll = fileIni.ReadString("Rule", "140CylinderAll",null).TrimEnd('\0');
            itemCode.s140CylinderCum = fileIni.ReadString("Rule", "140CylinderCum", null).TrimEnd('\0');
            itemCode.s140DrillingAll = fileIni.ReadString("Rule", "140DrillingAll", null).TrimEnd('\0');
            itemCode.s140DrillingCum = fileIni.ReadString("Rule", "140DrillingCum", null).TrimEnd('\0');
            itemCode.s210CylinderAll = fileIni.ReadString("Rule", "210CylinderAll", null).TrimEnd('\0');
            itemCode.s210CylinderCum = fileIni.ReadString("Rule", "210CylinderCum", null).TrimEnd('\0');
            itemCode.s210DrillingAll = fileIni.ReadString("Rule", "210DrillingAll", null).TrimEnd('\0');
            itemCode.s210DrillingCum = fileIni.ReadString("Rule", "210DrillingCum", null).TrimEnd('\0');
            itemCode.sActiveProcess = fileIni.ReadString("Rule", "ActiveProcess", null).TrimEnd('\0');
            itemCode.sAllDate = fileIni.ReadString("Rule", "AllDate", null).TrimEnd('\0');           
            itemCode.sChangeMoney = fileIni.ReadString("Rule", "ChangeMoney", null).TrimEnd('\0');
            itemCode.sChangeNum = fileIni.ReadString("Rule", "ChangeNum", null).TrimEnd('\0');
            itemCode.sDelayDay = fileIni.ReadString("Rule", "DelayDay", null).TrimEnd('\0');
            itemCode.sMaterialName = fileIni.ReadString("Rule", "MaterialName", null).TrimEnd('\0');
            itemCode.sMaterialNum = fileIni.ReadString("Rule", "MaterialNum", null).TrimEnd('\0');
            itemCode.sMaterialNumCumulative = fileIni.ReadString("Rule", "MaterialNumCumulative", null).TrimEnd('\0');
            itemCode.sMaterialNumNoUse = fileIni.ReadString("Rule", "MaterialNumNoUse", null).TrimEnd('\0');
            itemCode.sMaterialNumToday = fileIni.ReadString("Rule", "MaterialNumToday", null).TrimEnd('\0');
            itemCode.sMaterialUnit = fileIni.ReadString("Rule", "MaterialUnit", null).TrimEnd('\0');
            itemCode.sOriMoney = fileIni.ReadString("Rule", "OriMoney", null).TrimEnd('\0');
            itemCode.sProjectItem = fileIni.ReadString("Rule", "ProjectItem", null).TrimEnd('\0');
            itemCode.sProjectNum = fileIni.ReadString("Rule", "ProjectNum", null).TrimEnd('\0');
            itemCode.sProjectNumCumulative = fileIni.ReadString("Rule", "ProjectNumCumulative", null).TrimEnd('\0');
            itemCode.sProjectNumToday = fileIni.ReadString("Rule", "ProjectNumToday", null).TrimEnd('\0');
            itemCode.sProjectUnit = fileIni.ReadString("Rule", "ProjectUnit", null).TrimEnd('\0');
            itemCode.sProjectName = fileIni.ReadString("Rule", "ProjectName", null).TrimEnd('\0');
            for(int i = 0;i < 10; i++)
            {
                itemCode.sRemark[i] = fileIni.ReadString("Rule", "Remark" + i, null).TrimEnd('\0');
                itemCode.sSummary[i] = fileIni.ReadString("Rule", "Summary" + i, null).TrimEnd('\0');
            }
                 
            itemCode.sScheduleProcess = fileIni.ReadString("Rule", "ScheduleProcess", null).TrimEnd('\0');
            itemCode.sStartDate = fileIni.ReadString("Rule", "StartDate", null).TrimEnd('\0');        
            itemCode.sWriteDate = fileIni.ReadString("Rule", "WriteDate", null).TrimEnd('\0');
            itemCode.sEndDate = fileIni.ReadString("Rule", "EndDate", null).TrimEnd('\0');

            itemCode.sSecondNumToday = fileIni.ReadString("Rule", "SecondItemNumToday", null).TrimEnd('\0');
            itemCode.nSecondItemStart = fileIni.ReadInteger("Rule", "SecondItemStart", 0);
            itemCode.nSecondItemEnd = fileIni.ReadInteger("Rule", "SecondItemEnd", 0);
        }

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog loaddFileDialog = new OpenFileDialog();
            loaddFileDialog.Title = "讀入設定檔";
            loaddFileDialog.InitialDirectory = ".\\";
            loaddFileDialog.Filter = "ini files(*.*)|*.ini";
            if (loaddFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileIni = new IniRead(loaddFileDialog.FileName);
                labelSetting.Text = Path.GetFileName(loaddFileDialog.FileName);
                loadSetting();
                
                ShowCode();

            }
        }

        private void tbRemark6_TextChanged(object sender, EventArgs e)
        {

        }
    }



    class ProjectLog
    {
        public string sSheetName;
        public string sProjectName;                  //施工名稱
        public string sWriteDate;                    //填報日期
        public string sStartDate;                    //開工日期
        public string sEndDate;                      //完工日期
        public int nAllDate;                         //契約工期
        public int nCumulativeDate;                  //累計工期
        public int nSurplusDate;                     //剩餘
        public double fScheduleProcessToday;          //預定進度(本日)
        public double fScheduleProcessCumulative;     //預定進度(累計)

        public int nProjectChangeNum;                //契約更變次數
        public int nProjectDalayNum;                 //契約展延天數

        public double fActiveProcessToday;            //實際進度(本日)
        public double fActiveProcessCumulative;       //實際進度(累計)

        public int nOriMoney;                        //契約金額(原契約)
        public int nChangeMoney;                     //契約金額(更變契約)

        public List<ProjectItem> ProjectItem;
        public List<MaterialItem> MaterialItem;


        public int n140CylinderAll;
        public int n140CylinderCum;
        public int n210CylinderAll;
        public int n210CylinderCum;

        public int n140DrillingAll;
        public int n140DrillingCum;
        public int n210DrillingAll;
        public int n210DrillingCum;


        public String Remark;
        public String Summary;

    }

    class  ItemCode
    {
        public string sProjectName;
        public string sWriteDate;
        public string sAllDate;
        public string sStartDate;
        public string sEndDate;
        public string sChangeNum;
        public string sDelayDay;
        public string sScheduleProcess;
        public string sActiveProcess;
        public string sOriMoney;
        public string sChangeMoney;

        public int nProjectItemStart;
        public int nProjectItemEnd;

        public string sProjectItem;
        public string sProjectUnit;
        public string sProjectNum;
        public string sProjectNumToday;
        public string sProjectNumCumulative;

        public int nSecondItemStart;
        public int nSecondItemEnd;
        public string sSecondNumToday;

        public string sMaterialName;
        public string sMaterialUnit;
        public string sMaterialNum;
        public string sMaterialNumToday;
        public string sMaterialNumCumulative;
        public string sMaterialNumNoUse;


        public string [] sRemark = new string[10];


        public string [] sSummary = new string[10];


        public string s140CylinderAll;
        public string s140CylinderCum;
        public string s210CylinderAll;
        public string s210CylinderCum;

        public string s140DrillingAll;
        public string s140DrillingCum;
        public string s210DrillingAll;
        public string s210DrillingCum;


    }








    class ProjectItem
    {
        public string sConstructionProject;           //施工項目
        public string sUint;                          //單位
        public double dMumber;                           //契約數量
        public double dNumberToday;                      //本日完成數量
        public double dNumberCumulative;                 //累計完成數量
        public string sRemark;                        //備註

    }

    class MaterialItem
    {
        public string sMetrialName;                    //材料名稱
        public string sUint;                            //單位
        public double dMumber;                             //數量
        public double dUseNumToday;                      //本日使用數量
        public double dUseNumCumulative;                 //累計使用數量
        public double dNoUseNum;                         //在庫數量
    }

    class TestItem
    {
        public string sTestName;                    //應做試體及試驗
        public string sUint;                            //單位
        public double dMumber;                             //數量
        public double dUseNumToday;                      //本日使用數量
        public double dUseNumCumulative;                 //累計使用數量
    }

    


}
