using System;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
using StanleyDriver_RS232;
using HslCommunication.Profinet.Melsec;
using ClosedXML.Excel;
using EasyModbus;
using TaigeneDataCollection;
using System.IO;
using System.Text;

namespace _18_203
{
    public partial class Form1 : Form
    {
        #region --定義初始化檔案--
        private string initailFilePath = @"D:\RP20-701\Resource\initailFile.xlsx";
        private string FilePathParts1 = @"D:\RP20-701\Resource\Parts1.xlsx";
        private string filePath = @"D:\DataFile\";
        private string timerfilename = @"D:\RP20-701\Resource\record.txt";
        private string Parts1IDCheckCode = "";
        private int BarcodeHeader=3;
        private int NumOfAxis=3;
        private string MainIDCheckCode = "";
        TextBox[] tbScrewDataShowTorqueValue;
        TextBox[] tbScrewDataShowTorqueResult;
        TextBox[] tbScrewDataShowAngleValue;
        TextBox[] tbScrewDataShowAngleResult;
        #endregion
        #region --定義PLC--
        private MelsecMcNet PLC1;
        private string PlcIpAddress="10.5.3.112";
        private int PlcPort=6000;
        //PLC->PC
        private string PLCConneted = "D6001";
        private string PLCBarcodeCheckStartFlag = "D6002";
        private string PLCResetCountStatus = "D6003";
        #endregion
        #region --定義序列埠--
        private SerialPort[] spAx = new SerialPort[7];
        #endregion
        #region --定義鎖付--
        private StanleyScrewData ssd;
        #endregion
        #region --定義機器相關--
        private string oldMainIDCheckCode = "";
        private string stringNoBarcode = "NoBarcode";
        private int NoBarcodeSeriealNo = 0, DataSaveCount = 0;
        private int SaveCountAx1 = 0, SaveCountAx2 = 0, SaveCountAx3 = 0, SaveCountAx4 = 0;
        //private ObjectBarcodeAndCount CurrentcountParts1 = new ObjectBarcodeAndCount();
        //private ObjectBarcodeAndCount CurrentcountSite = new ObjectBarcodeAndCount();
        //private ObjectBarcodeAndCount CurrentcountRS = new ObjectBarcodeAndCount();
        //private ObjectBarcodeAndCount CurrentcountTopCap = new ObjectBarcodeAndCount();
        //private ObjectBarcodeAndCount CurrentcountBottomCap = new ObjectBarcodeAndCount();
        #endregion
        #region --定義上報資料相關--
        private const int numOfParts = 3;
        private csDataCollection TCSData;
        private ModbusServer TCS;
        private Int16 _machineRunSetting = 0;
        int _PowerOnTime;
        #endregion
        #region --Form--
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            tbScrewDataShowTorqueValue = new TextBox[4] { tbTorqueValue_1, tbTorqueValue_2, tbTorqueValue_3, tbTorqueValue_4 };
            tbScrewDataShowTorqueResult = new TextBox[4] { tbTorqueResult_1, tbTorqueResult_2, tbTorqueResult_3, tbTorqueResult_4 };
            tbScrewDataShowAngleValue = new TextBox[4] { tbAngleValue_1, tbAngleValue_2, tbAngleValue_3, tbAngleValue_4 };
            tbScrewDataShowAngleResult = new TextBox[4] { tbAngleResult_1, tbAngleResult_2, tbAngleResult_3, tbAngleResult_4 };
            //初始檔案取得
            XLWorkbook wb = new XLWorkbook(initailFilePath);
            var ws = wb.Worksheet(1);
            //計時檔案取得
            GetTimerFile();
            //物件品號
            //NumOfAxis = Convert.ToInt32(ws.Cell("A2").Value);
            //PlcIpAddress = Convert.ToString(ws.Cell("B2").Value);
            //PlcPort = Convert.ToInt32(ws.Cell("C2").Value);
            //BarcodeHeader = Convert.ToInt32(ws.Cell("D2").Value);
            Parts1IDCheckCode = Convert.ToString(ws.Cell("E2").Value);
            MainIDCheckCode = Convert.ToString(ws.Cell("F2").Value);
            //MainIDCheckCode = MainIDCheckCode.Remove(3);
            oldMainIDCheckCode = MainIDCheckCode;
            //取得各物件數量條碼
            ssd = new StanleyScrewData(NumOfAxis);
            textBoxPartsID1.GotFocus += Tb_ScrewBarcode_GotFocus;
            textBoxPartsID2.GotFocus += TextBoxPartID2_GotFocus;
            //*****************************
            //建立通訊埠
            //*****************************
            for (int i = 1; i <= 6; i++)
            {
                spAx[i] = new SerialPort();
                spAx[i].PortName = "COM" + i.ToString();
                spAx[i].BaudRate = 9600;
                spAx[i].Parity = Parity.None;
                spAx[i].DataBits = 8;
                spAx[i].StopBits = StopBits.One;
            }
            //手持式條碼機(螺絲)
            spAx[5] = new SerialPort();
            spAx[5].PortName = "COM5";
            spAx[5].BaudRate = 115200;
            spAx[5].Parity = Parity.Even;
            spAx[5].DataBits = 8;
            spAx[5].StopBits = StopBits.One;
            //固定式條碼機(工件)
            spAx[6] = new SerialPort();
            spAx[6].PortName = "COM6";
            spAx[6].BaudRate = 115200;
            spAx[6].Parity = Parity.Even;
            spAx[6].DataBits = 8;
            spAx[6].StopBits = StopBits.One;
            #region --通訊埠開啟,開發可PASS--
            for (int i = 1; i <= 6; i++)
            {
                spAx[i].Open();
            }
            #endregion
            spAx[1].DataReceived += ax1_DataReceived;
            spAx[2].DataReceived += ax2_DataReceived;
            spAx[3].DataReceived += ax3_DataReceived;
            spAx[4].DataReceived += ax4_DataReceived;
            spAx[5].DataReceived += Hand_DataReceived;
            spAx[6].DataReceived += Fix_DataReceived;
            //建立PLC通訊
            PLC1 = new MelsecMcNet();
            PLC1.IpAddress = PlcIpAddress;
            PLC1.Port = PlcPort;
            PLC1.ConnectTimeOut = 2000;
            PLC1.NetworkNumber = 0x00;
            PLC1.NetworkStationNumber = 0x00;
            //Server
            TCSData = new csDataCollection(numOfParts, NumOfAxis);
            TCS = new ModbusServer();
            TCS.Listen();
            SetBasicSetting();
            timer1.Enabled = true;
            t1s.Interval = 1000;
            t1s.Tick += T1s_Tick;
            t1s.Start();
            
        }
        private void TextBoxPartID2_GotFocus(object sender, EventArgs e)
        {
            currentTextBox = sender as TextBox;
        }
        private void GetTimerFile()
        {
            string[] ss = File.ReadAllLines(timerfilename);
            _PowerOnTime = Convert.ToInt32(ss[0]);
        }
        private void PutTimerFile()
        {
            File.WriteAllText(timerfilename, _PowerOnTime.ToString());
        }
        private void Tb_ScrewBarcode_GotFocus(object sender, EventArgs e)
        {
            currentTextBox = sender as TextBox;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //myData.SaveTime();
            try
            {
                timer1.Enabled = false;
                t1s.Stop();
                PutTimerFile();
                TCS.StopListening();
                foreach (var item in spAx)
                {
                    item.Close();
                }
            }
            catch (Exception)
            {
                ;
            }

        }
        #endregion
        #region --序列埠相關--
        TextBox currentTextBox;
        private bool CheckMainIDIsOK(string mainID)
        {
            string cutstring = mainID.Remove(3);
            if (cutstring.Equals(MainIDCheckCode.Remove(3)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //固定條碼
        private void Fix_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            ssd.ItemBarcode = "";
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                string ss = System.Text.Encoding.ASCII.GetString(buffer);
                //資料丟入memory
                ssd.ItemBarcode = ss;
                ssd.Parts1Barcode = textBoxPartsID1.Text;
                ssd.Parts2Barcode = textBoxPartsID2.Text;
                //檢查BARCODE
                MainIDCheckCode = PLC1.ReadString("D5018", 10).Content;
                TCSLoseControl = PLC1.ReadBool("M1200").Content;
                if (TCSLoseControl)
                {
                    bool result = CheckMainIDIsOK(ss);
                    if (result)
                    {
                        //UPDATE
                        TCSData.setMainID(ss);
                        TCSData.FlagMainIDChanged = 7001;
                        wPLCs[12] = 2000;
                    }
                    else
                    {
                        wPLCs[12] = 3000;
                    }

                }
                else
                {
                    //UPDATE
                    TCSData.setMainID(ss);
                    TCSData.FlagMainIDChanged = 7001;
                }
                PLC1.Write("D1012", ss);
                DisplayTextBox(tb_ItemBarcode, ssd.ItemBarcode);
                ssd.checkFileExist(filePath, ssd.ItemBarcode);
            }
            catch (Exception error)
            {
                DisplayTextBox(tb_Message, "");
                DisplayTextBox(tb_Message, $"固定槍讀取:{error.ToString()}");
            }
        }
        //手持條碼
        private void Hand_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                string ss = "";
                ss = System.Text.Encoding.ASCII.GetString(buffer);
                //檢查barcode
                if (currentTextBox == textBoxPartsID1)
                {
                    //UPDATE
                    TCSData.PartIDs[0].setPartID(ss);
                    TCSData.FlagPLCPartsID = 2001;
                    //HMI
                    PLC1.Write("D1524", ss);
                    DisplayTextBox(textBoxPartsID1, ss);
                }
                else
                {
                    if (currentTextBox==textBoxPartsID2)
                    {
                        //UPDATE
                        TCSData.PartIDs[1].setPartID(ss);
                        TCSData.FlagPLCPartsID = 2001;
                        //HMI
                        PLC1.Write("D2048", ss);
                        DisplayTextBox(textBoxPartsID2, ss);
                    }
                    else
                    {
                        if (currentTextBox == tbParts3Id)
                        {
                            //UPDATE
                            TCSData.PartIDs[2].setPartID(ss);
                            TCSData.FlagPLCPartsID = 2001;
                            //HMI
                            //PLC1.Write("D2048", ss);
                            DisplayTextBox(tbParts3Id, ss);
                        }
                    }
                }
            }
            catch (Exception error)
            {
                DisplayTextBox(tb_Message, "");
                DisplayTextBox(tb_Message, $"手持條碼:{error.ToString()}");
            }
        }
        //軸4
        private void ax4_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                ssd.GetRs232ScrewData(buffer, 4);
                ssd._sd[3].ScrewDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                //GetObjectBarcode();
                //show
                DisplayTextBox(tb_Datetime_Ax4, ssd._sd[3].ScrewDateTime);
                DisplayTextBox(tb_Torque_Ax4, ssd._sd[3].TorqueResult);
                DisplayTextBox(tb_Angle_Ax4, ssd._sd[3].AngleResult);
                DisplayTextBox(tb_Overrall_Ax4, ssd._sd[3].OverrallStatus);
                //傳到myDate
                TCSData.ScrewRows[3].SetTorqueValue(ssd._sd[3].TorqueResult);
                TCSData.ScrewRows[3].SetTorqueJudgement(ssd._sd[3].TorqueStatus);
                TCSData.ScrewRows[3].SetAngleValue(ssd._sd[3].AngleResult);
                TCSData.ScrewRows[3].SetAngleJudgement(ssd._sd[3].AngleStatus);

                SaveCountAx4++;
            }
            catch (Exception error)
            {
                DisplayTextBox(tb_Message, "");
                DisplayTextBox(tb_Message, $"AX4:{error.ToString()}");
            }
        }
        //軸3
        private void ax3_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                ssd.GetRs232ScrewData(buffer, 3);
                ssd._sd[2].ScrewDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                //GetObjectBarcode();
                //show
                DisplayTextBox(tb_Datetime_Ax3, ssd._sd[2].ScrewDateTime);
                DisplayTextBox(tb_Torque_Ax3, ssd._sd[2].TorqueResult);
                DisplayTextBox(tb_Angle_Ax3, ssd._sd[2].AngleResult);
                DisplayTextBox(tb_Overrall_Ax3, ssd._sd[2].OverrallStatus);
                //傳到myDate
                TCSData.ScrewRows[2].SetTorqueValue(ssd._sd[2].TorqueResult);
                TCSData.ScrewRows[2].SetTorqueJudgement(ssd._sd[2].TorqueStatus);
                TCSData.ScrewRows[2].SetAngleValue(ssd._sd[2].AngleResult);
                TCSData.ScrewRows[2].SetAngleJudgement(ssd._sd[2].AngleStatus);

                SaveCountAx3++;
            }
            catch (Exception error)
            {
                DisplayTextBox(tb_Message, "");
                DisplayTextBox(tb_Message, $"AX3:{error.ToString()}");
            }
        }
        //軸2
        private void ax2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                ssd.GetRs232ScrewData(buffer, 2);
                ssd._sd[1].ScrewDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                //GetObjectBarcode();
                //show
                DisplayTextBox(tb_Datetime_Ax2, ssd._sd[1].ScrewDateTime);
                DisplayTextBox(tb_Torque_Ax2, ssd._sd[1].TorqueResult);
                DisplayTextBox(tb_Angle_Ax2, ssd._sd[1].AngleResult);
                DisplayTextBox(tb_Overrall_Ax2, ssd._sd[1].OverrallStatus);
                //傳到myDate
                TCSData.ScrewRows[1].SetTorqueValue(ssd._sd[1].TorqueResult);
                TCSData.ScrewRows[1].SetTorqueJudgement(ssd._sd[1].TorqueStatus);
                TCSData.ScrewRows[1].SetAngleValue(ssd._sd[1].AngleResult);
                TCSData.ScrewRows[1].SetAngleJudgement(ssd._sd[1].AngleStatus);

                SaveCountAx2++;
            }
            catch (Exception error)
            {
                DisplayTextBox(tb_Message, "");
                DisplayTextBox(tb_Message, $"AX2:{error.ToString()}");
            }
        }
        //軸1
        private void ax1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                ssd.GetRs232ScrewData(buffer, 1);
                ssd._sd[0].ScrewDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                //GetObjectBarcode();
                //show
                DisplayTextBox(tb_Datetime_Ax1, ssd._sd[0].ScrewDateTime);
                DisplayTextBox(tb_Torque_Ax1, ssd._sd[0].TorqueResult);
                DisplayTextBox(tb_Angle_Ax1, ssd._sd[0].AngleResult);
                DisplayTextBox(tb_Overrall_Ax1, ssd._sd[0].OverrallStatus);
                //傳到myDate
                TCSData.ScrewRows[0].SetTorqueValue(ssd._sd[0].TorqueResult);
                TCSData.ScrewRows[0].SetTorqueJudgement(ssd._sd[0].TorqueStatus);
                TCSData.ScrewRows[0].SetAngleValue(ssd._sd[0].AngleResult);
                TCSData.ScrewRows[0].SetAngleJudgement(ssd._sd[0].AngleStatus);

                SaveCountAx1++;
            }
            catch (Exception error)
            {
                DisplayTextBox(tb_Message, "");
                DisplayTextBox(tb_Message, $"AX1:{error.ToString()}");
            }
        }
        #endregion
        #region --Timer--
        System.Windows.Forms.Timer t1s = new System.Windows.Forms.Timer();
        bool TCSLoseControl = false;
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            DataSaveCount = SaveCountAx1 + SaveCountAx2 + SaveCountAx3 + SaveCountAx4;
            //檢查是否全部資料都存檔,供給不使用本體條碼用
            if (DataSaveCount >= NumOfAxis)
            {
                ssd.SaveScrewData(filePath, ssd.ItemBarcode, NumOfAxis);
                SaveCountAx1 = 0;
                SaveCountAx2 = 0;
                SaveCountAx3 = 0;
                SaveCountAx4 = 0;
            }
            //PLC讀取
            ReadPLCDevice();
            EnableOpWorkLeavel();
            EnableWorkTypenPartsStock();
            EnablePartsControl();
            EnableProductControl();
            FlagMainIDChanged();
            WritePLCDevice();
            //Server
            Server2UpdateData(40000);//更新Modbus資料
            UpdateTheData();//更新上報資料
            MappingData(40000, TCSData);//更新Modbus資料
            ShowPartsCount();
            //ClearFlagPartsID();
            ReflashDisplay();
            timer1.Start();
        }
        private void T1s_Tick(object sender, EventArgs e)
        {
            t1s.Stop();
            _PowerOnTime++;
            TimeSpan ts = new TimeSpan(0, 0, _PowerOnTime);
            TCSData.BasicSettingYear = (short)DateTime.Now.Year;
            TCSData.BasicSettingMonth = (short)DateTime.Now.Month;
            TCSData.BasicSettingDay = (short)DateTime.Now.Day;
            TCSData.BasicSettingHour = (short)DateTime.Now.Hour;
            TCSData.BasicSettingMin = (short)DateTime.Now.Minute;
            TCSData.BasicSettingSec = (short)DateTime.Now.Second;
            TCSData.OeePowerOnHour = (short)ts.Hours;
            TCSData.OeePowerOnMin = (short)ts.Minutes;
            TCSData.OeePowerOnSec = (short)ts.Seconds;
            t1s.Start();
        }
        //人員操作資格
        private void EnableOpWorkLeavel()
        {
            if (TCSData.EnableOperator == 1)
            {
                //HMI
                wPLCs[0] = 0;
                //UPDATE
                //KH
                //SAVE
            }
            else
            {
                //HMI
                wPLCs[0] = 1;
                //UPDATE
                //KH
                //SAVE
            }
        }
        //設備機種,材料
        private void EnableWorkTypenPartsStock()
        {
            if (TCSData.EnableWorkTypesOrStockOut == 1001)
            {
                wPLCs[1] = 0;
                wPLCs[2] = 0;
            }
            else
            {
                if (TCSData.EnableWorkTypesOrStockOut == 1011)
                {
                    wPLCs[1] = 1;
                }
                else
                {
                    if (TCSData.EnableWorkTypesOrStockOut == 1111)
                    {
                        wPLCs[2] = 1;
                    }
                }
            }
        }
        //TCS 材料控管
        private void EnablePartsControl()
        {
            if (TCSData.EnableParts == 2001)
            {
                wPLCs[3] = 0;
                wPLCs[4] = 0;
            }
            else
            {
                if (TCSData.EnableParts == 2011)
                {
                    wPLCs[3] = 1;
                }
                else
                {
                    if (TCSData.EnableParts == 2111)
                    {
                        wPLCs[4] = 1;
                    }
                }
            }
        }
        //TCS 生產管控
        private void EnableProductControl()
        {
            if (TCSData.EnableMachineRun == 3001)
            {
                wPLCs[5] = 0;
                wPLCs[6] = 0;
                wPLCs[7] = 0;
                wPLCs[8] = 0;
                wPLCs[9] = 0;
            }
            else
            {
                if (TCSData.EnableMachineRun == 3011)
                {
                    wPLCs[5] = 1;
                }
                else
                {
                    if (TCSData.EnableMachineRun == 3111)
                    {
                        wPLCs[6] = 1;
                    }
                    else
                    {
                        if (TCSData.EnableMachineRun == 3100)
                        {
                            wPLCs[7] = 1;
                        }
                        else
                        {
                            if (TCSData.EnableMachineRun == 3101)
                            {
                                wPLCs[8] = 1;
                            }
                            else
                            {
                                if (TCSData.EnableMachineRun == 3110)
                                {
                                    wPLCs[9] = 1;
                                }
                            }
                        }
                    }
                }
            }
        }
        //上位控制
        private void pbUpperControl_Click(object sender, EventArgs e)
        {
            TCSLoseControl = PLC1.ReadBool("M1200").Content;
            TCSLoseControl = !TCSLoseControl;
            if (TCSLoseControl)
            {
                pbUpperControl.Text = "上位未控制";
            }
            else
            {
                pbUpperControl.Text = "上位控制中";
            }
            PLC1.Write("M1200", TCSLoseControl);
        }
        //顯示材料數量
        private void ShowPartsCount()
        {
            if (TCSData.PartIDs[0]!=null)
            {
                lbScrewCount.Text = TCSData.PartIDs[0].PartsCount.ToString();
            }
        }
        //MAIN ID 切換
        private void FlagMainIDChanged()
        {
            wPLCs[10] = TCSData.FlagMainIDChanged;
        }
        #endregion
        #region --螢幕顯示--
        private void ReflashDisplay()
        {
            #region --上報資料頁--
            //上端
            tbEnableOp.Text = TCSData.EnableOperator.ToString();
            tbEnableWorkTypeOrStockOut.Text = TCSData.EnableWorkTypesOrStockOut.ToString();
            tbEnableParts.Text = TCSData.EnableParts.ToString();
            tbEnableMachineRun.Text = TCSData.EnableMachineRun.ToString();
            tbFlagTCSRunFinish.Text = TCSData.FlagMainIDChanged.ToString();
            tbFlagTCSPartsID.Text = TCSData.FlagWorkTypeChanged.ToString();
            //狀態
            tbMachineStatus.Text = TCSData.MachineStatus.ToString();
            tbWorkID.Text = TCSData.WorkID.ToString();
            tbFlagPLCPartsID.Text = TCSData.FlagPLCPartsID.ToString();
            tbFlagPLCRunFinish.Text = TCSData.FlagPLCRunFinish.ToString();
            //基本設定
            tbBsettingLineNo.Text = TCSData.BasicSettingLineNo.ToString();
            tbBsettingWorkID.Text = TCSData.BasicSettingWorkType[0].ToString() + TCSData.BasicSettingWorkType[1].ToString();
            tbBSettingYear.Text = TCSData.BasicSettingYear.ToString();
            tbBSettingMonth.Text = TCSData.BasicSettingMonth.ToString();
            tbBSettingDay.Text = TCSData.BasicSettingDay.ToString();
            tbBSettingHour.Text = TCSData.BasicSettingHour.ToString();
            tbBSettingMin.Text = TCSData.BasicSettingMin.ToString();
            tbBSettingSec.Text = TCSData.BasicSettingSec.ToString();
            tbBsettingStationID.Text = TCSData.BasicSettingEquipmentNo.ToString();
            //人員ID
            //tbStaffReadFlag.Text = myData.StaffReadFlag.ToString();
            //tbStaffID.Text = myData.GetStaffID();
            //材料ID
            tbPartsData_1.Text = TCSData.PartIDs[0].getPartID();
            tbPartsCount_1.Text = TCSData.PartIDs[0].PartsCount.ToString();
            tbPartsData_2.Text = TCSData.PartIDs[1].getPartID();
            tbPartsCount_2.Text = TCSData.PartIDs[1].PartsCount.ToString();
            tbPartsData_3.Text = TCSData.PartIDs[2].getPartID();
            tbPartsCount_3.Text = TCSData.PartIDs[2].PartsCount.ToString();
            //OEE
            tbOeePowerOnHours.Text = TCSData.OeePowerOnHour.ToString();
            tbOeePowerOnMins.Text = TCSData.OeePowerOnMin.ToString();
            tbOeePowerOnSecs.Text = TCSData.OeePowerOnSec.ToString();
            tbOeeMachineRunHours.Text = TCSData.OeeMachineRunHour.ToString();
            tbOeeMachineRunMins.Text = TCSData.OeeMachineRunMin.ToString();
            tbOeeMachineRunSecs.Text = TCSData.OeeMachineRunSec.ToString();
            tbOeeCycleTime.Text = TCSData.OeeCycleTime.ToString();
            tbOeeOutputCount.Text = TCSData.OeeTotalCount.ToString();
            tbOeeOKPartsCount.Text = TCSData.OeeOkCount.ToString();
            tbOeeNGPartsCount.Text = TCSData.OeeNgCount.ToString();
            tbOeeResetCountHour.Text = TCSData.OeeResetCountHour.ToString();
            tbOeeResetCountMin.Text = TCSData.OeeResetCountMin.ToString();
            tbOeeResetCountSec.Text = TCSData.OeeResetCountSec.ToString();
            //製程ID
            tbMainID1.Text = TCSData.getMainID();
            //設備參數及測試數據
            for (int i = 0; i < NumOfAxis; i++)
            {
                tbScrewDataShowTorqueValue[i].Text= TCSData.ScrewRows[i].TorqueValue.ToString();
                tbScrewDataShowTorqueResult[i].Text= TCSData.ScrewRows[i].TorqueJudgement.ToString();
                tbScrewDataShowAngleValue[i].Text = TCSData.ScrewRows[i].AngleValue.ToString();
                tbScrewDataShowAngleResult[i].Text = TCSData.ScrewRows[i].AngleJudgement.ToString();
            }
            #endregion
        }
        delegate void displayTextBoxCallback(TextBox tb,string showText);
        private void DisplayTextBox(TextBox tb,string showText)
        {
            if (tb.InvokeRequired)
            {
                displayTextBoxCallback d = new displayTextBoxCallback(DisplayTextBox);
                this.Invoke(d, new object[] { tb, showText });
            }
            else
            {
                tb.Text = showText;
            }
        }
        #endregion
        #region --PLC相關--
        short[] rPLCs = new short[100];
        short[] wPLCs = new short[100];
        //讀取PLC
        private void ReadPLCDevice()
        {
            //取得PLC狀態
            rPLCs = PLC1.ReadInt16("W0", 100).Content;
            TCSData.MachineStatus = rPLCs[0];
            TCSData.WorkID = rPLCs[1];
            TCSData.FlagWorkTypeChanged = rPLCs[2];
            if (rPLCs[2]==2001)
            {
                SetBasicSetting();
            }
            TCSData.OeeCycleTime = rPLCs[3];
            TCSData.OverrallJudgement = rPLCs[4];
            TCSData.OeeTotalCount = rPLCs[5];
            TCSData.OeeOkCount = rPLCs[6];
            TCSData.OeeNgCount = rPLCs[7];
            if (TCSData.FlagPLCRunFinish==0)
            {
                TCSData.FlagPLCRunFinish = rPLCs[8];
            }
            
            //清除count
            short st= PLC1.ReadInt16(PLCResetCountStatus).Content;
            if (st!=0)
            {
                TCSData.OeeResetCountHour = (short)DateTime.Now.Hour;
                TCSData.OeeResetCountMin = (short)DateTime.Now.Minute;
                TCSData.OeeResetCountSec = (short)DateTime.Now.Second;
                PLC1.Write(PLCResetCountStatus, 0);
            }
            //reset pb
            if (rPLCs[9]==1)
            {
                wPLCs[12] = 0;
            }
        }
        private void WritePLCDevice()
        {
            wPLCs[11] = TCSData.FlagPLCRunFinish;
            PLC1.Write("W200", wPLCs);
            //PC連線清零
            PLC1.Write(PLCConneted, 0);
        }
        #endregion
        #region --上報資料相關--
        /// <summary>
        /// 更新上報資料
        /// </summary>
        private void UpdateTheData()
        {
            //狀態
            //PLCStateMachineCommandStatus = myData.m;
            //myData.StateMachineRunStartFlag = _machineRunSetting;
            //PLCStateMachineCommandFlagStatus = myData.StateMachineRunStartFlag;
            //myData.StateMachineWorkType = 0;
            //PLC1.Write(PLCStateMachineCommand, PLCStateMachineCommandStatus);
            //PLC1.Write(PLCStateMachineCommandFlag, PLCStateMachineCommandFlagStatus);
            //基本設定
            //myData.BsettingLineNo = 1;
            //myData.BsettingStationID = 1;
            //人員ID
            //myData.StaffReadFlag = 0;
            //myData.SetStaffID(" ");
            //材料ID
            //myData.PartsData[0].SetPartsID(CurrentcountParts1.Barcode);
            //myData.PartsData[0].count = Convert.ToInt16( CurrentcountParts1.Count);
            //if (NumOfAxis==3)
            //{
            //    myData.PartsData[1].SetPartsID(CurrentcountSite.Barcode);
            //    myData.PartsData[1].count = Convert.ToInt16(CurrentcountSite.Count);
            //}
            //if (NumOfAxis == 4)
            //{
            //    myData.PartsData[1].SetPartsID(CurrentcountBottomCap.Barcode);
            //    myData.PartsData[1].count = Convert.ToInt16(CurrentcountBottomCap.Count);
            //}
            //製程ID
            //if (PLCResetMachineCountStatus>0)
            //{
            //    myData.SetResetMachineCountTime();
            //    PLCResetMachineCountStatus = 0;
            //    PLC1.Write(PLCResetMachineCountFlag, (Int16)0);
            //}
        }
        /// <summary>
        /// 取得伺服器參數
        /// </summary>
        /// <param name="startAddress">資料起始位置</param>
        private void Server2UpdateData(int startAddress)
        {
            int baseAddress = startAddress - 40000 + 1;
            //狀態
            //myData.StateMachineCommand= myServer.holdingRegisters[baseAddress + 0];
            //myData.StateResonOfMachineStop = myServer.holdingRegisters[baseAddress + 1];
        }
        /// <summary>
        /// 傳入機器上報參數
        /// </summary>
        /// <param name="startAddress">資料起始位置</param>
        /// <param name="usd">上報資料</param>
        private void MappingData(int startAddress, csDataCollection dc)
        {
            int baseAddress = startAddress - 40000 + 1;
            dc.EnableOperator = TCS.holdingRegisters[baseAddress + 7];
            dc.EnableWorkTypesOrStockOut = TCS.holdingRegisters[baseAddress + 8];
            dc.EnableParts = TCS.holdingRegisters[baseAddress + 9];
            dc.EnableMachineRun = TCS.holdingRegisters[baseAddress + 10];

            for (int i = 0; i < dc.PartIDs.Length; i++)
            {
                for (int j = 0; j < dc.PartIDs[i].ID.Length; j++)
                {
                    TCS.holdingRegisters[baseAddress + 75 + j + (45 * i)] = dc.PartIDs[i].ID[j];
                }
                dc.PartIDs[i].PartsCount = TCS.holdingRegisters[baseAddress + 115 + (45 * i)];
            }
            dc.OeeMachineRunHour = TCS.holdingRegisters[baseAddress + 578];
            dc.OeeMachineRunMin = TCS.holdingRegisters[baseAddress + 579];
            dc.OeeMachineRunSec = TCS.holdingRegisters[baseAddress + 580];

            TCS.holdingRegisters[baseAddress + 3] = dc.MachineStatus;
            TCS.holdingRegisters[baseAddress + 11] = dc.WorkID;
            //材料旗標
            if (TCS.holdingRegisters[baseAddress + 12] == 2000)
            {
                dc.FlagPLCPartsID = 0;
                TCS.holdingRegisters[baseAddress + 12] = dc.FlagPLCPartsID;
            }
            else
            {
                if (TCS.holdingRegisters[baseAddress + 12] != 2001 && dc.FlagPLCPartsID==2001)
                {
                    TCS.holdingRegisters[baseAddress + 12] = dc.FlagPLCPartsID;
                }
            }
            //設備完成旗標
            if (TCS.holdingRegisters[baseAddress + 13] == 3000)
            {
                dc.FlagPLCRunFinish = 0;
                TCS.holdingRegisters[baseAddress + 13] = dc.FlagPLCRunFinish;
            }
            else
            {
                if (TCS.holdingRegisters[baseAddress + 13] != 3001 && TCS.holdingRegisters[baseAddress + 13] != 3002 && dc.FlagPLCRunFinish==3001)
                {
                    TCS.holdingRegisters[baseAddress + 13] = dc.FlagPLCRunFinish;
                }
            }
            //製程ID變更旗標
            if (TCS.holdingRegisters[baseAddress + 14] == 7000)
            {
                dc.FlagMainIDChanged = 0;
                TCS.holdingRegisters[baseAddress + 14] = dc.FlagMainIDChanged;
            }
            else
            {
                if (TCS.holdingRegisters[baseAddress + 14] != 7001 && dc.FlagMainIDChanged==7001)
                {
                    TCS.holdingRegisters[baseAddress + 14] = dc.FlagMainIDChanged;
                }
            }
            //PLC換機種旗標
            if (TCS.holdingRegisters[baseAddress + 15] == 2000)
            {
                dc.FlagWorkTypeChanged = TCS.holdingRegisters[baseAddress + 15];
            }
            else
            {
                TCS.holdingRegisters[baseAddress + 15] = dc.FlagWorkTypeChanged;
            }
            TCS.holdingRegisters[baseAddress + 25] = dc.BasicSettingLineNo;
            TCS.holdingRegisters[baseAddress + 26] = dc.BasicSettingWorkType[0];
            TCS.holdingRegisters[baseAddress + 27] = dc.BasicSettingWorkType[1];
            TCS.holdingRegisters[baseAddress + 28] = dc.BasicSettingYear;
            TCS.holdingRegisters[baseAddress + 29] = dc.BasicSettingMonth;
            TCS.holdingRegisters[baseAddress + 30] = dc.BasicSettingDay;
            TCS.holdingRegisters[baseAddress + 31] = dc.BasicSettingHour;
            TCS.holdingRegisters[baseAddress + 32] = dc.BasicSettingMin;
            TCS.holdingRegisters[baseAddress + 33] = dc.BasicSettingSec;
            TCS.holdingRegisters[baseAddress + 34] = dc.BasicSettingEquipmentNo;



            TCS.holdingRegisters[baseAddress + 575] = dc.OeePowerOnHour;
            TCS.holdingRegisters[baseAddress + 576] = dc.OeePowerOnMin;
            TCS.holdingRegisters[baseAddress + 577] = dc.OeePowerOnSec;
            TCS.holdingRegisters[baseAddress + 581] = dc.OeeCycleTime;
            TCS.holdingRegisters[baseAddress + 582] = dc.OeeTotalCount;
            TCS.holdingRegisters[baseAddress + 583] = dc.OeeOkCount;
            TCS.holdingRegisters[baseAddress + 584] = dc.OeeNgCount;
            TCS.holdingRegisters[baseAddress + 585] = dc.OeeResetCountHour;
            TCS.holdingRegisters[baseAddress + 586] = dc.OeeResetCountMin;
            TCS.holdingRegisters[baseAddress + 587] = dc.OeeResetCountSec;

            TCS.holdingRegisters[baseAddress + 675] = dc.MainID[0];
            TCS.holdingRegisters[baseAddress + 676] = dc.MainID[1];
            TCS.holdingRegisters[baseAddress + 677] = dc.MainID[2];
            TCS.holdingRegisters[baseAddress + 678] = dc.MainID[3];
            TCS.holdingRegisters[baseAddress + 679] = dc.MainID[4];
            TCS.holdingRegisters[baseAddress + 680] = dc.MainID[5];
            TCS.holdingRegisters[baseAddress + 681] = dc.MainID[6];
            TCS.holdingRegisters[baseAddress + 682] = dc.MainID[7];
            TCS.holdingRegisters[baseAddress + 683] = dc.MainID[8];
            TCS.holdingRegisters[baseAddress + 684] = dc.MainID[9];
            TCS.holdingRegisters[baseAddress + 685] = dc.MainID[10];
            TCS.holdingRegisters[baseAddress + 686] = dc.MainID[11];

            for (int i = 0; i < dc.ScrewRows.Length; i++)
            {
                TCS.holdingRegisters[baseAddress + 1075 + (2 * i)] = dc.ScrewRows[i].TorqueValue;
                TCS.holdingRegisters[baseAddress + 1076 + (2 * i)] = dc.ScrewRows[i].TorqueJudgement;
                TCS.holdingRegisters[baseAddress + 1087 + (2 * i)] = dc.ScrewRows[i].AngleValue;
                TCS.holdingRegisters[baseAddress + 1088 + (2 * i)] = dc.ScrewRows[i].AngleJudgement;
            }

            TCS.holdingRegisters[baseAddress + 1374] = dc.OverrallJudgement;
        }
        //旗標清除
        private void ClearFlagPartsID()
        {
            if (TCSData.FlagMainIDChanged == 7000)
            {
                TCSData.FlagMainIDChanged = 0;
            }
            if (TCSData.FlagPLCPartsID == 2000)
            {
                TCSData.FlagPLCPartsID = 0;
            }
            if (TCSData.FlagPLCRunFinish == 3000)
            {
                TCSData.FlagPLCRunFinish = 0;
            }
            if (TCSData.FlagWorkTypeChanged == 2000)
            {
                TCSData.FlagWorkTypeChanged = 0;
            }
        }
        private void SetBasicSetting()
        {
            TCSData.BasicSettingLineNo = 2;
            TCSData.BasicSettingEquipmentNo = 1;
            MainIDCheckCode = PLC1.ReadString("D5018", 10).Content;
            string s = MainIDCheckCode.Remove(3);
            byte[] bs = Encoding.ASCII.GetBytes(s);
            if (bs.Length == 3)
            {
                int j = bs[1] << 8;
                j += bs[0];
                TCSData.BasicSettingWorkType[0] = (short)j;
                TCSData.BasicSettingWorkType[1] = bs[2];
            }
        }
        #endregion
        #region --其他--
        //測試用
        private void button1_Click(object sender, EventArgs e)
        {
            string checkText = tb_ItemBarcode.Text;
        }
        #endregion
    }
}
