using System;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
using StanleyDriver_RS232;
using HslCommunication.Profinet.Melsec;
using ClosedXML.Excel;
using EasyModbus;

namespace _18_203
{
    public partial class Form1 : Form
    {
        #region --定義初始化檔案--
        private string initailFilePath = @"D:\RP20-702\Resource\initailFile.xlsx";
        private string FilePathScrewCount = @"D:\RP20-702\Resource\ScrewCount.xlsx";
        private string filePath = @"D:\DataFile\";
        private string ScrewPartNo = "";
        private int BarcodeHeader=3;
        private int NumOfAxis=2;
        private string BodyBarcodeCheckCode = "";
        TextBox[] tbTorqueLimiteOfMaxArray;
        TextBox[] tbTorqueLimiteOfMinArray;
        TextBox[] tbAngleLimiteOfMaxArray;
        TextBox[] tbAngleLimiteOfMinArray;
        TextBox[] tbScrewDataShowTorqueValue;
        TextBox[] tbScrewDataShowTorqueResult;
        TextBox[] tbScrewDataShowAngleValue;
        TextBox[] tbScrewDataShowAngleResult;
        #endregion
        #region --定義PLC--
        private MelsecMcNet PLC1;
        private string PlcIpAddress="10.5.3.113";
        private int PlcPort=6000;
        //PLC->PC
        private int PLCBarcodeUseStatus = 0;
        private int PLCResetMachineCountStatus = 0;
        private int PLCStateMachineCommandStatus = 0;
        private int PLCStateMachineCommandFlagStatus = 0;
        //PC->PLC
        private string PLCConneted = "D2040";
        private string PLCBarcodeCheckStartFlag = "D2042";
        private string PLCBarcodeCheck = "D2044";
        private string PLCBarcodeHeaderNOK = "D2047";
        private string PLCOutOfScrew = "D2177";
        private string PLCSaveScrewDataFinish = "D2181";
        private string PLCResetMachineCountFlag = "D2182";
        private string PLCBarcodeUse = "D6000";
        private string PLCMachineStatus = "W0";
        private string PLCFinalResult = "W1";
        private string PLCCycleTime = "W2";
        private string PLCTotalCount = "W3";
        private string PLCOKCount = "W4";
        private string PLCNGCount = "W5";
        private string PLCMachineMode = "W6";
        private string PLCMachineErrorCode = "W7";
        private string PLCStateMachineCommand = "W8";
        private string PLCStateMachineCommandFlag = "W9";
        #endregion
        #region --定義序列埠--
        private SerialPort[] spAx = new SerialPort[7];
        #endregion
        #region --定義鎖付--
        private StanleyScrewData ssd;
        #endregion
        #region --定義機器相關--
        private string oldCheckBarcodeHeader = "";
        private string stringNoBarcode = "NoBarcode";
        private int NoBarcodeSeriealNo = 0, DataSaveCount = 0;
        private ObjectBarcodeAndCount CurrentcountScrew = new ObjectBarcodeAndCount();
        private ObjectBarcodeAndCount CurrentcountSite = new ObjectBarcodeAndCount();
        private ObjectBarcodeAndCount CurrentcountRS = new ObjectBarcodeAndCount();
        private ObjectBarcodeAndCount CurrentcountTopCap = new ObjectBarcodeAndCount();
        private ObjectBarcodeAndCount CurrentcountBottomCap = new ObjectBarcodeAndCount();
        #endregion
        #region --定義上報資料相關--
        private const int numOfParts = 1;//part 0:螺絲;part1:固定座
        private UpdateServerData myData;
        private ModbusServer myServer;
        private Int16 _machineRunSetting = 0;
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
            tbTorqueLimiteOfMaxArray = new TextBox[4] { tbTorqueLimiteOfMax_1, tbTorqueLimiteOfMax_2, tbTorqueLimiteOfMax_3, tbTorqueLimiteOfMax_4 };
            tbTorqueLimiteOfMinArray = new TextBox[4] { tbTorqueLimiteOfMin_1, tbTorqueLimiteOfMin_2, tbTorqueLimiteOfMin_3, tbTorqueLimiteOfMin_4 };
            tbAngleLimiteOfMaxArray = new TextBox[4] { tbAngleLimiteOfMax_1, tbAngleLimiteOfMax_2, tbAngleLimiteOfMax_3, tbAngleLimiteOfMax_4 };
            tbAngleLimiteOfMinArray = new TextBox[4] { tbAngleLimiteOfMin_1, tbAngleLimiteOfMin_2, tbAngleLimiteOfMin_3, tbAngleLimiteOfMin_4 };
            //初始檔案取得
            XLWorkbook wb = new XLWorkbook(initailFilePath);
            var ws = wb.Worksheet(1);
            //物件品號
            NumOfAxis = Convert.ToInt32(ws.Cell("A2").Value);
            PlcIpAddress = Convert.ToString(ws.Cell("B2").Value);
            PlcPort = Convert.ToInt32(ws.Cell("C2").Value);
            BarcodeHeader = Convert.ToInt32(ws.Cell("D2").Value);
            ScrewPartNo = Convert.ToString(ws.Cell("E2").Value);
            BodyBarcodeCheckCode = Convert.ToString(ws.Cell("F2").Value);
            oldCheckBarcodeHeader = BodyBarcodeCheckCode;
            //display物件品號
            tb_ScrewPartNo.Text = ScrewPartNo;
            tbBodyBarcodeCheckCode.Text = BodyBarcodeCheckCode;
            //取得各物件數量條碼
            ssd = new StanleyScrewData(NumOfAxis);
            GetObjectData();

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
            spAx[5].DataReceived += ax5_DataReceived;
            spAx[6].DataReceived += ax6_DataReceived;
            //建立PLC通訊
            PLC1 = new MelsecMcNet();
            PLC1.IpAddress = PlcIpAddress;
            PLC1.Port = PlcPort;
            PLC1.ConnectTimeOut = 2000;
            PLC1.NetworkNumber = 0x00;
            PLC1.NetworkStationNumber = 0x00;
            try
            {
                //Server
                myData = new UpdateServerData(1, numOfParts, NumOfAxis);
                myServer = new ModbusServer();
                myServer.Listen();
            
                //計時器開啟
                timer1.Enabled = true;
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = "FORM_LOAD:" + error.ToString();
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //myData.SaveTime();
            try
            {
                timer1.Enabled = false;
                myServer.StopListening();
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
        //固定條碼
        private void ax6_DataReceived(object sender, SerialDataReceivedEventArgs e)
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
                ssd.ScrewBarcode = CurrentcountScrew.Barcode;
                ssd.SiteBarcode = CurrentcountSite.Barcode;
                //檢查BARCODE
                BarCodeCheck();
                myData.SetProductBodyID(ss);
                tb_ItemBarcode.Text = ssd.ItemBarcode;
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = "固定槍讀取:" + error.ToString();
            }
        }
        //手持條碼
        private void ax5_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(800);
            try
            {
                Byte[] buffer = new Byte[1024];
                Int32 length = (sender as SerialPort).Read(buffer, 0, buffer.Length);
                Array.Resize(ref buffer, length);
                string ss = "";
                ss = System.Text.Encoding.ASCII.GetString(buffer);
                //解析條碼內容,分辨是何種條碼
                string[] newstring = ss.Split('*');
                if (newstring.Length == 5)
                {
                    //螺絲條碼
                    if (newstring[1] == ScrewPartNo)
                    {
                        ObjectBarcodeAndCount obj = new ObjectBarcodeAndCount();
                        obj.Barcode = ss;
                        obj.Count = Convert.ToInt32(newstring[3]);
                        obj.AddNewData(FilePathScrewCount);
                    }
                    //如果都不符合
                    if (newstring[1] != ScrewPartNo)
                    {
                        MessageBox.Show("無符合的品號");
                    }
                }
                else
                {
                    MessageBox.Show("輸入條碼異常:請確認條碼的正確性!!");
                }
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = "手持條碼:" + error.ToString();
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
                GetObjectBarcode();
                ////扣除數量
                //CurrentcountScrew.Count = CurrentcountScrew.Count - 1;//螺絲-1
                ////判斷沒螺絲就取新資料
                //if (CurrentcountScrew.Count <= 0)
                //{
                //    CurrentcountScrew.ReloadData(FilePathScrewCount);
                //    myData.PartsData[0].SetChangePartTime();
                //}
                DataSaveCount++;
                //show
                tb_Datetime_Ax4.Text = ssd._sd[3].ScrewDateTime;
                tb_Torque_Ax4.Text = ssd._sd[3].TorqueResult;
                tb_Angle_Ax4.Text = ssd._sd[3].AngleResult;
                tb_Overrall_Ax4.Text = ssd._sd[3].OverrallStatus;
                //傳到myDate
                myData.SetScrewData(4, ssd._sd[3]);
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = error.ToString();
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
                GetObjectBarcode();
                ////扣除數量
                //CurrentcountScrew.Count = CurrentcountScrew.Count - 1;//螺絲-1
                ////判斷沒螺絲就取新資料
                //if (CurrentcountScrew.Count <= 0)
                //{
                //    CurrentcountScrew.ReloadData(FilePathScrewCount);
                //    myData.PartsData[0].SetChangePartTime();
                //}
                DataSaveCount++;
                //show
                tb_Datetime_Ax3.Text = ssd._sd[2].ScrewDateTime;
                tb_Torque_Ax3.Text = ssd._sd[2].TorqueResult;
                tb_Angle_Ax3.Text = ssd._sd[2].AngleResult;
                tb_Overrall_Ax3.Text = ssd._sd[2].OverrallStatus;
                //傳到myDate
                myData.SetScrewData(3, ssd._sd[2]);
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = "AX3:"+error.ToString();
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
                GetObjectBarcode();
                ////扣除數量
                //CurrentcountScrew.Count = CurrentcountScrew.Count - 1;//螺絲-1
                ////判斷沒螺絲就取新資料
                //if (CurrentcountScrew.Count <= 0)
                //{
                //    CurrentcountScrew.ReloadData(FilePathScrewCount);
                //    myData.PartsData[0].SetChangePartTime();
                //}
                DataSaveCount++;
                //show
                tb_Datetime_Ax2.Text = ssd._sd[1].ScrewDateTime;
                tb_Torque_Ax2.Text = ssd._sd[1].TorqueResult;
                tb_Angle_Ax2.Text = ssd._sd[1].AngleResult;
                tb_Overrall_Ax2.Text = ssd._sd[1].OverrallStatus;
                //傳到myDate
                myData.SetScrewData(2, ssd._sd[1]);
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = "Ax2:"+error.ToString();
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
                GetObjectBarcode();
                ////扣除數量
                //CurrentcountScrew.Count--;//螺絲-1
                ////判斷沒螺絲就取新資料
                //if (CurrentcountScrew.Count <= 0)
                //{
                //    CurrentcountScrew.ReloadData(FilePathScrewCount);
                //    myData.PartsData[0].SetChangePartTime();
                //}
                DataSaveCount++;
                //show
                tb_Datetime_Ax1.Text = ssd._sd[0].ScrewDateTime;
                tb_Torque_Ax1.Text = ssd._sd[0].TorqueResult;
                tb_Angle_Ax1.Text = ssd._sd[0].AngleResult;
                tb_Overrall_Ax1.Text = ssd._sd[0].OverrallStatus;
                //傳到myDate
                myData.SetScrewData(1, ssd._sd[0]);
            }
            catch (Exception error)
            {
                tb_Message.Text = "";
                tb_Message.Text = "Ax1:"+error.ToString();
            }
        }
        #endregion
        #region --Timer--
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            //PC連線清零
            PLC1.Write(PLCConneted, 0);
            //檢查是否全部資料都存檔,供給不使用本體條碼用
            if (DataSaveCount >= NumOfAxis)
            {
                if (PLCBarcodeUseStatus == 0)
                {
                    ssd.SaveScrewData(filePath, ssd.ItemBarcode,NumOfAxis);
                }
                else
                {
                    string x = stringNoBarcode + NoBarcodeSeriealNo.ToString();
                    ssd.SaveScrewData(filePath, x,NumOfAxis);
                    NoBarcodeSeriealNo++;
                }
                CurrentcountScrew.SaveCountToFile(FilePathScrewCount);
                
                //PC->PLC寫檔完成
                PLC1.Write(PLCSaveScrewDataFinish, 1);
                DataSaveCount = 0;
                tb_ItemBarcode.Text = "";
            }
            //PLC讀取
            ReadPLCDevice();
            //Server
            Server2UpdateData(40000);//更新Modbus資料
            UpdateTheData();//更新上報資料
            UpdateData2Server(40000, myData);//更新Modbus資料
                                             //檢查物件數量
            CheckNoOfObject();
            ReflashDisplay();
            timer1.Enabled = true;
        }
        #endregion
        #region --重取材料--
        //品號存檔
        private void pb_SetPartNo_Click(object sender, EventArgs e)
        {
            ScrewPartNo = tb_ScrewPartNo.Text;
            BodyBarcodeCheckCode = tbBodyBarcodeCheckCode.Text;
            XLWorkbook wb = new XLWorkbook(initailFilePath);
            var ws = wb.Worksheet(1);
            ws.Cell("E2").Value = ScrewPartNo;
            ws.Cell("F2").Value = BodyBarcodeCheckCode;
            ws.Columns().AdjustToContents();
            wb.Save();
            ws.Dispose();
            wb.Dispose();
        }
        private void pb_LoadScrew_Click(object sender, EventArgs e)
        {
            CurrentcountScrew.LoadData(FilePathScrewCount);
        }
        private void pb_ReLoadScrew_Click(object sender, EventArgs e)
        {
            CurrentcountScrew.ReloadData(FilePathScrewCount);
            myData.PartsData[0].SetChangePartTime();
        }
        //檢查工件條碼
        private void BarCodeCheck()
        {
            BodyBarcodeCheckCode = PLC1.ReadString("D5018", 10).Content;
            tbBodyBarcodeCheckCode.Text = BodyBarcodeCheckCode;
            if (!oldCheckBarcodeHeader.Equals(BodyBarcodeCheckCode))
            {
                pb_SetPartNo_Click(null, null);
            }
            string itemBarcode = tb_ItemBarcode.Text;//取得現在的BARCODE
            //檢查條碼頭碼是否和現在使用TYPE相同
            try
            {
                string newss = itemBarcode.Remove(BodyBarcodeCheckCode.Length);//去掉工件條碼尾巴不用的字元
                if (newss.Equals(BodyBarcodeCheckCode))//比較頭碼和工作條碼
                {
                    //OK
                    PLC1.Write(PLCBarcodeCheck, (Int16)0);
                    oldCheckBarcodeHeader = BodyBarcodeCheckCode;
                }
                else
                {
                    //NOK
                    PLC1.Write(PLCBarcodeHeaderNOK, (Int16)1);
                }
            }
            catch (Exception e)
            {
                tb_Message.Text = e.ToString();
            }
            PLC1.Write(PLCBarcodeCheckStartFlag, (Int16)0);
        }
        //取得物件BARCODE
        private void GetObjectBarcode()
        {
            ssd.ScrewBarcode = CurrentcountScrew.Barcode;
            ssd.SiteBarcode = CurrentcountSite.Barcode;
            ssd.RsBarcode = CurrentcountRS.Barcode;
            ssd.TopCapBarcode = CurrentcountTopCap.Barcode;
            ssd.BottomCapBarcode = CurrentcountBottomCap.Barcode;
        }
        //取得物件資料
        private void GetObjectData()
        {
            CurrentcountScrew.LoadData(FilePathScrewCount);
            //if (NumOfAxis == 3)
            //{
            //    CurrentcountSite.LoadData(FilePathSiteCount);
            //}
            //CurrentcountRS.LoadData(FilePathRSCount);
            //CurrentcountTopCap.LoadData(FilePathTopCapCount);
            //if (NumOfAxis==4)
            //{
            //    CurrentcountBottomCap.LoadData(FilePathBottomCapCount);
            //}
        }
        #endregion
        #region --螢幕顯示--
        private void ReflashDisplay()
        {
            #region --主頁--
            //tb_ItemBarcode.Text = ssd.ItemBarcode;
            tb_ScrewBarcode.Text = CurrentcountScrew.Barcode;
            tb_CountOfScrew.Text = CurrentcountScrew.Count.ToString();
            #endregion
            #region --工程頁--
            tb_CountOfScrew_E.Text = CurrentcountScrew.Count.ToString();
            #endregion
            #region --上報資料頁--
            //狀態
            tbStateMachineCommand.Text = myData.StateMachineCommand.ToString();
            tbStateResonOfMachineStop.Text = myData.StateResonOfMachineStop.ToString();
            tbStateMachineRunStartFlag.Text = myData.StateMachineRunStartFlag.ToString();
            tbStateMachineStatus.Text = myData.StateMachineStatus.ToString();
            tbStateMachineMode.Text = myData.StateMachineMode.ToString();
            tbStateMachineErrorFlag.Text = myData.StateMachineErrorFlag.ToString();
            tbStateMachineWorkType.Text = myData.StateMachineWorkType.ToString();
            //基本設定
            tbBsettingLineNo.Text = myData.BsettingLineNo.ToString();
            tbBsettingWorkID.Text = myData.GetBSettingWorkID();
            tbBSettingYear.Text = myData.BSettingYear.ToString();
            tbBSettingMonth.Text = myData.BSettingMonth.ToString();
            tbBSettingDay.Text = myData.BSettingDay.ToString();
            tbBSettingHour.Text = myData.BSettingHour.ToString();
            tbBSettingMin.Text = myData.BSettingMin.ToString();
            tbBSettingSec.Text = myData.BSettingSec.ToString();
            tbBsettingStationID.Text = myData.BsettingStationID.ToString();
            //人員ID
            tbStaffReadFlag.Text = myData.StaffReadFlag.ToString();
            tbStaffID.Text = myData.GetStaffID();
            //材料ID
            //tbPartsData_1.Text = myData.PartsData[0].GetPartsID();
            //tbPartsCount_1.Text = myData.PartsData[0].count.ToString();
            //tbPartsChangeHour_1.Text = myData.PartsData[0].changePartHour.ToString();
            //tbPartsChangeMin_1.Text = myData.PartsData[0].changePartMin.ToString();
            //tbPartsChangeSec_1.Text = myData.PartsData[0].changePartSec.ToString();
            //tbPartsData_2.Text = myData.PartsData[1].GetPartsID();
            //tbPartsCount_2.Text = myData.PartsData[1].count.ToString();
            //tbPartsChangeHour_2.Text = myData.PartsData[1].changePartHour.ToString();
            //tbPartsChangeMin_2.Text = myData.PartsData[1].changePartMin.ToString();
            //tbPartsChangeSec_2.Text = myData.PartsData[1].changePartSec.ToString();
            //OEE
            tbOeePowerOnHours.Text = myData.OeePowerOnHours.ToString();
            tbOeePowerOnMins.Text = myData.OeePowerOnMins.ToString();
            tbOeePowerOnSecs.Text = myData.OeePowerOnSecs.ToString();
            tbOeeMachineRunHours.Text = myData.OeeMachineRunHours.ToString();
            tbOeeMachineRunMins.Text = myData.OeeMachineRunMins.ToString();
            tbOeeMachineRunSecs.Text = myData.OeeMachineRunSecs.ToString();
            tbOeeCycleTime.Text = myData.OeeCycleTime.ToString();
            tbOeeOutputCount.Text = myData.OeeOutputCount.ToString();
            tbOeeOKPartsCount.Text = myData.OeeOKPartsCount.ToString();
            tbOeeNGPartsCount.Text = myData.OeeNGPartsCount.ToString();
            tbOeeResetCountHour.Text = myData.OeeResetCountHour.ToString();
            tbOeeResetCountMin.Text = myData.OeeResetCountMin.ToString();
            tbOeeResetCountSec.Text = myData.OeeResetCountSec.ToString();
            //製程ID
            tbProductBodyID.Text = myData.GetProductBodyID();
            //設備參數及測試數據
            for (int i = 0; i < NumOfAxis; i++)
            {
                tbTorqueLimiteOfMaxArray[i].Text = myData.ParaScrewSetting[i].TorqueLimiteOfMax.ToString();
                tbTorqueLimiteOfMinArray[i].Text = myData.ParaScrewSetting[i].TorqueLimiteOfMin.ToString();
                tbAngleLimiteOfMaxArray[i].Text = myData.ParaScrewSetting[i].AngleLimiteOfMax.ToString();
                tbAngleLimiteOfMinArray[i].Text = myData.ParaScrewSetting[i].AngleLimiteOfMin.ToString();
                tbScrewDataShowTorqueValue[i].Text= myData.DataScrewData[i].TorqueValue.ToString();
                tbScrewDataShowTorqueResult[i].Text= myData.DataScrewData[i].TorqueResult.ToString();
                tbScrewDataShowAngleValue[i].Text = myData.DataScrewData[i].AngleValue.ToString();
                tbScrewDataShowAngleResult[i].Text = myData.DataScrewData[i].AngleResult.ToString();
            }
            #endregion
        }
        #endregion
        #region --手動輸入條碼--
        private void pbInputBarcode7110_Click(object sender, EventArgs e)
        {
            tbInputMessage.Text = " ";
            string ss = tbInputBarcode.Text;
            //解析條碼內容,分辨是何種條碼
            string[] newstring = ss.Split('*');
            if (newstring.Length == 5)
            {
                //螺絲條碼
                if (newstring[1] == ScrewPartNo)
                {
                    ObjectBarcodeAndCount obj = new ObjectBarcodeAndCount();
                    obj.Barcode = ss;
                    obj.Count = Convert.ToInt32(newstring[3]);
                    obj.AddNewData(FilePathScrewCount);
                }
                //如果都不符合
                if (newstring[1] != ScrewPartNo)
                {
                    MessageBox.Show("無符合的品號");
                }
            }
            else
            {
                tbInputMessage.Text = "輸入條碼異常:請確認條碼的正確性!!";
            }
        }
        #endregion
        #region --PLC相關--
        //讀取PLC
        private void ReadPLCDevice()
        {
            //取得PLC狀態
            PLCBarcodeUseStatus = PLC1.ReadInt16(PLCBarcodeUse).Content;
            PLCResetMachineCountStatus = PLC1.ReadInt16(PLCResetMachineCountFlag).Content;
            //上報資料
            myData.StateMachineStatus = PLC1.ReadInt16(PLCMachineStatus).Content;
            string ss = BodyBarcodeCheckCode;
            myData.SetBSettingWorkID(ss);
            myData.FinalResult = PLC1.ReadInt16(PLCFinalResult).Content;
            myData.OeeCycleTime = PLC1.ReadInt16(PLCCycleTime).Content;
            myData.OeeOutputCount = PLC1.ReadInt16(PLCTotalCount).Content;
            myData.OeeOKPartsCount = PLC1.ReadInt16(PLCOKCount).Content;
            myData.OeeNGPartsCount = PLC1.ReadInt16(PLCNGCount).Content;
            myData.StateMachineMode = PLC1.ReadInt16(PLCMachineMode).Content;
            myData.StateMachineErrorFlag = PLC1.ReadInt16(PLCMachineErrorCode).Content;
        }
        //檢查物件數量傳送給PLC
        private void CheckNoOfObject()
        {
            //螺絲
            if (CurrentcountScrew.Count <= 0)
            {
                PLC1.Write(PLCOutOfScrew, (Int16)0);
            }
            else
            {
                PLC1.Write(PLCOutOfScrew, (Int16)1);
            }
        }
        #endregion
        #region --上報資料相關--
        /// <summary>
        /// 更新上報資料
        /// </summary>
        private void UpdateTheData()
        {
            //狀態
            PLCStateMachineCommandStatus = myData.StateMachineCommand;
            myData.StateMachineRunStartFlag = _machineRunSetting;
            PLCStateMachineCommandFlagStatus = myData.StateMachineRunStartFlag;
            myData.StateMachineWorkType = 0;
            PLC1.Write(PLCStateMachineCommand, PLCStateMachineCommandStatus);
            PLC1.Write(PLCStateMachineCommandFlag, PLCStateMachineCommandFlagStatus);
            //基本設定
            myData.BsettingLineNo = 1;
            myData.BsettingStationID = 1;
            //人員ID
            myData.StaffReadFlag = 0;
            myData.SetStaffID(" ");
            //材料ID
            myData.PartsData[0].SetPartsID(CurrentcountScrew.Barcode);
            myData.PartsData[0].count = Convert.ToInt16( CurrentcountScrew.Count);
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
            if (PLCResetMachineCountStatus>0)
            {
                myData.SetResetMachineCountTime();
                PLCResetMachineCountStatus = 0;
                PLC1.Write(PLCResetMachineCountFlag, 0);
            }
        }
        /// <summary>
        /// 取得伺服器參數
        /// </summary>
        /// <param name="startAddress">資料起始位置</param>
        private void Server2UpdateData(int startAddress)
        {
            int baseAddress = startAddress - 40000 + 1;
            //狀態
            myData.StateMachineCommand= myServer.holdingRegisters[baseAddress + 0];
            myData.StateResonOfMachineStop = myServer.holdingRegisters[baseAddress + 1];
        }
        /// <summary>
        /// 傳入機器上報參數
        /// </summary>
        /// <param name="startAddress">資料起始位置</param>
        /// <param name="usd">上報資料</param>
        private void UpdateData2Server(int startAddress, UpdateServerData usd)
        {
            int baseAddress = startAddress - 40000 + 1;
            myServer.holdingRegisters[baseAddress + 2] = usd.StateMachineRunStartFlag;
            myServer.holdingRegisters[baseAddress + 3] = usd.StateMachineStatus;
            myServer.holdingRegisters[baseAddress + 4] = usd.StateMachineMode;
            myServer.holdingRegisters[baseAddress + 5] = usd.StateMachineErrorFlag;
            myServer.holdingRegisters[baseAddress + 6] = usd.StateMachineWorkType;
            //基本設定
            myServer.holdingRegisters[baseAddress + 25] = usd.BsettingLineNo;
            myServer.holdingRegisters[baseAddress + 26] = usd.BsettingWorkID[0];
            myServer.holdingRegisters[baseAddress + 27] = usd.BsettingWorkID[1];
            myServer.holdingRegisters[baseAddress + 28] = usd.BSettingYear;
            myServer.holdingRegisters[baseAddress + 29] = usd.BSettingMonth;
            myServer.holdingRegisters[baseAddress + 30] = usd.BSettingDay;
            myServer.holdingRegisters[baseAddress + 31] = usd.BSettingHour;
            myServer.holdingRegisters[baseAddress + 32] = usd.BSettingMin;
            myServer.holdingRegisters[baseAddress + 33] = usd.BSettingSec;
            myServer.holdingRegisters[baseAddress + 34] = usd.BsettingStationID;
            //人員ID
            myServer.holdingRegisters[baseAddress + 50] = usd.StaffReadFlag;
            for (int i = 0; i < usd.StaffID.Length; i++)
            {
                myServer.holdingRegisters[baseAddress + 51+i] = usd.StaffID[i];
            }
            //材料ID
            for (int i = 0; i <35; i++)
            {
                //myServer.holdingRegisters[baseAddress + 75 + i] = usd.PartsData[0].id[i];
                //myServer.holdingRegisters[baseAddress + 120 + i] = usd.PartsData[1].id[i];
            }
            //myServer.holdingRegisters[baseAddress + 110] = usd.PartsData[0].count;
            //myServer.holdingRegisters[baseAddress + 111] = usd.PartsData[0].changePartHour;
            //myServer.holdingRegisters[baseAddress + 112] = usd.PartsData[0].changePartMin;
            //myServer.holdingRegisters[baseAddress + 113] = usd.PartsData[0].changePartSec;
            //myServer.holdingRegisters[baseAddress + 155] = usd.PartsData[1].count;
            //myServer.holdingRegisters[baseAddress + 156] = usd.PartsData[1].changePartHour;
            //myServer.holdingRegisters[baseAddress + 157] = usd.PartsData[1].changePartMin;
            //myServer.holdingRegisters[baseAddress + 158] = usd.PartsData[1].changePartSec;
            //OEE
            myServer.holdingRegisters[baseAddress + 575] = (short)usd.OeePowerOnHours;
            myServer.holdingRegisters[baseAddress + 576] = usd.OeePowerOnMins;
            myServer.holdingRegisters[baseAddress + 577] = usd.OeePowerOnSecs;
            myServer.holdingRegisters[baseAddress + 578] = (short)usd.OeeMachineRunHours;
            myServer.holdingRegisters[baseAddress + 579] = usd.OeeMachineRunMins;
            myServer.holdingRegisters[baseAddress + 580] = usd.OeeMachineRunSecs;
            myServer.holdingRegisters[baseAddress + 581] = usd.OeeCycleTime;
            myServer.holdingRegisters[baseAddress + 582] = usd.OeeOutputCount;
            myServer.holdingRegisters[baseAddress + 583] = usd.OeeOKPartsCount;
            myServer.holdingRegisters[baseAddress + 584] = usd.OeeNGPartsCount;
            myServer.holdingRegisters[baseAddress + 585] = usd.OeeResetCountHour;
            myServer.holdingRegisters[baseAddress + 586] = usd.OeeResetCountMin;
            myServer.holdingRegisters[baseAddress + 587] = usd.OeeResetCountSec;
            //製程ID
            for (int i = 0; i < usd.ProductBodyID.Length; i++)
            {
                myServer.holdingRegisters[baseAddress + 675+i] = usd.ProductBodyID[i];
            }
            //設備參數及測試數據
            for (int i = 0; i < NumOfAxis; i++)
            {
                myServer.holdingRegisters[baseAddress + 775 + (i * 2)] = usd.DataScrewData[i].TorqueValue;
                myServer.holdingRegisters[baseAddress + 776 + (i * 2)] = usd.DataScrewData[i].TorqueResult;
                myServer.holdingRegisters[baseAddress + 787 + (i * 2)] = Convert.ToInt16( usd.DataScrewData[i].AngleValue);
                myServer.holdingRegisters[baseAddress + 788 + (i * 2)] = Convert.ToInt16( usd.DataScrewData[i].AngleValue);
                myServer.holdingRegisters[baseAddress + 1075+(i*2)] = usd.DataScrewData[i].TorqueValue;
                myServer.holdingRegisters[baseAddress + 1076+(i*2)] = usd.DataScrewData[i].TorqueResult;
                myServer.holdingRegisters[baseAddress + 1087 + (i * 2)] = Convert.ToInt16( usd.DataScrewData[i].AngleValue);
                myServer.holdingRegisters[baseAddress + 1088 + (i * 2)] = Convert.ToInt16( usd.DataScrewData[i].AngleResult);
            }
            myServer.holdingRegisters[baseAddress + 1374] = usd.FinalResult;
        }
        // 設定生產旗標
        private void pbSetMachineRunSetting_Click(object sender, EventArgs e)
        {
            try
            {
                short oldvalue = _machineRunSetting;
                short newvalue = Convert.ToInt16(tbMachineRunSetting.Text);
                if (newvalue == 0 || newvalue == 1)
                {
                    _machineRunSetting = newvalue;
                }
                else
                {
                    _machineRunSetting = oldvalue;
                    MessageBox.Show("錯誤!!請輸入0:Off或1:On");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("錯誤!!輸入非數值資料!!"); 
            }
        }
        #endregion
        #region --其他--
        private void pb_GetBarcodeHeader_Click(object sender, EventArgs e)
        {
        }
        //測試用
        private void button1_Click(object sender, EventArgs e)
        {
            BarCodeCheck();
        }

        #endregion
    }
}
