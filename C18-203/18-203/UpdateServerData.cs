using System;
using System.Collections.Generic;
using System.Windows.Forms;
using StanleyDriver_RS232;
using ClosedXML.Excel;
using System.IO;

namespace _18_203
{
    public class UpdateServerData
    {
        #region --一般define--
        private string recordFilePath = @"D:\RP20-702\Resource\record.txt";
        private string backupFilePath= @"D:\RP20-702\Resource\backuprecord.txt";
        private int _noOfOeeTime;
        private int PowerOnTotalTime,MachineRunTotalTime;
        private Timer t1 = new Timer();
        #endregion
        #region --Create--
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="noOfOeeTime">計算開機時間及運轉時間的數量</param>
        /// <param name="numOfParts">材料數</param>
        /// <param name="numOfAxis">軸數</param>
        public UpdateServerData(int noOfOeeTime, int numOfParts, int numOfAxis)
        {
            #region --建構物件--
            //檢查檔案
            //如果沒檔案就建立新檔
            _noOfOeeTime = noOfOeeTime;
            bool _fileExist = File.Exists(recordFilePath);
            if (_fileExist)
            {
                StreamReader sr = new StreamReader(recordFilePath);
                PowerOnTotalTime = Convert.ToInt32(sr.ReadLine());
                MachineRunTotalTime = Convert.ToInt32(sr.ReadLine());
                sr.Dispose();
            }
            else
            {
                string rat= File.ReadAllText(backupFilePath);
                File.WriteAllText(recordFilePath, rat);
                StreamReader sr = new StreamReader(recordFilePath);
                PowerOnTotalTime = Convert.ToInt32(sr.ReadLine());
                MachineRunTotalTime = Convert.ToInt32(sr.ReadLine());
                sr.Dispose();
            }
            for (int i = 0; i < numOfParts; i++)
            {
                parts axis = new parts();
                _partData.Add(axis);
            }
            for (int i = 0; i < numOfAxis; i++)
            {
                ScrewSetting setting = new ScrewSetting();
                _paraScrewSetting.Add(setting);
                ServerScrewData data = new ServerScrewData();
                _dataScrewData.Add(data);
            }
            #endregion
            #region --OEE時間--
            t1.Interval = 1000;
            t1.Tick += T1_Tick;
            t1.Start();
            #endregion
        }
        ~UpdateServerData()
        {
            string str = PowerOnTotalTime.ToString() + Environment.NewLine + MachineRunTotalTime.ToString();
            File.WriteAllText(recordFilePath, str);
        }
        #endregion
        #region --OEE計時--
        private void T1_Tick(object sender, EventArgs e)
        {
            t1.Stop();
            setDateTime();
            PowerOnTotalTime++;
            if (_stateMachineStatus==1)
            {
                MachineRunTotalTime++;
            }
            SetOeePowerOnTime();
            SetOeeMachineRunTime();
            t1.Start();
        }
        #endregion
        #region --狀態--
        //--Define--
        private Int16 _stateMachineCommand = 0;
        private Int16 _stateResonOfMachineStop = 0;
        private Int16 _stateMachineRunStartFlag = 0;
        private Int16 _stateMachineStatus = 0;
        private Int16 _stateMachineMode = 0;
        private Int16 _stateMachineErrorFlag = 0;
        private Int16 _stateMachineWorkType = 0;
        //--Get,Set--
        /// <summary>
        /// 生產許可=e.g.0:none;1:Stop
        /// </summary>
        public Int16 StateMachineCommand
        {
            get { return _stateMachineCommand; }
            set { _stateMachineCommand = value; }
        }
        /// <summary>
        /// 生產禁止原因=e.g.0:無異常;1:資料格式錯誤
        /// </summary>
        public Int16 StateResonOfMachineStop
        {
            get { return _stateResonOfMachineStop; }
            set { _stateResonOfMachineStop = value; }
        }
        /// <summary>
        /// 生產許可旗標=e.g.0: Off ; 1: On
        /// </summary>
        public Int16 StateMachineRunStartFlag
        {
            get { return _stateMachineRunStartFlag; }
            set{_stateMachineRunStartFlag = value;}
        }
        /// <summary>
        /// 設備動作=e.g.0:待機;1:開始測試;2:完成
        /// </summary>
        public Int16 StateMachineStatus
        {
            get { return _stateMachineStatus; }
            set { _stateMachineStatus = value; }
        }
        /// <summary>
        /// 設備模式=e.g.0:自動;1:半自動/試作;2:手動
        /// </summary>
        public Int16 StateMachineMode
        {
            get { return _stateMachineMode; }
            set { _stateMachineMode = value; }
        }
        /// <summary>
        /// 設備異常原因=e.g.0:無;1:氣壓不足;
        /// </summary>
        public Int16 StateMachineErrorFlag
        {
            get { return _stateMachineErrorFlag; }
            set { _stateMachineErrorFlag = value; }
        }
        /// <summary>
        /// 標準品生產?=e.g.0:生產品;1:標準品;
        /// </summary>
        public Int16 StateMachineWorkType
        {
            get { return _stateMachineWorkType; }
            set { _stateMachineWorkType = value; }
        }
        #endregion
        #region --基本設定--
        //--Define--
        private Int16 _bsettingLineNo = 0;
        private Int16[] _bsettingWorkID = new Int16[2] { 0, 0 };
        private Int16 _bsettingYear = 0;
        private Int16 _bsettingMonth = 0;
        private Int16 _bsettingDay = 0;
        private Int16 _bsettingHour = 0;
        private Int16 _bsettingMin = 0;
        private Int16 _bsettingSec = 0;
        private Int16 _bsettingStationID = 0;
        //--Get,Set--
        /// <summary>
        /// LINE No.
        /// </summary>
        public Int16 BsettingLineNo
        {
            get { return _bsettingLineNo; }
            set { _bsettingLineNo = value; }
        }
        /// <summary>
        /// 機種=197
        /// </summary>
        public Int16[] BsettingWorkID
        {
            get { return _bsettingWorkID; }
        }
        /// <summary>
        /// 設備即時-年
        /// </summary>
        public Int16 BSettingYear { get { return _bsettingYear; } }
        /// <summary>
        /// 設備即時-月
        /// </summary>
        public Int16 BSettingMonth { get { return _bsettingMonth; } }
        /// <summary>
        /// 設備即時-日
        /// </summary>
        public Int16 BSettingDay { get { return _bsettingDay; } }
        /// <summary>
        /// 設備即時-時
        /// </summary>
        public Int16 BSettingHour { get { return _bsettingHour; } }
        /// <summary>
        /// 設備即時-分
        /// </summary>
        public Int16 BSettingMin { get { return _bsettingMin; } }
        /// <summary>
        /// 設備即時-秒
        /// </summary>
        public Int16 BSettingSec { get { return _bsettingSec; } }
        /// <summary>
        /// 設備工位
        /// </summary>
        public Int16 BsettingStationID
        {
            get { return _bsettingStationID; }
            set { _bsettingStationID = value; }
        }
        //--Funtion--
        /// <summary>
        /// 取得當前時間
        /// </summary>
        private void setDateTime()
        {
            _bsettingYear = Convert.ToInt16(DateTime.Now.Year);
            _bsettingMonth = Convert.ToInt16(DateTime.Now.Month);
            _bsettingDay = Convert.ToInt16(DateTime.Now.Day);
            _bsettingHour = Convert.ToInt16(DateTime.Now.Hour);
            _bsettingMin = Convert.ToInt16(DateTime.Now.Minute);
            _bsettingSec = Convert.ToInt16(DateTime.Now.Second);
        }
        /// <summary>
        /// Set機種
        /// </summary>
        /// <param name="bodyBarcodeHeader">本體條碼Part碼</param>
        public void SetBSettingWorkID(string bodyBarcodeHeader)
        {
            if (bodyBarcodeHeader!=null)
            {
                string ss = bodyBarcodeHeader;
                string[] nss = new string[2];
                nss[0] = ss.Remove(2);
                nss[1] = ss.Remove(0, 2);
                _bsettingWorkID[0] = Convert.ToInt16(nss[0]);
                _bsettingWorkID[1] = Convert.ToInt16(nss[1]);
            }
        }
        /// <summary>
        /// Get string 機種
        /// </summary>
        /// <returns>機種字串</returns>
        public string GetBSettingWorkID()
        {
            string ss = _bsettingWorkID[0].ToString() + _bsettingWorkID[1].ToString();
            return ss;
        }
        #endregion
        #region --人員ID--
        //--Define--
        private byte[] m_staffID = new byte[10];
        private Int16 _staffReadFlag = 0;
        private Int16[] _staffID = new Int16[5]{0,0,0,0,0};
        //--Get,Set--
        /// <summary>
        /// 讀取人員ID旗標=0:Off;1:On
        /// </summary>
        public Int16 StaffReadFlag
        {
            get { return _staffReadFlag; }
            set { _staffReadFlag = value; }
        }
        /// <summary>
        /// 人員ID
        /// </summary>
        public Int16[] StaffID
        {
            get { return _staffID; }
        }
        //--Funtion--
        /// <summary>
        /// 設定人員ID條碼
        /// </summary>
        /// <param name="instaffId">人員ID條碼</param>
        public void SetStaffID(string instaffId)
        {
            string ss = instaffId.Trim();
            int ivalue = ss.Length;
            if (ivalue <= 10)
            {
                m_staffID = System.Text.Encoding.ASCII.GetBytes(ss);
                Array.Resize(ref m_staffID, 10);
                for (int i = 0; i < 5; i++)
                {
                    _staffID[i] = Convert.ToInt16(m_staffID[i * 2] << 8);
                    _staffID[i] = Convert.ToInt16(_staffID[i] + m_staffID[i * 2 + 1]);
                }
            }
            else
            {
                MessageBox.Show("人員ID碼長度有誤,請確認!!");
            }
        }
        /// <summary>
        /// 取得人員ID條碼
        /// </summary>
        /// <returns>人員ID條碼</returns>
        public string GetStaffID()
        {
            string ss = System.Text.Encoding.ASCII.GetString(m_staffID);
            return ss;
        }
        #endregion
        #region --材料ID--
        //--Define--
        //private byte[] m_id = new byte[80];
        private List<parts> _partData=new List<parts>();
        //--Get,Set--
        /// <summary>
        /// 材料ID
        /// </summary>
        public List<parts> PartsData
        {
            get { return _partData; }
            set { _partData = value; }
        }
        #endregion
        #region --稼動率overall equipment effectiveness--
        //--Define--
        private UInt16 _oeePowerOnHours=0;
        private Int16 _oeePowerOnMins=0;
        private Int16 _oeePowerOnSecs=0;
        private UInt16 _oeeMachineRunHours=0;
        private Int16 _oeeMachineRunMins=0;
        private Int16 _oeeMachineRunSecs=0;
        private Int16 _oeeCycleTime=0;
        private Int16 _oeeOutputCount=0;
        private Int16 _oeeOKPartsCount=0;
        private Int16 _oeeNGPartsCount=0;
        private Int16 _oeeResetCountHour=0;
        private Int16 _oeeResetCountMin=0;
        private Int16 _oeeResetCountSec=0;
        //--Get,Set--
        /// <summary>
        /// 電源投入總時間- 時
        /// </summary>
        public UInt16 OeePowerOnHours{get { return _oeePowerOnHours; }}
        /// <summary>
        /// 電源投入總時間- 分
        /// </summary>
        public Int16 OeePowerOnMins{get { return _oeePowerOnMins; }}
        /// <summary>
        /// 電源投入總時間- 秒
        /// </summary>
        public Int16 OeePowerOnSecs{get { return _oeePowerOnSecs; }}
        /// <summary>
        /// 機台運轉總時間- 時
        /// </summary>
        public UInt16 OeeMachineRunHours{get { return _oeeMachineRunHours; }}
        /// <summary>
        /// 機台運轉總時間- 分
        /// </summary>
        public Int16 OeeMachineRunMins{get { return _oeeMachineRunMins; }}
        /// <summary>
        /// 機台運轉總時間- 秒
        /// </summary>
        public Int16 OeeMachineRunSecs{get { return _oeeMachineRunSecs; }}
        /// <summary>
        /// 循環時間-秒
        /// </summary>
        public Int16 OeeCycleTime
        {
            get { return _oeeCycleTime; }
            set { _oeeCycleTime = value; }
        }
        /// <summary>
        /// 生產數
        /// </summary>
        public Int16 OeeOutputCount
        {
            get { return _oeeOutputCount; }
            set { _oeeOutputCount = value; }
        }
        /// <summary>
        ///良品數
        /// </summary>
        public Int16 OeeOKPartsCount
        {
            get { return _oeeOKPartsCount; }
            set { _oeeOKPartsCount = value; }
        }
        /// <summary>
        ///不良品數
        /// </summary>
        public Int16 OeeNGPartsCount
        {
            get { return _oeeNGPartsCount; }
            set { _oeeNGPartsCount = value; }
        }
        /// <summary>
        /// 生產數歸零的設定時間- 時
        /// </summary>
        public Int16 OeeResetCountHour
        {
            get { return _oeeResetCountHour; }
            set { _oeeResetCountHour = value; }
        }
        /// <summary>
        /// 生產數歸零的設定時間- 分
        /// </summary>
        public Int16 OeeResetCountMin
        {
            get { return _oeeResetCountMin; }
            set { _oeeResetCountMin = value; }
        }
        /// <summary>
        /// 生產數歸零的設定時間- 秒
        /// </summary>
        public Int16 OeeResetCountSec
        {
            get { return _oeeResetCountSec; }
            set { _oeeResetCountSec = value; }
        }
        //--Funtion
        /// <summary>
        /// 設定電源投入時間
        /// </summary>
        private void SetOeePowerOnTime()
        {
            TimeSpan ts1 = new TimeSpan(0, 0, 0, PowerOnTotalTime);
            int ts1Days = ts1.Days;
            int ts1Hours = ts1.Hours;
            int ts1Mins = ts1.Minutes;
            int ts1Secs = ts1.Seconds;
            _oeePowerOnHours = Convert.ToUInt16(ts1Days * 24 + ts1Hours);
            _oeePowerOnMins = Convert.ToInt16(ts1Mins);
            _oeePowerOnSecs = Convert.ToInt16(ts1Secs);
        }
        /// <summary>
        /// 設定電源投入時間
        /// </summary>
        private void SetOeeMachineRunTime()
        {
            TimeSpan ts1 = new TimeSpan(0, 0, 0, MachineRunTotalTime);
            int ts1Days = ts1.Days;
            int ts1Hours = ts1.Hours;
            int ts1Mins = ts1.Minutes;
            int ts1Secs = ts1.Seconds;
            _oeeMachineRunHours = Convert.ToUInt16(ts1Days * 24 + ts1Hours);
            _oeeMachineRunMins = Convert.ToInt16(ts1Mins);
            _oeeMachineRunSecs = Convert.ToInt16(ts1Secs);
        }
        /// <summary>
        /// 設定重置生產數時間
        /// </summary>
        public void SetResetMachineCountTime()
        {
            _oeeResetCountHour = (short)DateTime.Now.Hour;
            _oeeResetCountMin = (short)DateTime.Now.Minute;
            _oeeResetCountSec = (short)DateTime.Now.Second;
        }
        public void SaveTime()
        {
            string str = PowerOnTotalTime.ToString() + Environment.NewLine + MachineRunTotalTime.ToString();
            File.WriteAllText(recordFilePath, str);
        }
        #endregion
        #region --製程ID--
        //--Define--
        private byte[] m_productBodyID = new byte[20];
        private Int16[] _productBodyID=new Int16[10] { 0,0,0,0,0,0,0,0,0,0};
        //--Get,Set--
        /// <summary>
        /// 半成品ID
        /// </summary>
        public Int16[] ProductBodyID
        {
            get { return _productBodyID; }
        }
        //Funtion
        /// <summary>
        /// Set製程ID
        /// </summary>
        /// <param name="inBodyID">本體條碼</param>
        public void SetProductBodyID(string inBodyID)
        {
            string ss= inBodyID.Trim();
            int ivalue = ss.Length;
            if (ivalue<=20)
            {
                m_productBodyID = System.Text.Encoding.ASCII.GetBytes(ss);
                Array.Resize(ref m_productBodyID, 20);
                for (int i = 0; i < 10; i++)
                {
                    _productBodyID[i] = Convert.ToInt16(m_productBodyID[i*2] << 8);
                    _productBodyID[i] = Convert.ToInt16(_productBodyID[i] + m_productBodyID[i*2+1]);
                }
            }
            else
            {
                MessageBox.Show("本體條碼長度有誤,請確認!! 字串為:"+inBodyID+"...");
            }
        }
        /// <summary>
        /// 取得本體ID
        /// </summary>
        /// <returns>本體ID</returns>
        public string GetProductBodyID()
        {
            string ss = System.Text.Encoding.ASCII.GetString(m_productBodyID);
            return ss;
        }
        #endregion
        #region --設備參數--
        //--Define--
        private List<ScrewSetting> _paraScrewSetting = new List<ScrewSetting>();
        //--Get,Set--
        /// <summary>
        /// 設備參數
        /// </summary>
        public List<ScrewSetting> ParaScrewSetting
        {
            get { return _paraScrewSetting; }
            set { _paraScrewSetting = value; }
        }
        #endregion
        #region --測試數據--
        //--Define--
        private List<ServerScrewData> _dataScrewData = new List<ServerScrewData>();
        private Int16 _finalResult=0;
        //--Get,Set--
        /// <summary>
        /// 測試數據
        /// </summary>
        public List<ServerScrewData> DataScrewData
        {
            get { return _dataScrewData; }
            set { _dataScrewData = value; }
        }
        /// <summary>
        /// 產出結果=0:OK;1:NG
        /// </summary>
        public Int16 FinalResult
        {
            get { return _finalResult; }
            set { _finalResult = value; }
        }
        //--Funtion
        /// <summary>
        /// Set鎖付資料
        /// </summary>
        /// <param name="axisNo">軸編號</param>
        /// <param name="ssd">鎖付資料</param>
        public void SetScrewData(int axisNo,ScrewData sd)
        {
            ScrewData sda = sd;
            int i = axisNo - 1;
            //*******************************************************
            double ddt = Convert.ToDouble(sda.TorqueResult);
            ddt = ddt * 100;
            short stt = Convert.ToInt16( ddt);
            _dataScrewData[i].TorqueValue = stt;
            //*******************************************************
            double dda = Convert.ToDouble(sda.AngleResult);
            dda = dda * 10;
            ushort sta = Convert.ToUInt16(dda);
            _dataScrewData[i].AngleValue = sta;
            //*******************************************************
            string sst = sda.TorqueStatus.Trim();
            if (sst=="A")
            {
                _dataScrewData[i].TorqueResult = 0;
            }
            else
            {
                _dataScrewData[i].TorqueResult = 1;
            }
            //*******************************************************
            string ssa = sda.AngleStatus.Trim();
            if (ssa == "A")
            {
                _dataScrewData[i].AngleResult = 0;
            }
            else
            {
                _dataScrewData[i].AngleResult = 1;
            }
        }
        #endregion
    }
    #region --Class--
    public class parts
    {
        public parts(){}
        private byte[] m_id = new byte[70];
        public Int16[] id=new Int16[35];
        public Int16 count = 0;
        public Int16 changePartHour = 0;
        public Int16 changePartMin = 0;
        public Int16 changePartSec = 0;
        #region --Funtion--
        /// <summary>
        /// 設定材料ID
        /// </summary>
        /// <param name="inPartID">材料ID</param>
        public void SetPartsID(string inPartID)
        {
            string ss = inPartID.Trim();
            int ivalue = ss.Length;
            if (ivalue <= 70 & ivalue > 0)
            {
                m_id = System.Text.Encoding.ASCII.GetBytes(ss);
                Array.Resize(ref m_id, 70);
                for (int i = 0; i < 35; i++)
                {
                    id[i] = Convert.ToInt16(m_id[i * 2] << 8);
                    id[i] = Convert.ToInt16(id[i] + m_id[i * 2 + 1]);
                }
            }
        }
        /// <summary>
        /// 取得材料ID
        /// </summary>
        /// <returns>材料ID</returns>
        public string GetPartsID()
        {
            string ss = System.Text.Encoding.ASCII.GetString(m_id);
            return ss;
        }
        /// <summary>
        /// 設定換材料時間
        /// </summary>
        public void SetChangePartTime()
        {
            changePartHour = (short)DateTime.Now.Hour;
            changePartMin = (short)DateTime.Now.Minute;
            changePartSec = (short)DateTime.Now.Second;
        }
        #endregion
    }
    public class ScrewSetting
    {
        public Int16 TorqueLimiteOfMax = 0;
        public Int16 TorqueLimiteOfMin = 0;
        public Int16 AngleLimiteOfMax = 0;
        public Int16 AngleLimiteOfMin = 0;
    }
    public class ServerScrewData
    {
        public Int16 TorqueValue = 0;
        public Int16 TorqueResult = 0;
        public UInt16 AngleValue = 0;
        public Int16 AngleResult = 0;
    }
    #endregion
}
