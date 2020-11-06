using System;
using System.IO;
using ClosedXML.Excel;

namespace StanleyDriver_RS232
{
    interface IstanleyScrewData
    {
        void checkFileExist(string filePath, string barcode);
        void SaveScrewData(string filePath, string barcode,int numOfAxis);
        void GetRs232ScrewData(byte[] rs232data,int axisno);
    }

    public class StanleyScrewData:IstanleyScrewData
    {
        private string fileEnd = @".xlsx";
        private string fileYear = DateTime.Now.ToString("yyyy");
        private string fileMonth = DateTime.Now.ToString("MM");
        private string fileDate = DateTime.Now.ToString("MMdd");
        private string _itembarcode, _parts1barcode;
        
        public string ItemBarcode { get { return _itembarcode; } set { _itembarcode = value; } }            //工件條碼
        public string Parts1Barcode { get { return _parts1barcode; } set { _parts1barcode = value; } }         //工件條碼
        //public string SiteBarcode { get { return _sitebarcode; } set { _sitebarcode = value; } }         //固定承座條碼
        //public string RsBarcode { get { return _rsbarcode; } set { _rsbarcode = value; } }                  //RS條碼
        //public string TopCapBarcode { get { return _topcapbarcode; } set { _topcapbarcode = value; } }                  //上蓋板條碼
        //public string BottomCapBarcode { get { return _bottomcapbarcode; } set { _bottomcapbarcode = value; } }         //下蓋板條碼
        public ScrewData[] _sd;

        /// <summary>
        ///建構子 
        /// </summary>
        /// <param name="inNumberOfAxis">軸數量</param>
        public StanleyScrewData(int inNumberOfAxis)
        {
            _sd = new ScrewData[inNumberOfAxis];
            for (int i = 0; i < inNumberOfAxis; i++)
            {
                _sd[i] = new ScrewData();
            }
        }

        //檢查檔案是否存在,不存在就建立新檔
        public void checkFileExist(string filePath,string filename)
        {
            string fullfileName = filePath + fileYear + @"\" + fileMonth + @"\" + fileDate + @"\" + filename.TrimEnd() + fileEnd;

            bool fileExist = File.Exists(fullfileName);
            if (!fileExist)
            {
                XLWorkbook wb = new XLWorkbook();
                wb.AddWorksheet("sheet1");
                var ws = wb.Worksheet(1);
                ws.Cell("A1").Value = "軸號碼";
                ws.Cell("B1").Value = "JOB號碼";
                ws.Cell("C1").Value = "扭力輸出";
                ws.Cell("D1").Value = "扭力判定,A=OK,L=Low,H=High";
                ws.Cell("E1").Value = "角度輸出";
                ws.Cell("F1").Value = "角度判定,A=OK,L=Low,H=High";
                ws.Cell("G1").Value = "總合判定,A=OK,R=NOK";
                ws.Cell("H1").Value = "鎖付日期時間";
                ws.Cell("I1").Value = "材料條碼";
                ws.Columns().AdjustToContents();
                wb.SaveAs(fullfileName);
                ws.Dispose();
                wb.Dispose();
            }
        }
        /// <summary>
        /// 待所有軸都載入後,儲存鎖付檔案
        /// </summary>
        /// <param name="filePath">檔案根目錄</param>
        /// <param name="filename">檔案名稱,即本體條碼</param>
        /// <param name="inNumberOfAxis">軸數量</param>
        public void SaveScrewData(string filePath,string filename,int inNumberOfAxis)
        {
            checkFileExist(filePath,filename);
            string fullfileName = filePath + fileYear + @"\" + fileMonth + @"\" + fileDate + @"\" + filename.TrimEnd() + fileEnd;
            XLWorkbook wb = new XLWorkbook(fullfileName);
            var ws = wb.Worksheet(1);
            for (int i = 0; i < inNumberOfAxis; i++)
            {
                ws.Cell(i + 2, 1).Value = _sd[i].SpindleNumber;
                ws.Cell(i + 2, 2).Value = _sd[i].JobNumber;
                ws.Cell(i + 2, 3).Value = _sd[i].TorqueResult;
                ws.Cell(i + 2, 4).Value = _sd[i].TorqueStatus;
                ws.Cell(i + 2, 5).Value = _sd[i].AngleResult;
                ws.Cell(i + 2, 6).Value = _sd[i].AngleStatus;
                ws.Cell(i + 2, 7).Value = _sd[i].OverrallStatus;
                ws.Cell(i + 2, 8).Value = _sd[i].ScrewDateTime;
                ws.Cell(i + 2, 9).Value = Parts1Barcode;
                //ws.Cell(i + 2, 10).Value = SiteBarcode;
            }
            ws.Columns().AdjustToContents();
            wb.SaveAs(fullfileName);
            ws.Dispose();
            wb.Dispose();
        }

        //DRIVER RS232取得鎖付資料解析
        public void GetRs232ScrewData(byte[] rs232data,int axisno)
        {
            string str = System.Text.Encoding.Default.GetString(rs232data);
            char c = ',';
            string[] substrings = str.Split(c);
            _sd[axisno-1].SpindleNumber = axisno.ToString();
            _sd[axisno-1].JobNumber = substrings[1];
            _sd[axisno-1].TorqueResult = substrings[2];
            _sd[axisno-1].TorqueStatus = substrings[3];
            _sd[axisno-1].AngleResult = substrings[4];
            _sd[axisno-1].AngleStatus = substrings[5];
            _sd[axisno-1].OverrallStatus = substrings[6];
        }
    }
    public class ScrewData
    {
        private string _spindlenumber, _jobnumber, _torqueresult, _torquestatus, _angleresult, _anglestatus, _overrallstatus, _screwdatetime;

        public string JobNumber { get { return _spindlenumber; } set { _spindlenumber = value; } }          //軸號碼
        public string SpindleNumber { get { return _jobnumber; } set { _jobnumber = value; } }               //JOB號碼
        public string TorqueResult { get { return _torqueresult; } set { _torqueresult = value; } }         //扭力輸出
        public string TorqueStatus { get { return _torquestatus; } set { _torquestatus = value; } }         //扭力判定A=OK,L=Low,H=High
        public string AngleResult { get { return _angleresult; } set { _angleresult = value; } }            //角度輸出
        public string AngleStatus { get { return _anglestatus; } set { _anglestatus = value; } }            //角度判定A=OK,L=Low,H=High
        public string OverrallStatus { get { return _overrallstatus; } set { _overrallstatus = value; } }   //總合判定A=OK,R=NOK
        public string ScrewDateTime { get { return _screwdatetime; } set { _screwdatetime = value; } }          //鎖付的時間
    }
}
