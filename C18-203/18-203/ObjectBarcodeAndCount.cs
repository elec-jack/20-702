using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _18_203
{
    class ObjectBarcodeAndCount
    {
        private string _barcode;
        private int _count;

        public string Barcode { get { return _barcode; } set {_barcode=value; } }
        public int Count { get { return _count; } set { _count = value; } }

        public ObjectBarcodeAndCount() { }

        //檢查檔案是否存在,不存在就建立新檔
        public void checkFileExist(string p)
        {
            bool fileExist = File.Exists(p);
            if (!fileExist)
            {
                XLWorkbook wb = new XLWorkbook();
                wb.AddWorksheet("sheet1");
                wb.SaveAs(p);
                wb.Dispose();
            }
        }
        //數量存檔
        public void SaveCountToFile(string fp)
        {
            XLWorkbook wb = new XLWorkbook(fp);
            var ws = wb.Worksheet(1);
            ws.Cell(1, 2).Value = Count;
            ws.Columns().AdjustToContents();
            wb.Save();
            ws.Dispose();
            wb.Dispose();
        }
        //新增資料
        public void AddNewData(string ff)
        {
            checkFileExist(ff);
            IXLWorkbook wb = new XLWorkbook(ff);
            var ws = wb.Worksheet(1);
            try
            {
                int lastrow = ws.LastRowUsed().RowNumber();
                ws.Cell(lastrow + 1, 1).Value = Barcode;
                ws.Cell(lastrow + 1, 2).Value = Count;
                ws.Columns().AdjustToContents();
                wb.Save();
            }
            catch (Exception)
            {
                ws.Cell(1, 1).Value = Barcode;
                ws.Cell(1, 2).Value = Count;
                ws.Columns().AdjustToContents();
                wb.Save();
            }
            ws.Dispose();
            wb.Dispose();
        }
        //取得資料
        public void LoadData(string fullfilepath)
        {
            checkFileExist(fullfilepath);
            IXLWorkbook wb = new XLWorkbook(fullfilepath);
            var ws = wb.Worksheet(1);
            try
            {
                Barcode = ws.Cell(1, 1).Value.ToString();
                Count = Convert.ToInt32(ws.Cell(1, 2).Value);
            }
            catch (Exception)
            {
                MessageBox.Show("無資料,請確認補料條碼已掃描!!");
            }
            ws.Dispose();
            wb.Dispose();
        }
        //數量不足,更新資料
        public void ReloadData(string pp)
        {
            checkFileExist(pp);
            IXLWorkbook wb = new XLWorkbook(pp);
            var ws = wb.Worksheet(1);
            try
            {
                ws.Row(1).Delete();
                Barcode = ws.Cell(1, 1).Value.ToString();
                Count = Convert.ToInt32(ws.Cell(1, 2).Value);
            }
            catch (Exception)
            {
                Barcode = "";
                Count = 0;
                MessageBox.Show("無資料,請確認補料條碼已掃描!!");
            }
            wb.Save();
            ws.Dispose();
            wb.Dispose();
        }
    }
}
