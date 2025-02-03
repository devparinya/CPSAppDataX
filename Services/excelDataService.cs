using OfficeOpenXml;
using System.Data;

namespace CPSAppData.Services
{
    public class excelDataService
    {
        //GET HEADER EXCEL FILE
        public Dictionary<string, string>? doGetColumnHDFromExcel(string pathFileName,bool ismaster)
        {
            Dictionary<string, string>? headers = new Dictionary<string, string>();
            string[] s_heder = new string[] { "LedNumber", "WorkNo", "Maxmonth" };
            // อ่านไฟล์ Excel
            FileInfo fileInfo = new FileInfo(pathFileName);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(fileInfo)) 
            { 
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 
                // สมมติว่าใช้ sheet แรก 
                // ตรวจสอบว่ามีข้อมูลใน sheet หรือไม่
                if (worksheet.Dimension == null) 
                { 
                    return headers; 
                } // ดึงค่า Header จากแถวแรกและแปลงเป็น
                
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++) 
                { 
                    string header = worksheet.Cells[1, col].Text;
                    if (ismaster)
                    {
                        if (!s_heder.Contains(header))
                        {
                            headers.Add(header, header);
                        }
                    }
                    else
                    {
                        headers.Add(header, header);
                    }
                } // แสดงผลค่า header ใน Dictionary               
            }
            return headers;
        }

        //Convert Excel Data to Datatable 
        public DataTable ReadExcelToDataTable(string filePath)
        {
            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // เพิ่มคอลัมน์ใน DataTable
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text);
                }

                // เพิ่มแถวใน DataTable
                for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var row = dataTable.NewRow();
                    foreach (var cell in worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column])
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    dataTable.Rows.Add(row);
                }
            }

            return dataTable;
        }
        public DataTable excelToDataTable(string filePath, bool hasHeader = true)
        {
            var dt = new DataTable();
            var fi = new FileInfo(filePath);
            // Check if the file exists
            if (!fi.Exists)
                throw new Exception("File " + filePath + " Does Not Exists");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var xlPackage = new ExcelPackage(fi);
            // get the first worksheet in the workbook
            List<string> sheetname = getSheetNameEXCEL(filePath);
            var worksheet = xlPackage.Workbook.Worksheets[sheetname[0]];
            try
            {
                dt = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].ToDataTable(c =>
                {
                    c.FirstRowIsColumnNames = true;
                    c.AlwaysAllowNull = true;
                });
            }
            catch(Exception ex)
            {
                if (ex.Message.Contains("SkipNumberOfRowsStart"))
                {
                    MessageBox.Show(string.Format("ไฟล์ {0} \r\n ไม่มีข้อมูลกรุณาตรวจสอบ",Path.GetFileName(filePath)),"คำเตือน",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
                else
                {
                    return new DataTable();
                }
            }

            return dt;
        }
       
        private List<string> getSheetNameEXCEL(string pathdata)
        {
            List<string> sheetName = new List<string>();
            FileInfo fileInfo = new FileInfo(pathdata); 
            using (ExcelPackage package = new ExcelPackage(fileInfo)) 
            { 
                foreach (var worksheet in package.Workbook.Worksheets) 
                {
                    sheetName.Add(worksheet.Name); 
                } 
            }
            return sheetName;
        }
        public void ExportToExcel(DataTable dataTable, string[] captionth,string[] captioneng, string excelFilePath = "")
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DataCPS"); 
                // เพิ่มชื่อคอลัมน์
                for (int i = 0; i < dataTable.Columns.Count; i++) 
                {
                    worksheet.Cells[1, i + 1].Value = captionth[i]; 
                } // เพิ่มแถวข้อมูล
                for (int i = 0; i < dataTable.Rows.Count; i++) 
                { 
                    for (int j = 0; j < dataTable.Columns.Count; j++) 
                    { 
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j].ToString(); 
                    } 
                } // บันทึกไฟล์ Excel
                if (captioneng.Length > 0)
                {
                    var worksheet_cap = package.Workbook.Worksheets.Add("Caption");
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        worksheet_cap.Cells[1, i + 1].Value = captioneng[i];
                    } // เพิ่มแถวข้อมูล
                }
                FileInfo fi = new FileInfo(excelFilePath); 
                package.SaveAs(fi); 
            }
        }
        public void ExportToExcelAll(DataTable dataTable,string excelFilePath = "")
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("datacalulate");
                // เพิ่มชื่อคอลัมน์
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                } // เพิ่มแถวข้อมูล
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                } // บันทึกไฟล์ Excel
               
                FileInfo fi = new FileInfo(excelFilePath);
                package.SaveAs(fi);
            }
        }
        public void ExportExcelResult(DataTable dtresult,string resultpath,string typename)
        {
            if (!Path.Exists(resultpath)) return; 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DataCPS");
                // เพิ่มชื่อคอลัมน์
                for (int i = 0; i < dtresult.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dtresult.Columns[i].ColumnName;
                } // เพิ่มแถวข้อมูล
                for (int i = 0; i < dtresult.Rows.Count; i++)
                {
                    for (int j = 0; j < dtresult.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dtresult.Rows[i][j].ToString();
                    }
                }
                string fullpath = doGetFullNameResultFile(resultpath, typename);
                FileInfo fi = new FileInfo(fullpath);
                package.SaveAs(fi);
            }
        }
        private string doGetFullNameResultFile(string pathfile,string typename)
        {
            string fullfilename = string.Empty;
            int year_ = DateTime.Now.Year;
            int month_ = DateTime.Now.Month;
            int day_ = DateTime.Now.Day;
            int h_ = DateTime.Now.Hour;
            int m_ = DateTime.Now.Minute;
            int s_ = DateTime.Now.Second;
            int ms_ = DateTime.Now.Millisecond;
            string filename = string.Format("result_{0}_{1}{2}{3}{4}{5}{6}{7}.xlsx",typename,year_, doGetQNumByZeroStr(2,month_), doGetQNumByZeroStr(2,day_), doGetQNumByZeroStr(2,h_), doGetQNumByZeroStr(2,m_), doGetQNumByZeroStr(2,s_), doGetQNumByZeroStr(2,ms_));
            fullfilename = Path.Combine(pathfile,filename);
            return fullfilename;
        }
        private string doGetQNumByZeroStr(int qlen, int qnum)
        {
            string q_str = qnum.ToString();

            if (qlen == 4)
            {
                int length = q_str.Length;
                if (length == 4) return q_str;
                if (length == 3) return "0" + q_str;
                if (length == 2) return "00" + q_str;
                if (length == 1) return "000" + q_str;
            }
            if (qlen == 3)
            {
                int length = q_str.Length;
                if (length == 3) return q_str;
                if (length == 2) return "0" + q_str;
                if (length == 1) return "00" + q_str;
            }
            if (qlen == 2)
            {
                int length = q_str.Length;
                if (length == 2) return q_str;
                if (length == 1) return "0" + q_str;
            }
            return q_str;
        }

    }
}
    