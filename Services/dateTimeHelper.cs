using System.Data;
using System.Globalization;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CPSAppData.Services
{
    public class dateTimeHelper
    {
        string[] monthNames = { "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม" };
        CultureInfo cultureEN = new CultureInfo("en-US");
        CultureInfo cultureTH = new CultureInfo("th-TH");
        public DateTime convertDateValueXLSToDateTime(Double datexlsvalue)
        {
            return DateTime.FromOADate(datexlsvalue);
        }
        public Double convertDateTimeToDateValueXLS(DateTime datevalue)
        {
            return datevalue.ToOADate();
        }
        public string doGetDateThaiFromUIToPDF(string datadate)
        {
            string date_str = datadate.Replace("/", "");
            string y_ = date_str.Substring(4);
            string m_ = date_str.Substring(2, 2);
            string d_ = date_str.Substring(0, 2);

            DateTime date_arg = DateTime.Parse(string.Format("{0}-{1}-{2}", y_, m_, d_));            
            int date = Convert.ToInt16(date_arg.Day);
            int month = Convert.ToInt16(date_arg.Month);
            int year = Convert.ToInt16(date_arg.Year);

            int yearchk = year - 543;
            if (yearchk < 1950)
            {
                year = year + 543;
            }

            string namethai = monthNames[month - 1];
            return string.Format("{0} {1} {2}", date, namethai, year);
        }
        public string doGetDateUIToDB(string datadate)
        {
            CultureInfo culture = CultureInfo.GetCultureInfo("en-US");
            string date_str = datadate.Replace("/", "");
            string y_ = date_str.Substring(4);
            string m_ = date_str.Substring(2, 2);
            string d_ = date_str.Substring(0, 2);

            DateTime date_arg = DateTime.Parse(string.Format("{0}-{1}-{2}", y_, m_, d_), culture);
            int date = Convert.ToInt16(date_arg.Day);
            int month = Convert.ToInt16(date_arg.Month);
            int year = Convert.ToInt16(date_arg.Year);


            return string.Format("{0}{1}{2}", year, doGetQNumByZeroStr(2, month), doGetQNumByZeroStr(2, date));
        }
        public string doGetDateThaiFromDBToPDF(string datadate)
        {            
            if (datadate.Length == 8)
            {
                string y_ = datadate.Substring(0, 4);
                string m_ = datadate.Substring(4, 2);
                string d_ = datadate.Substring(6);
                int month = Convert.ToInt16(m_);
                int year = Convert.ToInt16(y_);

                string year_ = (year + 543).ToString();
                string namethai = monthNames[month - 1];
                return string.Format("{0} {1} {2}", d_, namethai, year_);
            }
            else
            {
                return "-";
            }
        }
        public string doGetShortDateFromDBToPDF(string datadate)
        {
            if (datadate.Length == 8)
            {
                string y_ = datadate.Substring(0, 4);
                string m_ = datadate.Substring(4, 2);
                string d_ = datadate.Substring(6);
                int month = Convert.ToInt16(m_);
                int year = Convert.ToInt16(y_);
                int day = Convert.ToInt16(d_);

                return string.Format("{0}/{1}/{2}", doGetQNumByZeroStr(2,day), doGetQNumByZeroStr(2, month), year);
            }
            else
            {
                return "-";
            }
        }
        public string doGetShortDateTHFromDBToPDF(string datadate)
        {
            if (datadate.Length == 8)
            {
                string y_ = datadate.Substring(0, 4);
                string m_ = datadate.Substring(4, 2);
                string d_ = datadate.Substring(6);
                int month = Convert.ToInt16(m_);
                int year = Convert.ToInt16(y_);
                int day = Convert.ToInt16(d_);
                string year_ = (year + 543).ToString();
                return string.Format("{0}/{1}/{2}", doGetQNumByZeroStr(2, day), doGetQNumByZeroStr(2, month), year_);
            }
            else
            {
                return "-";
            }
        }
        public string ConverDateTODBStr(DataRow rowdatedata,string colname)
        {
            string[] format = new string[] { "d/M/yyyy","d/MM/yyyy" , "dd/M/yyyy", "dd/MM/yyyy" };
            int index = 3;
            if (rowdatedata != null && colname != null)
            {            // DateTime dateValue;
                string? str_date = Convert.ToString(rowdatedata[colname] is DBNull ? "" : rowdatedata[colname]);
                string? str_datex = str_date;
                if (!string.IsNullOrEmpty(str_date))
                {
                    if(str_date.Contains(":"))
                    {
                        int indexlast = str_date.IndexOf(" ");
                        str_date = str_date.Substring(0, indexlast);
                    }

                    DateTime dateValue;
                    if (str_date.Contains("/"))
                    {
                        switch (str_date.Length)
                        {
                            case 10:
                                index = 3;
                                break;
                            case 9:
                                index = 1;
                                break;
                            case 8:
                                index = 0;
                                break;
                        }
                    }
                    else
                    {
                        index = 99;
                    }

                    try
                    {
                        if(index == 99)
                        {
                            Double dateOA = Convert.ToDouble(str_date);
                            DateTime datedata_excel = convertDateValueXLSToDateTime(dateOA);
                            string str_datedb = convertToDateStr(datedata_excel);
                            return str_datedb;
                        }
                        else
                        {
                            bool success = DateTime.TryParseExact(str_date, format[index], cultureEN, DateTimeStyles.None, out dateValue);
                            if (index == 1 && (!success))
                            {
                                index = 2;
                                DateTime.TryParseExact(str_date, format[index], cultureEN, DateTimeStyles.None, out dateValue);
                            }
                            Double dateOA = convertDateTimeToDateValueXLS(dateValue);
                            DateTime datedata_excel = convertDateValueXLSToDateTime(dateOA);
                            string str_datedb = convertToDateStr(datedata_excel);
                            return str_datedb;
                        }
                       
                    }
                    catch
                    {
                        return string.Format("10000101[{0}]", str_datex);
                    }
                }
            }
            return string.Empty;
        }
        private string convertToDateStr(DateTime datedata_excel)
        {            
            int yearint = datedata_excel.Year;
            if(yearint > 2500) // พศ
            {
                yearint = yearint - 543;
            }
           
            string year = yearint.ToString();
            string month = doGetQNumByZeroStr(2,datedata_excel.Month);
            string day = doGetQNumByZeroStr(2,datedata_excel.Day);
            return string.Format("{0}{1}{2}",year,month,day);
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