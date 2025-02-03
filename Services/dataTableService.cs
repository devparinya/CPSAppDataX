using System.Data;
using System.Security.Cryptography;

namespace CpsDataApp.Services
{
    public class dataTableService
    {

        public DataTable doCreateDataCPSMasterTemplate()
        {
            DataTable dataTemplate = new DataTable();
            dataTemplate.Columns.Add("CadeID", typeof(string));
            dataTemplate.Columns.Add("CadeStatus", typeof(string));
            //CARD 
            dataTemplate.Columns.Add("CardNo", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate", typeof(string));//ชำระครั้งล่าสุด            
            //DataCustomer
            dataTemplate.Columns.Add("CustomerName", typeof(string)); //ชื่อ สกุล
            dataTemplate.Columns.Add("CustomerTel", typeof(string));//รหัสบัตรประชาชน
            dataTemplate.Columns.Add("CustomerID", typeof(string));//รหัสบัตรประชาชน            
            dataTemplate.Columns.Add("LegalStatus", typeof(string));//สถานะคดี
            dataTemplate.Columns.Add("BlackNo", typeof(string));//เลขคดีดำ
            dataTemplate.Columns.Add("RedNo", typeof(string));//เลขคดีแดง
            dataTemplate.Columns.Add("JudgeDate", typeof(string));//วันพิพากษา        
            dataTemplate.Columns.Add("CourtName", typeof(string));
            dataTemplate.Columns.Add("LegalExecRemark", typeof(string)); //หมายเหตุบังคับคดี
            dataTemplate.Columns.Add("LegalExecDate", typeof(string)); //วันที่ยึด อายัดทรัพย์
            //Collector Name
            dataTemplate.Columns.Add("CollectorName", typeof(string)); // ชื่อ collcctor
            dataTemplate.Columns.Add("CollectorTeam", typeof(string)); //ทีม
            dataTemplate.Columns.Add("CollectorTel", typeof(string)); //โทร
            return dataTemplate;

        }
        public DataTable doCreateMasterCPSDataDuplicate()
        {
            DataTable dataTemplate = new DataTable();
            dataTemplate.Columns.Add("CaseID", typeof(string));
            dataTemplate.Columns.Add("CardStatus", typeof(string));
            //DataCustomer
            dataTemplate.Columns.Add("CustomerID", typeof(string));//รหัสบัตรประชาชน
            dataTemplate.Columns.Add("CustomerName", typeof(string)); //ชื่อ สกุล            
            dataTemplate.Columns.Add("CustomerTel", typeof(string));//
            dataTemplate.Columns.Add("BlackNo", typeof(string));//เลขคดีดำ
            dataTemplate.Columns.Add("RedNo", typeof(string));//เลขคดีแดง
            dataTemplate.Columns.Add("JudgeDate", typeof(string));//วันพิพากษา        
            dataTemplate.Columns.Add("CourtName", typeof(string));
            dataTemplate.Columns.Add("LegalExecRemark", typeof(string)); //หมายเหตุบังคับคดี
            dataTemplate.Columns.Add("LegalExecDate", typeof(string)); //วันที่ยึด อายัดทรัพย์
            dataTemplate.Columns.Add("LegalStatus", typeof(string));//สถานะคดี

            //CARD 1
            dataTemplate.Columns.Add("CardNo1", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt1", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt1", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt1", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt1", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate1", typeof(string));//ชำระครั้งล่าสุด
            //CARD 2
            dataTemplate.Columns.Add("CardNo2", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt2", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt2", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt2", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt2", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate2", typeof(string));//ชำระครั้งล่าสุด
            //CARD 3
            dataTemplate.Columns.Add("CardNo3", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt3", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt3", typeof(double));//ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt3", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt3", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate3", typeof(string));//ชำระครั้งล่าสุด
             //CARD 4        
            dataTemplate.Columns.Add("CardNo4", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt4", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt4", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt4", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt4", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate4", typeof(string));//ชำระครั้งล่าสุด
            //CARD 5
            dataTemplate.Columns.Add("CardNo5", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt5", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt5", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt5", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt5", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate5", typeof(string));//ชำระครั้งล่าสุด
            //CARD 6
            dataTemplate.Columns.Add("CardNo6", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt6", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt6", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt6", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt6", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate6", typeof(string));//ชำระครั้งล่าสุด
            
            //Collector Name
            dataTemplate.Columns.Add("CollectorName", typeof(string)); // ชื่อ collcctor
            dataTemplate.Columns.Add("CollectorTeam", typeof(string)); //ทีม
            dataTemplate.Columns.Add("CollectorTel", typeof(string)); //โทร
            return dataTemplate;
        }
        public DataTable doCreateMasterCPSDataAfterCalculate()
        {
            DataTable dataTemplate = new DataTable();
            dataTemplate.Columns.Add("CaseID", typeof(string));
            dataTemplate.Columns.Add("CardStatus", typeof(string));
            //DATA             
            dataTemplate.Columns.Add("CustomerID", typeof(string));//รหัสบัตรประชาชน
            dataTemplate.Columns.Add("CustomerName", typeof(string)); //ชื่อ สกุล
            dataTemplate.Columns.Add("CustomerTel", typeof(string));//รหัสบัตรประชาชน
            dataTemplate.Columns.Add("LegalStatus", typeof(string));//สถานะคดี
            dataTemplate.Columns.Add("BlackNo", typeof(string));//เลขคดีดำ
            dataTemplate.Columns.Add("RedNo", typeof(string));//เลขคดีแดง
            dataTemplate.Columns.Add("JudgeDate", typeof(string));//วันพิพากษา        
            dataTemplate.Columns.Add("CourtName", typeof(string));
            dataTemplate.Columns.Add("LedNumber", typeof(string));// ลำดับกรม
            dataTemplate.Columns.Add("WorkNo", typeof(string)); //เลขชุดงาน 
            //CARD 1
            dataTemplate.Columns.Add("CardNo1", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt1", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt1", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt1", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt1", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate1", typeof(string));//ชำระครั้งล่าสุด
            dataTemplate.Columns.Add("AccCloseAmnt1", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt1", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt1", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt1", typeof(double));
            //CARD 2
            dataTemplate.Columns.Add("CardNo2", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt2", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt2", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt2", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt2", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate2", typeof(string));//ชำระครั้งล่าสุด
            dataTemplate.Columns.Add("AccCloseAmnt2", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt2", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt2", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt2", typeof(double));
            //CARD 3
            dataTemplate.Columns.Add("CardNo3", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt3", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt3", typeof(double));//ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt3", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt3", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate3", typeof(string));//ชำระครั้งล่าสุด
            dataTemplate.Columns.Add("AccCloseAmnt3", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt3", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt3", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt3", typeof(double));
            //CARD 4        
            dataTemplate.Columns.Add("CardNo4", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt4", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt4", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt4", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt4", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate4", typeof(string));//ชำระครั้งล่าสุด
            dataTemplate.Columns.Add("AccCloseAmnt4", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt4", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt4", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt4", typeof(double));
            //CARD 5
            dataTemplate.Columns.Add("CardNo5", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt5", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt5", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt5", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt5", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate5", typeof(string));//ชำระครั้งล่าสุด
            dataTemplate.Columns.Add("AccCloseAmnt5", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt5", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt5", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt5", typeof(double));
            //CARD 6
            dataTemplate.Columns.Add("CardNo6", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt6", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt6", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt6", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt6", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate6", typeof(string));//ชำระครั้งล่าสุด
            dataTemplate.Columns.Add("AccCloseAmnt6", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt6", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt6", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt6", typeof(double));
            //DataCustomer
           
            dataTemplate.Columns.Add("LegalExecRemark", typeof(string)); //หมายเหตุบังคับคดี
            dataTemplate.Columns.Add("LegalExecDate", typeof(string)); //วันที่ยึด อายัดทรัพย์
            //Collector Name
            dataTemplate.Columns.Add("CollectorName", typeof(string)); // ชื่อ collcctor
            dataTemplate.Columns.Add("CollectorTeam", typeof(string)); //ทีม
            dataTemplate.Columns.Add("CollectorTel", typeof(string)); //โทร
            dataTemplate.Columns.Add("MaxMonth", typeof(string));
            return dataTemplate;
        }
        public DataTable doCreateMasterCPSDataTemplate()
        {
            DataTable dataTemplate = new DataTable();
            dataTemplate.Columns.Add("CaseID", typeof(string));
            dataTemplate.Columns.Add("CardStatus", typeof(string));
            //CARD 1
            dataTemplate.Columns.Add("CardNo1", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt1", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt1", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt1", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt1", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate1", typeof(string));//ชำระครั้งล่าสุด
            //CARD 2
            dataTemplate.Columns.Add("CardNo2", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt2", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt2", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt2", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt2", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate2", typeof(string));//ชำระครั้งล่าสุด
            //CARD 3
            dataTemplate.Columns.Add("CardNo3", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt3", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt3", typeof(double));//ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt3", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt3", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate3", typeof(string));//ชำระครั้งล่าสุด
                                                                     //CARD 4        
            dataTemplate.Columns.Add("CardNo4", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt4", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt4", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt4", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt4", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate4", typeof(string));//ชำระครั้งล่าสุด
            //CARD 5
            dataTemplate.Columns.Add("CardNo5", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt5", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt5", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt5", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt5", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate5", typeof(string));//ชำระครั้งล่าสุด
            //CARD 6
            dataTemplate.Columns.Add("CardNo6", typeof(string));
            dataTemplate.Columns.Add("JudgmentAmnt6", typeof(double)); //ยอดพิพากษา
            dataTemplate.Columns.Add("PrincipleAmnt6", typeof(double)); //ต้นเงิน
            dataTemplate.Columns.Add("PayAfterJudgAmt6", typeof(double));//ชำระหลังพิพากษา
            dataTemplate.Columns.Add("DeptAmnt6", typeof(double));//ภาระหนี้ปัจจุบัน
            dataTemplate.Columns.Add("LastPayDate6", typeof(string));//ชำระครั้งล่าสุด
            //DataCustomer
            dataTemplate.Columns.Add("CustomerName", typeof(string)); //ชื่อ สกุล
            dataTemplate.Columns.Add("CustomerTel", typeof(string));//รหัสบัตรประชาชน
            dataTemplate.Columns.Add("CustomerID", typeof(string));//รหัสบัตรประชาชน            
            dataTemplate.Columns.Add("LegalStatus", typeof(string));//สถานะคดี
            dataTemplate.Columns.Add("BlackNo", typeof(string));//เลขคดีดำ
            dataTemplate.Columns.Add("RedNo", typeof(string));//เลขคดีแดง
            dataTemplate.Columns.Add("JudgeDate", typeof(string));//วันพิพากษา        
            dataTemplate.Columns.Add("CourtName", typeof(string));
            dataTemplate.Columns.Add("LegalExecRemark", typeof(string)); //หมายเหตุบังคับคดี
            dataTemplate.Columns.Add("LegalExecDate", typeof(string)); //วันที่ยึด อายัดทรัพย์
            //Collector Name
            dataTemplate.Columns.Add("CollectorName", typeof(string)); // ชื่อ collcctor
            dataTemplate.Columns.Add("CollectorTeam", typeof(string)); //ทีม
            dataTemplate.Columns.Add("CollectorTel", typeof(string)); //โทร
            return dataTemplate;
        }
        public DataTable doCreateFestCustomTemplate()
        {
            DataTable dataTemplate = new DataTable();
            dataTemplate.Columns.Add("CustomerID", typeof(string));
            dataTemplate.Columns.Add("CustomerName", typeof(string));
            dataTemplate.Columns.Add("LedNumber", typeof(string));
            dataTemplate.Columns.Add("WorkNo", typeof(string)); 
            dataTemplate.Columns.Add("CardNo1", typeof(string));
            dataTemplate.Columns.Add("AccCloseAmnt1", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt1", typeof(double));
            dataTemplate.Columns.Add("Installment6Amnt1", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt1", typeof(double));
            dataTemplate.Columns.Add("Installment12Amnt1", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt1", typeof(double));
            dataTemplate.Columns.Add("Installment24Amnt1", typeof(double));
            dataTemplate.Columns.Add("CardNo2", typeof(string));
            dataTemplate.Columns.Add("AccCloseAmnt2", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt2", typeof(double));
            dataTemplate.Columns.Add("Installment6Amnt2", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt2", typeof(double));
            dataTemplate.Columns.Add("Installment12Amnt2", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt2", typeof(double));
            dataTemplate.Columns.Add("Installment24Amnt2", typeof(double));
            dataTemplate.Columns.Add("CardNo3", typeof(string));
            dataTemplate.Columns.Add("AccCloseAmnt3", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt3", typeof(double));
            dataTemplate.Columns.Add("Installment6Amnt3", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt3", typeof(double));
            dataTemplate.Columns.Add("Installment12Amnt3", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt3", typeof(double));
            dataTemplate.Columns.Add("Installment24Amnt3", typeof(double));
            dataTemplate.Columns.Add("CardNo4", typeof(string));
            dataTemplate.Columns.Add("AccCloseAmnt4", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt4", typeof(double));
            dataTemplate.Columns.Add("Installment6Amnt4", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt4", typeof(double));
            dataTemplate.Columns.Add("Installment12Amnt4", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt4", typeof(double));
            dataTemplate.Columns.Add("Installment24Amnt4", typeof(double));
            dataTemplate.Columns.Add("CardNo5", typeof(string));
            dataTemplate.Columns.Add("AccCloseAmnt5", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt5", typeof(double));
            dataTemplate.Columns.Add("Installment6Amnt5", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt5", typeof(double));
            dataTemplate.Columns.Add("Installment12Amnt5", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt5", typeof(double));
            dataTemplate.Columns.Add("Installment24Amnt5", typeof(double));
            dataTemplate.Columns.Add("CardNo6", typeof(string));
            dataTemplate.Columns.Add("AccCloseAmnt6", typeof(double));
            dataTemplate.Columns.Add("AccClose6Amnt6", typeof(double));
            dataTemplate.Columns.Add("Installment6Amnt6", typeof(double));
            dataTemplate.Columns.Add("AccClose12Amnt6", typeof(double));
            dataTemplate.Columns.Add("Installment12Amnt6", typeof(double));
            dataTemplate.Columns.Add("AccClose24Amnt6", typeof(double));
            dataTemplate.Columns.Add("Installment24Amnt6", typeof(double));
            dataTemplate.Columns.Add("LegalExecRemark", typeof(string));
            return dataTemplate;
        }
        public DataTable doCreateFestDataTemplate()
        {
            DataTable dataTemplate = new DataTable();
            dataTemplate.Columns.Add("WorkNo", typeof(string)); //เลขชุดงาน
            dataTemplate.Columns.Add("LedNumber", typeof(string));// ลำดับกรม
            dataTemplate.Columns.Add("CustomerID", typeof(string));//รหัสบัตรประชาชน
            dataTemplate.Columns.Add("CustomerName", typeof(string)); //ชื่อ สกุล
            dataTemplate.Columns.Add("LegalExecRemark", typeof(string));
            return dataTemplate;    
        }       
        public DataTable doCreateDataShowTable()
        {
            DataTable dtDataShow = new DataTable();
            dtDataShow.Columns.Add("IsSelect", typeof(bool));
            dtDataShow.Columns.Add("WorkNo", typeof(string));
            dtDataShow.Columns.Add("LedNumber", typeof(string)); 
            dtDataShow.Columns.Add("CustomerID", typeof(string)); 
            dtDataShow.Columns.Add("CustomerName", typeof(string));
            dtDataShow.Columns.Add("LegalStatus", typeof(string));
            dtDataShow.Columns.Add("LegalExecRemark", typeof(string));
            dtDataShow.Columns.Add("CardStatus", typeof(string));
            dtDataShow.Columns.Add("CaseID", typeof(string));
            return dtDataShow;
        }
        public DataTable doCreateDataDatailShowTable()
        {
            DataTable dtDataShow = new DataTable();
            dtDataShow.Columns.Add("IsSelect", typeof(bool));
            dtDataShow.Columns.Add("CustomerID", typeof(string));
            dtDataShow.Columns.Add("CustomerName", typeof(string));
            dtDataShow.Columns.Add("LegalStatus", typeof(string));
            dtDataShow.Columns.Add("CardStatus", typeof(string));
            dtDataShow.Columns.Add("LegalExecDate", typeof(string));
            dtDataShow.Columns.Add("LegalExecRemark", typeof(string));
            dtDataShow.Columns.Add("CaseID", typeof(string));
            return dtDataShow;
        }
        public DataTable doCreateResultDataTable()
        {
            DataTable dtresult = new DataTable();
            dtresult.Columns.Add("CustomerID", typeof(string));
            dtresult.Columns.Add("Result", typeof(string));
            dtresult.Columns.Add("Remark", typeof(string));
            return dtresult;
        }
    }
}

