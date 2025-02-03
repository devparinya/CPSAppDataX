namespace CpsDataApp.Models
{
    public class DataCPSCard
    {
        public string? CaseID { get; set; }
        public string? CardStatus { get; set; }
        //DATA 
        public string? LedNumber { get; set; }// ลำดับกรม
        public string? WorkNo { get; set; } //เลขชุดงาน
        public int Maxmonth { get; set; } //จำนวนสูงสุดให้ผ่อน 
        //CARD 1
        public string? CardNo { get; set; }
        public double JudgmentAmnt { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate { get; set; }//ชำระครั้งล่าสุด
        public double CapitalAmnt { get; set; } //ต้นเงินปัจจุบัน (Calculate)
        //DataCustomer
        public string? CustomerName { get; set; } //ชื่อ สกุล
        public string? CustomerID { get; set; }//รหัสบัตรประชาชน
        public string? CustomerTel { get; set; }//โทร

        public string? LegalStatus { get; set; }//สถานะคดี
        public string? BlackNo { get; set; }//เลขคดีดำ
        public string? RedNo { get; set; }//เลขคดีแดง
        public string? JudgeDate { get; set; }//วันพิพากษา        
        public string? CourtName { get; set; }//ชื่อศาล
        public string? LegalExecRemark { get; set; } //หมายเหตุบังคับคดี
        public string? LegalExecDate { get; set; } //วันที่ยึด อายัดทรัพย์
        //Collector Name
        public string? CollectorName { get; set; } // ชื่อ collcctor
        public string? CollectorTeam { get; set; } //ทีม
        public string? CollectorTel { get; set; } //โทร      
        //Calculate Data
        public double AccCloseAmnt { get; set; }
        public double AccClose6Amnt { get; set; }
        public double Installment6Amnt { get; set; }
        public double AccClose12Amnt { get; set; }
        public double Installment12Amnt { get; set; }
        public double AccClose24Amnt { get; set; }
        public double Installment24Amnt { get; set; }
        public string? CustomFlag { get; set; }
    }
}
