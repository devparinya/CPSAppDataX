using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace CpsDataApp.Models
{
    public class DataCPSPerson
    {
        public string? CaseID { get; set; }
        public string? CardStatus { get; set; }
        //DATA 
        public string? LedNumber { get; set; }// ลำดับกรม
        public int WorkNo { get; set; } //เลขชุดงาน
        int Maxmonth { get; set; } //จำนวนสูงสุดให้ผ่อน 
        //CARD 1
        public string? CardNo1 { get; set; }
        public double JudgmentAmnt1 { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt1 { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt1 { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt1 { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate1 { get; set; }//ชำระครั้งล่าสุด
        //CARD 2
        public string? CardNo2 { get; set; }
        public double JudgmentAmnt2 { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt2 { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt2 { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt2 { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate2 { get; set; }//ชำระครั้งล่าสุด
        //CARD 3
        public string? CardNo3 { get; set; }
        public double JudgmentAmnt3 { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt3 { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt3 { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt3 { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate3 { get; set; }//ชำระครั้งล่าสุด
        //CARD 4        
        public string? CardNo4 { get; set; }
        public double JudgmentAmnt4 { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt4 { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt4 { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt4 { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate4 { get; set; }//ชำระครั้งล่าสุด
        //CARD 5
        public string? CardNo5 { get; set; }
        public double JudgmentAmnt5 { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt5 { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt5 { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt5 { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate5 { get; set; }//ชำระครั้งล่าสุด
        //CARD 6
        public string? CardNo6 { get; set; }
        public double JudgmentAmnt6 { get; set; } //ยอดพิพากษา
        public double PrincipleAmnt6 { get; set; } //ต้นเงิน
        public double PayAfterJudgAmt6 { get; set; }//ชำระหลังพิพากษา
        public double DeptAmnt6 { get; set; }//ภาระหนี้ปัจจุบัน
        public string? LastPayDate6 { get; set; }//ชำระครั้งล่าสุด
        //DataCustomer
        public string? CustomerName { get; set; } //ชื่อ สกุล
        public string? CustomerID { get; set; }//รหัสบัตรประชาชน
        public string? CustomerTel { get; set; }//โทร
        public string? LegalStatus { get; set; }//สถานะคดี
        public string? BlackNo { get; set; }//เลขคดีดำ
        public string? RedNo { get; set; }//เลขคดีแดง
        public string? JudgeDate { get; set; }//วันพิพากษา        
        public string? CourtName { get; set; }
        public string? LegalExecRemark { get; set; } //หมายเหตุบังคับคดี --
        public string? LegalExecDate { get; set; } //วันที่ยึด อายัดทรัพย์ --
        //Collector Name
        public string? CollectorName { get; set; } // ชื่อ collcctor
        public string? CollectorTeam { get; set; } //ทีม
        public string? CollectorTel { get; set; } //โทร       
        public string? CustomFlag { get; set; }
    }
}
