﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CPSAppData.Models
{
    public class DataCPSMaster
    {
        public string? CaseID { get; set; }
        public string? CardStatus { get; set; }
        public int ListNo { get; set; } //เลขชุดงาน
        //DATA 
        public string? LedNumber { get; set; }// ลำดับกรม
        public int WorkNo { get; set; } //เลขชุดงาน

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
        public string? CustomFlag { get; set; }
    }
}
