using CPSAppData.Models;
using CPSAppData.Services;
using CpsDataApp.Models;

namespace QueueAppManager.Service
{
    public class calculateDataService
    {
        dateTimeHelper dateHelper =  new dateTimeHelper();        
        public void CaluLateData6Card(ref DataCPSCard cardCPS,SettingData setdata)
        {
            int decamnt = 2;
            int maxmonth = 24;
            double PrincipleAmnt = cardCPS.PrincipleAmnt; //AA ==
            double PayAfterJudgeAmnt = cardCPS.PayAfterJudgAmt; //AB ==
            double DeptAmnt = cardCPS.DeptAmnt; //AD ภาระหนี้ปัจจุบัน ==

            if (!string.IsNullOrEmpty(cardCPS.JudgeDate))
            {
                maxmonth = CaculateMaxmonth(setdata.MaxMonth, cardCPS.JudgeDate, setdata.FistDateInstall??string.Empty); 
            }
            double CapitalAmntBalance = 0;//*ต้นเงินปัจจุบัน 1  =IF((AA2-AB2)<0,0,(AA2-AB2))  **AC**
            CapitalAmntBalance = Round(PrincipleAmnt - PayAfterJudgeAmnt,decamnt);
            if (CapitalAmntBalance < 0) CapitalAmntBalance = 0;

            double AccCloseAmnt = 0; //*ปิดบัญชีงวดเดียว IF(AND(AC2=0, AD2=0), 0, IF(AC2<1000, MAX(AD2*10%, 1000), AC2)) **AE**
            if ((CapitalAmntBalance == 0) && (DeptAmnt == 0))
            {
                AccCloseAmnt = 0;
            }
            else
            {
                double DeptAmnt0Per = Round(DeptAmnt * 0.10, decamnt);
                if (CapitalAmntBalance < 1000)
                {
                    AccCloseAmnt = 1000;
                    if (DeptAmnt0Per > AccCloseAmnt) AccCloseAmnt = DeptAmnt0Per;
                }
                else
                {
                    AccCloseAmnt = CapitalAmntBalance;
                }
            }
            cardCPS.CapitalAmnt = CapitalAmntBalance;
            cardCPS.AccCloseAmnt = AccCloseAmnt;


            double AccClose6Amnt = 0;//*ผ่อนชำระ 6 งวด IF(AE2 = 1000, 0, IF(($AD2 - $AC2) * 0.2 + $AC2 < AE2, $AD2 * 0.2, ($AD2 - $AC2) * 0.2 + $AC2)) **AF**
            if (AccCloseAmnt == 1000)
            {
                AccClose6Amnt = 0;
            }
            else
            {
                double CheckValue = Round(((DeptAmnt - CapitalAmntBalance) * 0.2) + CapitalAmntBalance, decamnt);
                if (CheckValue < AccCloseAmnt)
                {
                    AccClose6Amnt = Round(DeptAmnt * 0.2, decamnt);
                }
                else
                {
                    AccClose6Amnt = CheckValue;
                }
            }
            double Installment6Amnt = Round(AccClose6Amnt / 6, decamnt); //ผ่อนชำระ 6 งวด งวดละ **AG**

            cardCPS.AccClose6Amnt = AccClose6Amnt;
            cardCPS.Installment6Amnt = Installment6Amnt;

            double AccClose12Amnt = 0;//*ผ่อนชำระ 12 งวด IF(AE2 = 1000, 0, IF(($AD2 - $AC2) * 0.3 + $AC2 < AE2, $AD2 * 0.3, ($AD2 - $AC2) * 0.3 + $AC2)) **AH**
            if (AccCloseAmnt == 1000)
            {
                AccClose12Amnt = 0;
            }
            else
            {
                double CheckValue = Round(((DeptAmnt - CapitalAmntBalance) * 0.4)+ CapitalAmntBalance, decamnt);
                if (CheckValue < AccCloseAmnt)
                {
                    AccClose12Amnt = Round(DeptAmnt * 0.4, decamnt);
                }
                else
                {
                    AccClose12Amnt = CheckValue;
                }
            }
            double Installment12Amnt = Round(AccClose12Amnt / 12, decamnt); //ผ่อนชำระ 12 งวด งวดละ **AI**

            cardCPS.AccClose12Amnt = AccClose12Amnt;
            cardCPS.Installment12Amnt = Installment12Amnt;



            double AccClose24Amnt = 0; //*ผ่อนชำระ 24 งวด IF(AE2 = 1000, 0, IF(($AD2 - $AC2) * 0.4 + $AC2 < AE2, $AD2 * 0.4, ($AD2 - $AC2) * 0.4 + $AC2)) **AJ**
            if (AccCloseAmnt == 1000)
            {
                AccClose24Amnt = 0;
            }
            else
            {
                double CheckValue = Round(((DeptAmnt - CapitalAmntBalance) * 0.6) + CapitalAmntBalance, decamnt);
                if (CheckValue < AccCloseAmnt)
                {
                    AccClose24Amnt = Round(DeptAmnt * 0.6, decamnt);
                }
                else
                {
                    AccClose24Amnt = CheckValue;
                }
            }
            double Installment24Amnt = Round(AccClose24Amnt / 24, decamnt); //ผ่อนชำระ 24 งวด งวดละ **AK**

            cardCPS.AccClose24Amnt = AccClose24Amnt;
            cardCPS.Installment24Amnt = Installment24Amnt;
            cardCPS.Maxmonth = maxmonth;
        }
        public void CaluLateData6CardNew(ref DataCPSCard cardCPS, SettingData setdata)
        {
            int decamnt = 2;
            int maxmonth = 24;
            double PrincipleAmnt = cardCPS.PrincipleAmnt; //ต้นเงิน
            double PayAfterJudgeAmnt = cardCPS.PayAfterJudgAmt; //จ่ายหลังพิพากษา
            double DeptAmnt = cardCPS.DeptAmnt; //ภาระหนี้ปัจจุบัน

            if (!string.IsNullOrEmpty(cardCPS.JudgeDate))
            {
                maxmonth = CaculateMaxmonth(setdata.MaxMonth, cardCPS.JudgeDate, setdata.FistDateInstall ?? string.Empty);
            }
            double PrincipleAmntBalance = 0;// ต้นเงินปัจจุบัน - จ่ายหลังพิพากษา
            PrincipleAmntBalance = Round(PrincipleAmnt - PayAfterJudgeAmnt, decamnt);
            if (PrincipleAmntBalance < 0) PrincipleAmntBalance = 0; 

            double AccCloseAmnt = 0; //*ปิดบัญชีงวดเดียว IF(AND(AC2=0, AD2=0), 0, IF(AC2<1000, MAX(AD2*10%, 1000), AC2)) **AE**
            if ((PrincipleAmntBalance == 0) && (DeptAmnt == 0))
            {
                AccCloseAmnt = 0;
            }
            else if((PrincipleAmntBalance == 0)&&(DeptAmnt>0)) //จ่ายเกินเงินต้น
            {
                AccCloseAmnt = Round(DeptAmnt * 0.10, decamnt); // 10% ของภาระหนี้
            }
            else
            {
                double DeptAmnt0Per = Round(DeptAmnt * 0.10, decamnt); // 10% ของภาระหนี้ 
                if (PrincipleAmntBalance < 1000) //
                {
                    AccCloseAmnt = 1000;
                    if (DeptAmnt0Per > AccCloseAmnt) AccCloseAmnt = PrincipleAmntBalance + DeptAmnt0Per;
                }
                else
                {

                    AccCloseAmnt = PrincipleAmntBalance;
                }
            }
            cardCPS.CapitalAmnt = PrincipleAmntBalance;
            cardCPS.AccCloseAmnt = AccCloseAmnt;


            double AccClose6Amnt = 0;//*ผ่อนชำระ 6 งวด IF(AE2 = 1000, 0, IF(($AD2 - $AC2) * 0.2 + $AC2 < AE2, $AD2 * 0.2, ($AD2 - $AC2) * 0.2 + $AC2)) **AF**
            if (AccCloseAmnt == 1000)
            {
                AccClose6Amnt = 0;
            }
            else
            {
                double CheckValue = Round(((DeptAmnt - PrincipleAmntBalance) * 0.2) + PrincipleAmntBalance, decamnt);
                if (CheckValue < AccCloseAmnt)
                {
                    AccClose6Amnt = Round(DeptAmnt * 0.2, decamnt);
                }
                else
                {
                    AccClose6Amnt = CheckValue;
                }
            }
            double Installment6Amnt = Round(AccClose6Amnt / 6, decamnt); //ผ่อนชำระ 6 งวด งวดละ **AG**

            cardCPS.AccClose6Amnt = AccClose6Amnt;
            cardCPS.Installment6Amnt = Installment6Amnt;

            double AccClose12Amnt = 0;//*ผ่อนชำระ 12 งวด IF(AE2 = 1000, 0, IF(($AD2 - $AC2) * 0.3 + $AC2 < AE2, $AD2 * 0.3, ($AD2 - $AC2) * 0.3 + $AC2)) **AH**
            if (AccCloseAmnt == 1000)
            {
                AccClose12Amnt = 0;
            }
            else
            {
                double CheckValue = Round(((DeptAmnt - PrincipleAmntBalance) * 0.3) + PrincipleAmntBalance, decamnt);
                if (CheckValue < AccCloseAmnt)
                {
                    AccClose12Amnt = Round(DeptAmnt * 0.3, decamnt);
                }
                else
                {
                    AccClose12Amnt = CheckValue;
                }
            }
            double Installment12Amnt = Round(AccClose12Amnt / 12, decamnt); //ผ่อนชำระ 12 งวด งวดละ **AI**

            cardCPS.AccClose12Amnt = AccClose12Amnt;
            cardCPS.Installment12Amnt = Installment12Amnt;



            double AccClose24Amnt = 0; //*ผ่อนชำระ 24 งวด IF(AE2 = 1000, 0, IF(($AD2 - $AC2) * 0.4 + $AC2 < AE2, $AD2 * 0.4, ($AD2 - $AC2) * 0.4 + $AC2)) **AJ**
            if (AccCloseAmnt == 1000)
            {
                AccClose24Amnt = 0;
            }
            else
            {
                double CheckValue = Round(((DeptAmnt - PrincipleAmntBalance) * 0.4) + PrincipleAmntBalance, decamnt);
                if (CheckValue < AccCloseAmnt)
                {
                    AccClose24Amnt = Round(DeptAmnt * 0.4, decamnt);
                }
                else
                {
                    AccClose24Amnt = CheckValue;
                }
            }
            double Installment24Amnt = Round(AccClose24Amnt / 24, decamnt); //ผ่อนชำระ 24 งวด งวดละ **AK**

            cardCPS.AccClose24Amnt = AccClose24Amnt;
            cardCPS.Installment24Amnt = Installment24Amnt;
            cardCPS.Maxmonth = maxmonth;
        }
        public void CaluLateData6CardAddValue(ref List<DataCPSCard> cardlist)
        {
            double sumCapitalAmnt = 0;
            double sumCapitalAmnt25per = 0;
            double avg25per = 0;

            sumCapitalAmnt = cardlist.Sum(item => (item.DeptAmnt));
            sumCapitalAmnt25per = Round((2.5/100)* sumCapitalAmnt, 0);
            if (sumCapitalAmnt25per < 2500) sumCapitalAmnt25per = 2500;
            int countcard = cardlist.Count;
            avg25per = Round(sumCapitalAmnt25per / countcard,0);
            foreach (DataCPSCard card in cardlist) 
            {
                card.AccCloseAmnt = Round(card.AccCloseAmnt+avg25per,0);
                card.Maxmonth = 1;
            }
            
        }
        public void CaluLateData6CardCustom(ref DataCPSCard cardCPS, List<FestCustom> customData,int indexrow)
        {            
            int decamnt = 2;
            double PrincipleAmnt = cardCPS.PrincipleAmnt; //AA ==
            double PayAfterJudgeAmnt = cardCPS.PayAfterJudgAmt; //AB ==
            double DeptAmnt = cardCPS.DeptAmnt; //AD ภาระหนี้ปัจจุบัน ==

            double CapitalAmntBalance = 0;//*ต้นเงินปัจจุบัน 1  =IF((AA2-AB2)<0,0,(AA2-AB2))  **AC**
            CapitalAmntBalance = Round(PrincipleAmnt - PayAfterJudgeAmnt, decamnt);
            if (CapitalAmntBalance < 0) CapitalAmntBalance = 0;
            cardCPS.CapitalAmnt = CapitalAmntBalance;
            cardCPS.Maxmonth = 24;
            if (customData.Count > 0)
            {
                SetCustomData6Card(ref cardCPS, customData[0], indexrow + 1);
            }
          }       
        public void SetCustomData6Card(ref DataCPSCard cardCPS, FestCustom CustomData, int cadrno)
        {
            switch (cadrno)
            {
                case 1:                   
                    cardCPS.AccCloseAmnt = CustomData.AccCloseAmnt1;
                    cardCPS.AccClose6Amnt = CustomData.AccClose6Amnt1;                    
                    cardCPS.AccClose12Amnt = CustomData.AccClose12Amnt1;                   
                    cardCPS.AccClose24Amnt = CustomData.AccClose24Amnt1;

                    cardCPS.Installment6Amnt =  CustomData.Installment6Amnt1;  //Round(CustomData.AccClose6Amnt1 / 6, decamnt);
                    cardCPS.Installment12Amnt = CustomData.Installment12Amnt1; //Round(CustomData.AccClose12Amnt1 / 12, decamnt);
                    cardCPS.Installment24Amnt = CustomData.Installment24Amnt1; //Round(CustomData.AccClose24Amnt1 / 24, decamnt);
                    break;
                case 2:
                    cardCPS.AccCloseAmnt = CustomData.AccCloseAmnt2;
                    cardCPS.AccClose6Amnt = CustomData.AccClose6Amnt2;
                    cardCPS.AccClose12Amnt = CustomData.AccClose12Amnt2;
                    cardCPS.AccClose24Amnt = CustomData.AccClose24Amnt2;

                    cardCPS.Installment6Amnt =  CustomData.Installment6Amnt2;  //Round(CustomData.AccClose6Amnt2 / 6, decamnt);
                    cardCPS.Installment12Amnt = CustomData.Installment12Amnt2; //Round(CustomData.AccClose12Amnt2 / 12, decamnt);
                    cardCPS.Installment24Amnt = CustomData.Installment24Amnt2; //Round(CustomData.AccClose24Amnt2 / 24, decamnt);
                    break;
                case 3:
                    cardCPS.AccCloseAmnt =  CustomData.AccCloseAmnt3;
                    cardCPS.AccClose6Amnt = CustomData.AccClose6Amnt3;
                    cardCPS.AccClose12Amnt = CustomData.AccClose12Amnt3;
                    cardCPS.AccClose24Amnt = CustomData.AccClose24Amnt3;

                    cardCPS.Installment6Amnt =  CustomData.Installment6Amnt3;  //Round(CustomData.AccClose6Amnt3 / 6, decamnt);
                    cardCPS.Installment12Amnt = CustomData.Installment12Amnt3; //Round(CustomData.AccClose12Amnt3 / 12, decamnt);
                    cardCPS.Installment24Amnt = CustomData.Installment24Amnt3; //Round(CustomData.AccClose24Amnt3 / 24, decamnt);
                    break;
                case 4:
                    cardCPS.AccCloseAmnt = CustomData.AccCloseAmnt4;
                    cardCPS.AccClose6Amnt = CustomData.AccClose6Amnt4;
                    cardCPS.AccClose12Amnt = CustomData.AccClose12Amnt4;
                    cardCPS.AccClose24Amnt = CustomData.AccClose24Amnt4;

                    cardCPS.Installment6Amnt =  CustomData.Installment6Amnt4;  //Round(CustomData.AccClose6Amnt4 / 6, decamnt);
                    cardCPS.Installment12Amnt = CustomData.Installment12Amnt4; //Round(CustomData.AccClose12Amnt4 / 12, decamnt);
                    cardCPS.Installment24Amnt = CustomData.Installment24Amnt4; //Round(CustomData.AccClose24Amnt4 / 24, decamnt);
                    break;
                case 5:
                    cardCPS.AccCloseAmnt = CustomData.AccCloseAmnt5;
                    cardCPS.AccClose6Amnt = CustomData.AccClose6Amnt5;
                    cardCPS.AccClose12Amnt = CustomData.AccClose12Amnt5;
                    cardCPS.AccClose24Amnt = CustomData.AccClose24Amnt5;

                    cardCPS.Installment6Amnt =  CustomData.Installment6Amnt5;  //Round(CustomData.AccClose6Amnt5 / 6, decamnt);
                    cardCPS.Installment12Amnt = CustomData.Installment12Amnt5; //Round(CustomData.AccClose12Amnt5 / 12, decamnt);
                    cardCPS.Installment24Amnt = CustomData.Installment24Amnt5; //Round(CustomData.AccClose24Amnt5 / 24, decamnt);
                    break;
                case 6:
                    cardCPS.AccCloseAmnt = CustomData.AccCloseAmnt6;
                    cardCPS.AccClose6Amnt = CustomData.AccClose6Amnt6;
                    cardCPS.AccClose12Amnt = CustomData.AccClose12Amnt6;
                    cardCPS.AccClose24Amnt = CustomData.AccClose24Amnt6;

                    cardCPS.Installment6Amnt =  CustomData.Installment6Amnt6;  //Round(CustomData.AccClose6Amnt6 / 6, decamnt);
                    cardCPS.Installment12Amnt = CustomData.Installment12Amnt6; //Round(CustomData.AccClose12Amnt6 / 12, decamnt);
                    cardCPS.Installment24Amnt = CustomData.Installment24Amnt6; //Round(CustomData.AccClose24Amnt6 / 24, decamnt);
                    break;
            }           
           
        }
        public double Round(double d, int digitAmnt)
        {
            return Math.Round(d, digitAmnt, MidpointRounding.AwayFromZero);
        }
        public int CaculateMaxmonth(int beforexpire, string judedate,string firstdate)
        {
            if (string.IsNullOrEmpty(judedate)) return 24;

            int year_jude = Convert.ToInt16(judedate.Substring(0, 4));
            int month_jude = Convert.ToInt16(judedate.Substring(4, 2));
            int date_jude = Convert.ToInt16(judedate.Substring (6, 2));
           
            int year_first = Convert.ToInt16(firstdate.Substring(0, 4));
            int month_first = Convert.ToInt16(firstdate.Substring(4, 2));
            int date_firste = Convert.ToInt16(firstdate.Substring(6, 2));

            DateTime judeDate = new DateTime(year_jude, month_jude, date_jude); // วันที่พิพากษา
            DateTime expiredate = judeDate.AddMonths(120); // วันที่สิ้นสุดอายุความ

            DateTime firstDate = new DateTime(year_first, month_first, date_firste); // วันที่ผ่อนงวดแรก
            DateTime install6Date = firstDate.AddMonths(6); // งวดสุดท้าย   6 เดือน
            DateTime install12Date = firstDate.AddMonths(12);// งวดสุดท้าย  12 เดือน
            DateTime install24Date = firstDate.AddMonths(24);// งวดสุดท้าย  24 เดือน

            int rema_expiremonth = ((expiredate.Year - firstDate.Year)*12) + (expiredate.Month - firstDate.Month);
            int rema_expire6month = ((expiredate.Year - install6Date.Year) * 12) + (expiredate.Month - install6Date.Month);
            int rema_expire12month = ((expiredate.Year - install12Date.Year) * 12) + (expiredate.Month - install12Date.Month);
            int rema_expire24month = ((expiredate.Year - install24Date.Year) * 12) + (expiredate.Month - install24Date.Month);


            if (install6Date.Day < firstDate.Day)  { rema_expire6month--; }
            if (install12Date.Day < firstDate.Day) { rema_expire12month--;}
            if (install24Date.Day < firstDate.Day) { rema_expire24month--;}

            if (rema_expire6month >= 0) { rema_expire6month = rema_expire6month - beforexpire; }
            if (rema_expire12month >= 0) { rema_expire12month = rema_expire12month - beforexpire; }
            if (rema_expire24month >= 0) { rema_expire24month = rema_expire12month - beforexpire; }

            if (rema_expiremonth <= 0)
            {
                return 24;
            }
            else
            {
                if(rema_expire24month >0) return 24;
                if(rema_expire12month >0) return 12;
                if(rema_expire6month >0) return 6;
            }
            return 0;
        }

        
    }
}