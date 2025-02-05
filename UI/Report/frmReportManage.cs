using CPSAppData.Models;
using CPSAppData.Services;
using CPSAppData.UI.BaseForm;
using CpsDataApp.Models;
using CpsDataApp.Services;
using QueueAppManager.Service;
using System.Collections;
using System.Data;
namespace CPSAppData.UI.Report
{
    public partial class frmReportManage : frmBaseWindow
    {
        #region Member
        SettingData setdata = new SettingData();
        dataTableService dtService = new dataTableService();
        reportService reportsrv = new reportService();
        calculateDataService calcsrv = new calculateDataService();
        sqliteDataService sqlitedsrv = new sqliteDataService();
        List<FestCustom> customData = new List<FestCustom>();
        dateTimeHelper dateHelper = new dateTimeHelper();
        DataTable datatableshow;
        DataTable? datatablexls;
        ArrayList ARRdataForPrint = new ArrayList();

        string typerange = string.Empty;
        #endregion
        #region Constructor
        public frmReportManage()
        {
            InitializeComponent();
            doLoadSettingData();
            datatableshow = dtService.doCreateDataShowTable();
            dataGridShow.DataSource = datatableshow;
            doSettingGridData();
        }
        #endregion
        #region Inittail 
        private void doSettingGridData()
        {
            dataGridShow.Columns["IsSelect"].HeaderText = "#";
            dataGridShow.Columns["WorkNo"].HeaderText = "ลำดับที่";
            dataGridShow.Columns["LedNumber"].HeaderText = "ลำดับกรม";
            dataGridShow.Columns["CustomerID"].HeaderText = "เลขที่บัตรประชาชน";
            dataGridShow.Columns["CustomerName"].HeaderText = "ชื่อ-นามสกุล";
            dataGridShow.Columns["CardStatus"].HeaderText = "สถานะ(หลัก/เสริม)";
            dataGridShow.Columns["LegalStatus"].HeaderText = "สถานะทางคดี";
            dataGridShow.Columns["LegalExecRemark"].HeaderText = "หมายเหตุบังคับคดี";
            dataGridShow.Columns["CaseID"].HeaderText = "CaseID";

            dataGridShow.Columns["IsSelect"].Width = 20;
            dataGridShow.Columns["WorkNo"].Width = 90;
            dataGridShow.Columns["LedNumber"].Width = 90;
            dataGridShow.Columns["CustomerID"].Width = 130;
            dataGridShow.Columns["CustomerName"].Width = 180;
            dataGridShow.Columns["LegalStatus"].Width = 110;
            dataGridShow.Columns["CardStatus"].Width = 110;            
            dataGridShow.Columns["LegalExecRemark"].Width = 230;
            dataGridShow.Columns["CaseID"].Width = 90;

            dataGridShow.Columns["IsSelect"].ReadOnly = false;
            dataGridShow.Columns["WorkNo"].ReadOnly = true;
            dataGridShow.Columns["LedNumber"].ReadOnly = true;
            dataGridShow.Columns["CustomerID"].ReadOnly = true;
            dataGridShow.Columns["CustomerName"].ReadOnly = true;
            dataGridShow.Columns["LegalStatus"].ReadOnly = true;            
            dataGridShow.Columns["CardStatus"].ReadOnly = true;
            dataGridShow.Columns["LegalExecRemark"].ReadOnly = true;
            dataGridShow.Columns["CaseID"].ReadOnly = true;

            dataGridShow.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }     
        private void doLoadSettingData()
        {
            setdata = sqlitedsrv.doLoadSettingData();
            if (setdata != null)
            {
                txt_path_c2.Text = setdata.C2PathPdf;
                txt_path_table.Text = setdata.TablePathPdf;
                txt_festNo.Text = setdata.FestNo;
                txt_datefest.Text = dateHelper.doGetShortDateFromDBToPDF(setdata.FestDate ?? "");
                txt_dateatcalculate.Text = dateHelper.doGetShortDateFromDBToPDF(setdata.DateAtCalulate ?? "");
                txt_festName.Text = setdata.FestName;
                txt_firstInstalldate.Text = dateHelper.doGetShortDateFromDBToPDF(setdata.FistDateInstall ?? "");
                txt_maxmonth.Text = Convert.ToString(setdata.MaxMonth);
            }
        }
        #endregion
        #region Data Prepare
        private List<DataCPSCard> doConvertToCardDataCPS(DataCPSPerson dataperson) // ConvertData 1 person 6 card
        {
            List<DataCPSCard> dataCPSCardList = new List<DataCPSCard>();

            if (dataperson != null)
            {
                if (dataperson.CustomFlag == "Y")
                {
                    string customid = dataperson.CustomerID ?? "";
                    string caseid = dataperson.CustomerID ?? "";
                    if (!string.IsNullOrEmpty(customid))
                    {
                        List<FestCustom> datacustom = sqlitedsrv.doGetDataCustomWithID(customid, caseid);
                        if (datacustom.Count > 0)
                        {

                            string cardno1 = datacustom[0].CardNo1 ?? "";
                            if (!string.IsNullOrEmpty(cardno1))
                            {
                                int indexcard1 = findindexCardNo(dataperson, cardno1);
                                if (indexcard1 >= 0)
                                {
                                    DataCPSCard datacard = new DataCPSCard();
                                    DataCPSPerson datartn = doGetDataCardWithIndex(indexcard1, dataperson);
                                    datacard.CardNo = cardno1;
                                    datacard.JudgmentAmnt = datartn.JudgmentAmnt1;//ยอดพิพากษา
                                    datacard.PrincipleAmnt = datartn.PrincipleAmnt1;//ต้นเงิน
                                    datacard.PayAfterJudgAmt = datartn.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                    datacard.DeptAmnt = datartn.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                    datacard.LastPayDate = datartn.LastPayDate1;//ชำระครั้งล่าสุด
                                    doSetDataHederCard(ref datacard, dataperson);
                                    dataCPSCardList.Add(datacard);
                                }
                            }
                            string cardno2 = datacustom[0].CardNo2 ?? "";
                            if (!string.IsNullOrEmpty(cardno2))
                            {
                                int indexcard2 = findindexCardNo(dataperson, cardno2);
                                if (indexcard2 >= 0)
                                {
                                    DataCPSCard datacard = new DataCPSCard();
                                    DataCPSPerson datartn = doGetDataCardWithIndex(indexcard2, dataperson);
                                    datacard.CardNo = cardno2;
                                    datacard.JudgmentAmnt = datartn.JudgmentAmnt1;//ยอดพิพากษา
                                    datacard.PrincipleAmnt = datartn.PrincipleAmnt1;//ต้นเงิน
                                    datacard.PayAfterJudgAmt = datartn.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                    datacard.DeptAmnt = datartn.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                    datacard.LastPayDate = datartn.LastPayDate1;//ชำระครั้งล่าสุด
                                    doSetDataHederCard(ref datacard, dataperson);
                                    dataCPSCardList.Add(datacard);
                                }
                            }
                            string cardno3 = datacustom[0].CardNo3 ?? "";
                            if (!string.IsNullOrEmpty(cardno3))
                            {
                                int indexcard3 = findindexCardNo(dataperson, cardno3);
                                if (indexcard3 >= 0)
                                {
                                    DataCPSCard datacard = new DataCPSCard();
                                    DataCPSPerson datartn = doGetDataCardWithIndex(indexcard3, dataperson);
                                    datacard.CardNo = cardno3;
                                    datacard.JudgmentAmnt = datartn.JudgmentAmnt1;//ยอดพิพากษา
                                    datacard.PrincipleAmnt = datartn.PrincipleAmnt1;//ต้นเงิน
                                    datacard.PayAfterJudgAmt = datartn.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                    datacard.DeptAmnt = datartn.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                    datacard.LastPayDate = datartn.LastPayDate1;//ชำระครั้งล่าสุด
                                    doSetDataHederCard(ref datacard, dataperson);
                                    dataCPSCardList.Add(datacard);
                                }
                            }
                            string cardno4 = datacustom[0].CardNo4 ?? "";
                            if (!string.IsNullOrEmpty(cardno4))
                            {
                                int indexcard4 = findindexCardNo(dataperson, cardno4);
                                if (indexcard4 >= 0)
                                {
                                    DataCPSCard datacard = new DataCPSCard();
                                    DataCPSPerson datartn = doGetDataCardWithIndex(indexcard4, dataperson);
                                    datacard.CardNo = cardno4;
                                    datacard.JudgmentAmnt = datartn.JudgmentAmnt1;//ยอดพิพากษา
                                    datacard.PrincipleAmnt = datartn.PrincipleAmnt1;//ต้นเงิน
                                    datacard.PayAfterJudgAmt = datartn.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                    datacard.DeptAmnt = datartn.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                    datacard.LastPayDate = datartn.LastPayDate1;//ชำระครั้งล่าสุด
                                    doSetDataHederCard(ref datacard, dataperson);
                                    dataCPSCardList.Add(datacard);
                                }
                            }
                            string cardno5 = datacustom[0].CardNo5 ?? "";
                            if (!string.IsNullOrEmpty(cardno5))
                            {
                                int indexcard5 = findindexCardNo(dataperson, cardno5);
                                if (indexcard5 >= 0)
                                {
                                    DataCPSCard datacard = new DataCPSCard();
                                    DataCPSPerson datartn = doGetDataCardWithIndex(indexcard5, dataperson);
                                    datacard.CardNo = cardno5;
                                    datacard.JudgmentAmnt = datartn.JudgmentAmnt1;//ยอดพิพากษา
                                    datacard.PrincipleAmnt = datartn.PrincipleAmnt1;//ต้นเงิน
                                    datacard.PayAfterJudgAmt = datartn.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                    datacard.DeptAmnt = datartn.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                    datacard.LastPayDate = datartn.LastPayDate1;//ชำระครั้งล่าสุด
                                    doSetDataHederCard(ref datacard, dataperson);
                                    dataCPSCardList.Add(datacard);
                                }
                            }
                            string cardno6 = datacustom[0].CardNo6 ?? "";
                            if (!string.IsNullOrEmpty(cardno6))
                            {
                                int indexcard6 = findindexCardNo(dataperson, cardno6);
                                if (indexcard6 >= 0)
                                {
                                    DataCPSCard datacard = new DataCPSCard();
                                    DataCPSPerson datartn = doGetDataCardWithIndex(indexcard6, dataperson);
                                    datacard.CardNo = cardno6;
                                    datacard.JudgmentAmnt = datartn.JudgmentAmnt1;//ยอดพิพากษา
                                    datacard.PrincipleAmnt = datartn.PrincipleAmnt1;//ต้นเงิน
                                    datacard.PayAfterJudgAmt = datartn.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                    datacard.DeptAmnt = datartn.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                    datacard.LastPayDate = datartn.LastPayDate1;//ชำระครั้งล่าสุด
                                    doSetDataHederCard(ref datacard, dataperson);
                                    dataCPSCardList.Add(datacard);
                                }
                            }
                        }
                    }

                }
                else
                {
                    for (int i = 0; i < 6; i++)
                    {
                        DataCPSCard datacard = new DataCPSCard();

                        switch (i)
                        {
                            case 0:
                                datacard.CardNo = dataperson.CardNo1;
                                datacard.JudgmentAmnt = dataperson.JudgmentAmnt1;//ยอดพิพากษา
                                datacard.PrincipleAmnt = dataperson.PrincipleAmnt1;//ต้นเงิน
                                datacard.PayAfterJudgAmt = dataperson.PayAfterJudgAmt1;//ชำระหลังพิพากษา
                                datacard.DeptAmnt = dataperson.DeptAmnt1;//ภาระหนี้ปัจจุบัน
                                datacard.LastPayDate = dataperson.LastPayDate1;//ชำระครั้งล่าสุด
                                datacard.CapitalAmnt = 0;//ต้นเงินปัจจุบัน (Calculate)
                                break;
                            case 1:
                                datacard.CardNo = dataperson.CardNo2;
                                datacard.JudgmentAmnt = dataperson.JudgmentAmnt2;//ยอดพิพากษา
                                datacard.PrincipleAmnt = dataperson.PrincipleAmnt2;//ต้นเงิน
                                datacard.PayAfterJudgAmt = dataperson.PayAfterJudgAmt2;//ชำระหลังพิพากษา
                                datacard.DeptAmnt = dataperson.DeptAmnt2;//ภาระหนี้ปัจจุบัน
                                datacard.LastPayDate = dataperson.LastPayDate2;//ชำระครั้งล่าสุด
                                datacard.CapitalAmnt = 0;//ต้นเงินปัจจุบัน (Calculate)
                                break;
                            case 2:
                                datacard.CardNo = dataperson.CardNo3;
                                datacard.JudgmentAmnt = dataperson.JudgmentAmnt3;//ยอดพิพากษา
                                datacard.PrincipleAmnt = dataperson.PrincipleAmnt3;//ต้นเงิน
                                datacard.PayAfterJudgAmt = dataperson.PayAfterJudgAmt3;//ชำระหลังพิพากษา
                                datacard.DeptAmnt = dataperson.DeptAmnt3;//ภาระหนี้ปัจจุบัน
                                datacard.LastPayDate = dataperson.LastPayDate3;//ชำระครั้งล่าสุด
                                datacard.CapitalAmnt = 0;//ต้นเงินปัจจุบัน (Calculate)
                                break;
                            case 3:
                                datacard.CardNo = dataperson.CardNo4;
                                datacard.JudgmentAmnt = dataperson.JudgmentAmnt4;//ยอดพิพากษา
                                datacard.PrincipleAmnt = dataperson.PrincipleAmnt4;//ต้นเงิน
                                datacard.PayAfterJudgAmt = dataperson.PayAfterJudgAmt4;//ชำระหลังพิพากษา
                                datacard.DeptAmnt = dataperson.DeptAmnt4;//ภาระหนี้ปัจจุบัน
                                datacard.LastPayDate = dataperson.LastPayDate4;//ชำระครั้งล่าสุด
                                datacard.CapitalAmnt = 0;//ต้นเงินปัจจุบัน (Calculate)
                                break;
                            case 4:
                                datacard.CardNo = dataperson.CardNo5;
                                datacard.JudgmentAmnt = dataperson.JudgmentAmnt5;//ยอดพิพากษา
                                datacard.PrincipleAmnt = dataperson.PrincipleAmnt5;//ต้นเงิน
                                datacard.PayAfterJudgAmt = dataperson.PayAfterJudgAmt5;//ชำระหลังพิพากษา
                                datacard.DeptAmnt = dataperson.DeptAmnt5;//ภาระหนี้ปัจจุบัน
                                datacard.LastPayDate = dataperson.LastPayDate5;//ชำระครั้งล่าสุด
                                datacard.CapitalAmnt = 0;//ต้นเงินปัจจุบัน (Calculate)
                                break;
                            case 5:
                                datacard.CardNo = dataperson.CardNo6;
                                datacard.JudgmentAmnt = dataperson.JudgmentAmnt6;//ยอดพิพากษา
                                datacard.PrincipleAmnt = dataperson.PrincipleAmnt6;//ต้นเงิน
                                datacard.PayAfterJudgAmt = dataperson.PayAfterJudgAmt6;//ชำระหลังพิพากษา
                                datacard.DeptAmnt = dataperson.DeptAmnt6;//ภาระหนี้ปัจจุบัน
                                datacard.LastPayDate = dataperson.LastPayDate6;//ชำระครั้งล่าสุด
                                datacard.CapitalAmnt = 0;//ต้นเงินปัจจุบัน (Calculate)
                                break;
                        }
                        if (string.IsNullOrEmpty(datacard.CardNo)) break;

                        datacard.WorkNo = dataperson.WorkNo;
                        datacard.LedNumber = dataperson.LedNumber;
                        datacard.CustomerName = dataperson.CustomerName; //ชื่อ สกุล
                        datacard.CustomerID = dataperson.CustomerID;//รหัสบัตรประชาชน
                        datacard.CustomerTel = dataperson.CustomerTel;//รหัสบัตรประชาชน
                        datacard.LegalStatus = dataperson.LegalStatus;//สถานะคดี
                        datacard.BlackNo = dataperson.BlackNo;//เลขคดีดำ
                        datacard.RedNo = dataperson.RedNo;//เลขคดีแดง
                        datacard.JudgeDate = dataperson.JudgeDate;//วันพิพากษา        
                        datacard.CourtName = dataperson.CourtName;//ชื่อศาล
                        datacard.LegalExecRemark = dataperson.LegalExecRemark; //หมายเหตุบังคับคดี
                        datacard.LegalExecDate = dataperson.LegalExecDate; //วันที่ยึด อายัดทรัพย์

                        datacard.CollectorName = dataperson.CollectorName; // ชื่อ collcctor
                        datacard.CollectorTeam = dataperson.CollectorTeam;//ทีม
                        datacard.CollectorTel = dataperson.CollectorTel; //โทร
                        datacard.CustomFlag = dataperson.CustomFlag;
                        datacard.CaseID = dataperson.CaseID;
                        datacard.CardStatus = dataperson.CardStatus;

                        datacard.AccCloseAmnt = 0;
                        datacard.AccClose6Amnt = 0;
                        datacard.Installment6Amnt = 0;
                        datacard.AccClose12Amnt = 0;
                        datacard.Installment12Amnt = 0;
                        datacard.AccClose24Amnt = 0;
                        datacard.Installment24Amnt = 0;
                        dataCPSCardList.Add(datacard);
                    }
                }
            }
            return dataCPSCardList;
        }
        private int findindexCardNo(DataCPSPerson dataCard, string cardno)
        {
            int indexrow = -1;
            if (dataCard.CardNo1 == cardno) return 1;
            if (dataCard.CardNo2 == cardno) return 2;
            if (dataCard.CardNo3 == cardno) return 3;
            if (dataCard.CardNo4 == cardno) return 4;
            if (dataCard.CardNo5 == cardno) return 5;
            if (dataCard.CardNo6 == cardno) return 6;
            return indexrow;
        }
        private List<DataCPSPerson> doConvertDataMasterToCPSPerson(List<DataCPSMaster> datamasterlist)
        {
            List<DataCPSPerson> dataCPSperson = new List<DataCPSPerson>();            
            for (int i = 0; i < datamasterlist.Count; i++)
            {               
                string customerid = datamasterlist[i].CustomerID??"";
                string caseid = datamasterlist[i].CaseID ?? "";              

                int rowindex  = dataCPSperson.FindIndex(item => item.CustomerID == customerid && item.CaseID == caseid);
                if(rowindex < 0)
                {
                    DataCPSPerson dataperson = new DataCPSPerson();
                    dataperson.WorkNo = datamasterlist[i].WorkNo;
                    dataperson.LedNumber = datamasterlist[i].LedNumber;
                    dataperson.CustomerName = datamasterlist[i].CustomerName;
                    dataperson.CustomerID = datamasterlist[i].CustomerID;
                    dataperson.CustomerTel = datamasterlist[i].CustomerTel;
                    dataperson.LegalStatus = datamasterlist[i].LegalStatus;
                    dataperson.BlackNo = datamasterlist[i].BlackNo;
                    dataperson.RedNo = datamasterlist[i].RedNo;
                    dataperson.JudgeDate = datamasterlist[i].JudgeDate;
                    dataperson.CourtName = datamasterlist[i].CourtName;
                    dataperson.LegalExecRemark = datamasterlist[i].LegalExecRemark;
                    dataperson.LegalExecDate = datamasterlist[i].LegalExecDate;
                    dataperson.CollectorName = datamasterlist[i].CollectorName;
                    dataperson.CollectorTeam = datamasterlist[i].CollectorTeam;
                    dataperson.CollectorTel = datamasterlist[i].CollectorTel;
                    dataperson.CustomFlag = datamasterlist[i].CustomFlag;
                    dataperson.CaseID = datamasterlist[i].CaseID;
                    dataperson.CardStatus = datamasterlist[i].CardStatus;
                    switch (i)
                    {
                        case 0:
                            dataperson.CardNo1 = datamasterlist[i].CardNo;
                            dataperson.JudgmentAmnt1 = datamasterlist[i].JudgmentAmnt;
                            dataperson.PrincipleAmnt1 = datamasterlist[i].PrincipleAmnt;
                            dataperson.PayAfterJudgAmt1 = datamasterlist[i].PayAfterJudgAmt;
                            dataperson.DeptAmnt1 = datamasterlist[i].DeptAmnt;
                            dataperson.LastPayDate1 = datamasterlist[i].LastPayDate;
                            break;
                        case 1:
                            dataperson.CardNo2 = datamasterlist[i].CardNo;
                            dataperson.JudgmentAmnt2 = datamasterlist[i].JudgmentAmnt;
                            dataperson.PrincipleAmnt2 = datamasterlist[i].PrincipleAmnt;
                            dataperson.PayAfterJudgAmt2 = datamasterlist[i].PayAfterJudgAmt;
                            dataperson.DeptAmnt2 = datamasterlist[i].DeptAmnt;
                            dataperson.LastPayDate2 = datamasterlist[i].LastPayDate;
                            break;
                        case 2:
                            dataperson.CardNo3 = datamasterlist[i].CardNo;
                            dataperson.JudgmentAmnt3 = datamasterlist[i].JudgmentAmnt;
                            dataperson.PrincipleAmnt3 = datamasterlist[i].PrincipleAmnt;
                            dataperson.PayAfterJudgAmt3 = datamasterlist[i].PayAfterJudgAmt;
                            dataperson.DeptAmnt3 = datamasterlist[i].DeptAmnt;
                            dataperson.LastPayDate3 = datamasterlist[i].LastPayDate;
                            break;
                        case 3:
                            dataperson.CardNo4 = datamasterlist[i].CardNo;
                            dataperson.JudgmentAmnt4 = datamasterlist[i].JudgmentAmnt;
                            dataperson.PrincipleAmnt4 = datamasterlist[i].PrincipleAmnt;
                            dataperson.PayAfterJudgAmt4 = datamasterlist[i].PayAfterJudgAmt;
                            dataperson.DeptAmnt4 = datamasterlist[i].DeptAmnt;
                            dataperson.LastPayDate4 = datamasterlist[i].LastPayDate;
                            break;
                        case 4:
                            dataperson.CardNo5 = datamasterlist[i].CardNo;
                            dataperson.JudgmentAmnt5 = datamasterlist[i].JudgmentAmnt;
                            dataperson.PrincipleAmnt5 = datamasterlist[i].PrincipleAmnt;
                            dataperson.PayAfterJudgAmt5 = datamasterlist[i].PayAfterJudgAmt;
                            dataperson.DeptAmnt5 = datamasterlist[i].DeptAmnt;
                            dataperson.LastPayDate5 = datamasterlist[i].LastPayDate;
                            break;
                        case 6:
                            dataperson.CardNo6 = datamasterlist[i].CardNo;
                            dataperson.JudgmentAmnt6 = datamasterlist[i].JudgmentAmnt;
                            dataperson.PrincipleAmnt6 = datamasterlist[i].PrincipleAmnt;
                            dataperson.PayAfterJudgAmt6 = datamasterlist[i].PayAfterJudgAmt;
                            dataperson.DeptAmnt6 = datamasterlist[i].DeptAmnt;
                            dataperson.LastPayDate6 = datamasterlist[i].LastPayDate;
                            break;
                    }

                    dataCPSperson.Add(dataperson);
                }
                else
                {
                    dataCPSperson[rowindex].WorkNo = datamasterlist[i].WorkNo;
                    dataCPSperson[rowindex].LedNumber = datamasterlist[i].LedNumber;
                    dataCPSperson[rowindex].CustomerName = datamasterlist[i].CustomerName;
                    dataCPSperson[rowindex].CustomerID = datamasterlist[i].CustomerID;
                    dataCPSperson[rowindex].CustomerTel = datamasterlist[i].CustomerTel;
                    dataCPSperson[rowindex].LegalStatus = datamasterlist[i].LegalStatus;
                    dataCPSperson[rowindex].BlackNo = datamasterlist[i].BlackNo;
                    dataCPSperson[rowindex].RedNo = datamasterlist[i].RedNo;
                    dataCPSperson[rowindex].JudgeDate = datamasterlist[i].JudgeDate;
                    dataCPSperson[rowindex].CourtName = datamasterlist[i].CourtName;
                    dataCPSperson[rowindex].LegalExecRemark = datamasterlist[i].LegalExecRemark;
                    dataCPSperson[rowindex].LegalExecDate = datamasterlist[i].LegalExecDate;
                    dataCPSperson[rowindex].CollectorName = datamasterlist[i].CollectorName;
                    dataCPSperson[rowindex].CollectorTeam = datamasterlist[i].CollectorTeam;
                    dataCPSperson[rowindex].CollectorTel = datamasterlist[i].CollectorTel;
                    dataCPSperson[rowindex].CustomFlag = datamasterlist[i].CustomFlag;
                    dataCPSperson[rowindex].CaseID = datamasterlist[i].CaseID;
                    dataCPSperson[rowindex].CardStatus = datamasterlist[i].CardStatus;
                    switch (i)
                    {
                        case 0:
                            dataCPSperson[rowindex].CardNo1 = datamasterlist[i].CardNo;
                            dataCPSperson[rowindex].JudgmentAmnt1 = datamasterlist[i].JudgmentAmnt;
                            dataCPSperson[rowindex].PrincipleAmnt1 = datamasterlist[i].PrincipleAmnt;
                            dataCPSperson[rowindex].PayAfterJudgAmt1 = datamasterlist[i].PayAfterJudgAmt;
                            dataCPSperson[rowindex].DeptAmnt1 = datamasterlist[i].DeptAmnt;
                            dataCPSperson[rowindex].LastPayDate1 = datamasterlist[i].LastPayDate;
                            break;
                        case 1:
                           dataCPSperson[rowindex].CardNo2 = datamasterlist[i].CardNo;
                           dataCPSperson[rowindex].JudgmentAmnt2 = datamasterlist[i].JudgmentAmnt;
                           dataCPSperson[rowindex].PrincipleAmnt2 = datamasterlist[i].PrincipleAmnt;
                           dataCPSperson[rowindex].PayAfterJudgAmt2 = datamasterlist[i].PayAfterJudgAmt;
                           dataCPSperson[rowindex].DeptAmnt2 = datamasterlist[i].DeptAmnt;
                            dataCPSperson[rowindex].LastPayDate2 = datamasterlist[i].LastPayDate;
                            break;
                        case 2:
                            dataCPSperson[rowindex].CardNo3 = datamasterlist[i].CardNo;
                            dataCPSperson[rowindex].JudgmentAmnt3 = datamasterlist[i].JudgmentAmnt;
                            dataCPSperson[rowindex].PrincipleAmnt3 = datamasterlist[i].PrincipleAmnt;
                            dataCPSperson[rowindex].PayAfterJudgAmt3 = datamasterlist[i].PayAfterJudgAmt;
                            dataCPSperson[rowindex].DeptAmnt3 = datamasterlist[i].DeptAmnt;
                            dataCPSperson[rowindex].LastPayDate3 = datamasterlist[i].LastPayDate;
                            break;
                        case 3:
                            dataCPSperson[rowindex].CardNo4 = datamasterlist[i].CardNo;
                            dataCPSperson[rowindex].JudgmentAmnt4 = datamasterlist[i].JudgmentAmnt;
                            dataCPSperson[rowindex].PrincipleAmnt4 = datamasterlist[i].PrincipleAmnt;
                            dataCPSperson[rowindex].PayAfterJudgAmt4 = datamasterlist[i].PayAfterJudgAmt;
                            dataCPSperson[rowindex].DeptAmnt4 = datamasterlist[i].DeptAmnt;
                            dataCPSperson[rowindex].LastPayDate4 = datamasterlist[i].LastPayDate;
                            break;
                        case 4:
                            dataCPSperson[rowindex].CardNo5 = datamasterlist[i].CardNo;
                            dataCPSperson[rowindex].JudgmentAmnt5 = datamasterlist[i].JudgmentAmnt;
                            dataCPSperson[rowindex].PrincipleAmnt5 = datamasterlist[i].PrincipleAmnt;
                            dataCPSperson[rowindex].PayAfterJudgAmt5 = datamasterlist[i].PayAfterJudgAmt;
                            dataCPSperson[rowindex].DeptAmnt5 = datamasterlist[i].DeptAmnt;
                            dataCPSperson[rowindex].LastPayDate5 = datamasterlist[i].LastPayDate;
                            break;
                        case 6:
                            dataCPSperson[rowindex].CardNo6 = datamasterlist[i].CardNo;
                            dataCPSperson[rowindex].JudgmentAmnt6 = datamasterlist[i].JudgmentAmnt;
                            dataCPSperson[rowindex].PrincipleAmnt6 = datamasterlist[i].PrincipleAmnt;
                            dataCPSperson[rowindex].PayAfterJudgAmt6 = datamasterlist[i].PayAfterJudgAmt;
                            dataCPSperson[rowindex].DeptAmnt6 = datamasterlist[i].DeptAmnt;
                            dataCPSperson[rowindex].LastPayDate6 = datamasterlist[i].LastPayDate;
                            break;
                    }
                }               
            }
            
            return dataCPSperson;
        }
        private DataCPSPerson doGetDataCardWithIndex(int index, DataCPSPerson sourcedata)
        {
            DataCPSPerson datacpscard = new DataCPSPerson();
            switch (index)
            {
                case 1:
                    datacpscard.JudgmentAmnt1 = sourcedata.JudgmentAmnt1;
                    datacpscard.PrincipleAmnt1 = sourcedata.PrincipleAmnt1;
                    datacpscard.PayAfterJudgAmt1 = sourcedata.PayAfterJudgAmt1;
                    datacpscard.DeptAmnt1 = sourcedata.DeptAmnt1;
                    datacpscard.LastPayDate1 = sourcedata.LastPayDate1;

                    break;
                case 2:
                    datacpscard.JudgmentAmnt1 = sourcedata.JudgmentAmnt2;
                    datacpscard.PrincipleAmnt1 = sourcedata.PrincipleAmnt2;
                    datacpscard.PayAfterJudgAmt1 = sourcedata.PayAfterJudgAmt2;
                    datacpscard.DeptAmnt1 = sourcedata.DeptAmnt2;
                    datacpscard.LastPayDate1 = sourcedata.LastPayDate2;
                    break;
                case 3:
                    datacpscard.JudgmentAmnt1 = sourcedata.JudgmentAmnt3;
                    datacpscard.PrincipleAmnt1 = sourcedata.PrincipleAmnt3;
                    datacpscard.PayAfterJudgAmt1 = sourcedata.PayAfterJudgAmt3;
                    datacpscard.DeptAmnt1 = sourcedata.DeptAmnt3;
                    datacpscard.LastPayDate1 = sourcedata.LastPayDate3;
                    break;
                case 4:
                    datacpscard.JudgmentAmnt1 = sourcedata.JudgmentAmnt4;
                    datacpscard.PrincipleAmnt1 = sourcedata.PrincipleAmnt4;
                    datacpscard.PayAfterJudgAmt1 = sourcedata.PayAfterJudgAmt4;
                    datacpscard.DeptAmnt1 = sourcedata.DeptAmnt4;
                    datacpscard.LastPayDate1 = sourcedata.LastPayDate4;
                    break;
                case 5:
                    datacpscard.JudgmentAmnt1 = sourcedata.JudgmentAmnt5;
                    datacpscard.PrincipleAmnt1 = sourcedata.PrincipleAmnt5;
                    datacpscard.PayAfterJudgAmt1 = sourcedata.PayAfterJudgAmt5;
                    datacpscard.DeptAmnt1 = sourcedata.DeptAmnt5;
                    datacpscard.LastPayDate1 = sourcedata.LastPayDate5;
                    break;
                case 6:
                    datacpscard.JudgmentAmnt1 = sourcedata.JudgmentAmnt6;
                    datacpscard.PrincipleAmnt1 = sourcedata.PrincipleAmnt6;
                    datacpscard.PayAfterJudgAmt1 = sourcedata.PayAfterJudgAmt6;
                    datacpscard.DeptAmnt1 = sourcedata.DeptAmnt6;
                    datacpscard.LastPayDate1 = sourcedata.LastPayDate6;
                    break;
            }
            return datacpscard;
        }
        private void doSetDataHederCard(ref DataCPSCard datacard, DataCPSPerson dataperson)
        {
            datacard.WorkNo = dataperson.WorkNo;
            datacard.LedNumber = dataperson.LedNumber;
            datacard.CustomerName = dataperson.CustomerName; //ชื่อ สกุล
            datacard.CustomerID = dataperson.CustomerID;//รหัสบัตรประชาชน
            datacard.CustomerTel = dataperson.CustomerTel;//รหัสบัตรประชาชน
            datacard.LegalStatus = dataperson.LegalStatus;//สถานะคดี
            datacard.BlackNo = dataperson.BlackNo;//เลขคดีดำ
            datacard.RedNo = dataperson.RedNo;//เลขคดีแดง
            datacard.JudgeDate = dataperson.JudgeDate;//วันพิพากษา        
            datacard.CourtName = dataperson.CourtName;//ชื่อศาล
            datacard.LegalExecRemark = dataperson.LegalExecRemark; //หมายเหตุบังคับคดี
            datacard.LegalExecDate = dataperson.LegalExecDate; //วันที่ยึด อายัดทรัพย์

            datacard.CollectorName = dataperson.CollectorName; // ชื่อ collcctor
            datacard.CollectorTeam = dataperson.CollectorTeam;//ทีม
            datacard.CollectorTel = dataperson.CollectorTel; //โทร
            datacard.CustomFlag = dataperson.CustomFlag;
            datacard.CaseID = dataperson.CaseID;
            datacard.CardStatus = dataperson.CardStatus;

            datacard.AccCloseAmnt = 0;
            datacard.AccClose6Amnt = 0;
            datacard.Installment6Amnt = 0;
            datacard.AccClose12Amnt = 0;
            datacard.Installment12Amnt = 0;
            datacard.AccClose24Amnt = 0;
            datacard.Installment24Amnt = 0;
        }
        private void doLodeDataWithRange(string typedata, string startdata, string enddata) //New
        {
            Cursor = Cursors.WaitCursor;
            List<DataCPSPerson> dataperson = new List<DataCPSPerson>();           
            switch (typedata)
            {
                case "LEDNO":
                    List<DataCPSMaster> datamasterled = sqlitedsrv.doGetDataCPSWithRangeLEDNumber(startdata, enddata);
                    dataperson = doConvertDataMasterToCPSPerson(datamasterled);
                    break;
                case "WORKNO":
                    int i_startdata = Convert.ToInt32(startdata);
                    int i_enddata = Convert.ToInt32(enddata);
                    List<DataCPSMaster> datamasterwno = sqlitedsrv.doGetDataCPSWithRangeWorkNo(i_startdata, i_enddata);
                    dataperson = doConvertDataMasterToCPSPerson(datamasterwno);
                    break;
            }

            if (dataperson.Count > 0)
            {
                doSetShowDataTable(dataperson);
            }
            else
            {
                MessageBox.Show("ไม่พบข้อมูล", "ค้นหา", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                datatableshow.Clear();
            }
            Cursor = Cursors.Default;
        }
        private bool doCreateC2DataReport(List<DataCPSCard> dataCardlist) // 1 person print
        {
            bool result = false;
            if (dataCardlist != null)
            {
                if (Path.Exists(txt_path_c2.Text))
                {
                    reportsrv.C2PathFile = txt_path_c2.Text;
                    result = reportsrv.doCreateC2PDFReport(dataCardlist, setdata);
                }
                else
                {
                    MessageBox.Show("Path C2 ไม่ถูกต้อง", "Warnning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    result = false;
                }
            }
            return result;
        }      
        private bool doCreateC2DataReportMerge(ArrayList dataarrcpscard)
        {
            bool result = false;
            string filename = string.Empty;
            if (dataarrcpscard != null)
            {
                if (Path.Exists(txt_path_c2.Text))
                {
                    reportsrv.C2PathFile = txt_path_c2.Text;
                    result = reportsrv.doCreateC2PDFMerge(dataarrcpscard, setdata, typerange);
                }
                else
                {
                    MessageBox.Show("Path C2 ไม่ถูกต้อง", "Warnning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    result = false;
                }
            }
            return result;
        }
        private bool doCreateDataTablePDFMerge(ArrayList dataarrcpscard)
        {
            bool result = false;
            List<DataCPSCard> cpscarddatalist = new List<DataCPSCard>();
            if (dataarrcpscard != null)
            {
                for (int i = 0; i < dataarrcpscard.Count; i++)
                {
                    var datacard = dataarrcpscard[i];
                    if (datacard != null)
                    {
                        cpscarddatalist = (List<DataCPSCard>)datacard;
                        for (int n = 0; n < cpscarddatalist.Count; n++)
                        {
                            DataCPSCard cardCPS = cpscarddatalist[n];
                            if (cpscarddatalist[n].CustomFlag == "Y")
                            {
                                if (i == 0) customData = sqlitedsrv.doGetDataCustomWithID(cardCPS.CustomerID ?? "", cardCPS.CaseID ?? "");
                                if (customData.Count > 0) calcsrv.CaluLateData6CardCustom(ref cardCPS, customData, i);
                            }
                            else
                            {
                                calcsrv.CaluLateData6Card(ref cardCPS, setdata);
                            }
                        }
                    }
                }

                if (Path.Exists(txt_path_table.Text))
                {
                    reportsrv.TablePathFile = txt_path_table.Text;
                    result = reportsrv.doCreateCPSTableReportMerge(dataarrcpscard, setdata, typerange);
                }
                else
                {
                    MessageBox.Show("Path ตารางเจรจา ไม่ถูกต้อง", "Warnning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    result = false;
                }
            }
            return result;
        }       
        private bool doCreateDataTablePDF(List<DataCPSCard> datacpscard)
        {
            
            bool result = false;
            if (datacpscard != null)
            {
                for (int i = 0; i < datacpscard.Count; i++)
                {
                    DataCPSCard cpscarddata = datacpscard[i];
                    if (cpscarddata.CustomFlag == "Y")
                    {
                        if (i == 0) customData = sqlitedsrv.doGetDataCustomWithID(cpscarddata.CustomerID ?? "", cpscarddata.CaseID ?? "");
                        if (customData.Count > 0) calcsrv.CaluLateData6CardCustom(ref cpscarddata, customData, i);
                    }
                    else
                    {
                        calcsrv.CaluLateData6Card(ref cpscarddata, setdata);
                    }
                    datacpscard[i] = cpscarddata;
                }
                if (Path.Exists(txt_path_table.Text))
                {
                    reportsrv.TablePathFile = txt_path_table.Text;
                    result = reportsrv.doCreateCPSTableReport(datacpscard, setdata);
                }
                else
                {
                    MessageBox.Show("Path ตารางเจรจา ไม่ถูกต้อง", "Warnning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    result = false;
                }
            }
            return result;
        } // 1 person print       
        private List<DataCPSCard> doConvertDataForReportCustIDCaseID(string customerid, string case_id) //new
        {
            List<DataCPSCard> dataCardList = new List<DataCPSCard>();   
            List<DataCPSMaster> datamasterlist = sqlitedsrv.doGetDataCPSMasterAllByCustomerID(customerid, case_id);
            if(datamasterlist.Count > 0)
            {
                List<DataCPSPerson> datapersonlist = doConvertDataMasterToCPSPerson(datamasterlist);
                dataCardList = doConvertToCardDataCPS(datapersonlist[0]);
            }
            return dataCardList;
        }
        private ArrayList doConvertDataForReportMerge(DataRow[] dataselected)
        {
            if(ARRdataForPrint.Count > 0) return ARRdataForPrint;
            ArrayList datacardarrlist = new ArrayList();
            if (dataselected == null) return datacardarrlist;            
            for (int i = 0; i < dataselected.Length; i++)
            {
                List<DataCPSCard> datacardlist = new List<DataCPSCard>();
                string? customerid = Convert.ToString(dataselected[i]["CustomerID"] is DBNull ? "" : dataselected[i]["CustomerID"]);
                string? case_id = Convert.ToString(dataselected[i]["CaseID"] is DBNull ? "" : dataselected[i]["CaseID"]);
                if (!(string.IsNullOrEmpty(customerid) || (string.IsNullOrEmpty(case_id))))
                {
                    List<DataCPSMaster> datamasterlist = sqlitedsrv.doGetDataCPSMasterAllByCustomerID(customerid, case_id);
                    if (datamasterlist.Count > 0)
                    {
                        List<DataCPSPerson> datapersonlist = doConvertDataMasterToCPSPerson(datamasterlist);
                        datacardlist = doConvertToCardDataCPS(datapersonlist[0]);
                        datacardarrlist.Add(datacardlist);
                    }
                }
            }
            ARRdataForPrint = datacardarrlist;
            return ARRdataForPrint;
        }
        #endregion
        #region Export Excel
        private void doCrreateTableExportWithRange()
        {
            DataRow[]? datafind;
            if (datatableshow != null)
            {
                if (typerange == "W")
                {
                    datafind = datatableshow.Select("IsSelect = true", "WorkNo");
                    ArrayList arrdataexport = doConvertDataForExport(datafind);
                    doCreateTableExport(arrdataexport);
                }
                if (typerange == "L")
                {
                    datafind = datatableshow.Select("IsSelect = true", "LedNumber");
                    ArrayList arrdataexport = doConvertDataForExport(datafind);
                    doCreateTableExport(arrdataexport);
                }
            }
        }
        private void doCreateExportDataTableXLS(List<DataCPSCard> datacpslist)
        {
            if (datatablexls == null) return;
            if (datacpslist == null) return;
            DataRow rowdataxls = datatablexls.NewRow();
            for (int i = 0; i < datacpslist.Count; i++)
            {
                if (i == 0)
                {
                    rowdataxls["MaxMonth"] = datacpslist[i].Maxmonth;
                    rowdataxls["CustomerID"] = datacpslist[i].CustomerID;
                    rowdataxls["CustomerName"] = datacpslist[i].CustomerName;
                    rowdataxls["CaseID"] = datacpslist[i].CaseID;
                    rowdataxls["CardStatus"] = datacpslist[i].CardStatus;
                    rowdataxls["CustomerTel"] = datacpslist[i].CustomerTel;
                    rowdataxls["LegalStatus"] = datacpslist[i].LegalStatus;
                    rowdataxls["BlackNo"] = datacpslist[i].BlackNo;
                    rowdataxls["RedNo"] = datacpslist[i].RedNo;
                    rowdataxls["JudgeDate"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].JudgeDate ?? string.Empty);
                    rowdataxls["CourtName"] = datacpslist[i].CourtName;
                    rowdataxls["LedNumber"] = datacpslist[i].LedNumber;
                    rowdataxls["WorkNo"] = datacpslist[i].WorkNo;
                    rowdataxls["LegalExecRemark"] = datacpslist[i].LegalExecRemark;
                    rowdataxls["LegalExecDate"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LegalExecDate ?? string.Empty);
                    rowdataxls["CollectorName"] = datacpslist[i].CollectorName;
                    rowdataxls["CollectorTeam"] = datacpslist[i].CollectorTeam;
                    rowdataxls["CollectorTel"] = datacpslist[i].CollectorTel;
                }

                switch (i)
                {
                    case 0:
                        rowdataxls["CardNo1"] = datacpslist[i].CardNo;
                        rowdataxls["JudgmentAmnt1"] = datacpslist[i].JudgmentAmnt;
                        rowdataxls["PrincipleAmnt1"] = datacpslist[i].PrincipleAmnt;
                        rowdataxls["PayAfterJudgAmt1"] = datacpslist[i].PayAfterJudgAmt;
                        rowdataxls["DeptAmnt1"] = datacpslist[i].DeptAmnt;
                        rowdataxls["LastPayDate1"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LastPayDate ?? string.Empty);
                        rowdataxls["AccCloseAmnt1"] = datacpslist[i].AccCloseAmnt;
                        rowdataxls["AccClose6Amnt1"] = datacpslist[i].AccClose6Amnt;
                        rowdataxls["AccClose12Amnt1"] = datacpslist[i].AccClose12Amnt;
                        rowdataxls["AccClose24Amnt1"] = datacpslist[i].AccClose24Amnt;
                        break;
                    case 1:
                        rowdataxls["CardNo2"] = datacpslist[i].CardNo;
                        rowdataxls["JudgmentAmnt2"] = datacpslist[i].JudgmentAmnt;
                        rowdataxls["PrincipleAmnt2"] = datacpslist[i].PrincipleAmnt;
                        rowdataxls["PayAfterJudgAmt2"] = datacpslist[i].PayAfterJudgAmt;
                        rowdataxls["DeptAmnt2"] = datacpslist[i].DeptAmnt;
                        rowdataxls["LastPayDate2"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LastPayDate ?? string.Empty);
                        rowdataxls["AccCloseAmnt2"] = datacpslist[i].AccCloseAmnt;
                        rowdataxls["AccClose6Amnt2"] = datacpslist[i].AccClose6Amnt;
                        rowdataxls["AccClose12Amnt2"] = datacpslist[i].AccClose12Amnt;
                        rowdataxls["AccClose24Amnt2"] = datacpslist[i].AccClose24Amnt;
                        break;
                    case 2:
                        rowdataxls["CardNo3"] = datacpslist[i].CardNo;
                        rowdataxls["JudgmentAmnt3"] = datacpslist[i].JudgmentAmnt;
                        rowdataxls["PrincipleAmnt3"] = datacpslist[i].PrincipleAmnt;
                        rowdataxls["PayAfterJudgAmt3"] = datacpslist[i].PayAfterJudgAmt;
                        rowdataxls["DeptAmnt3"] = datacpslist[i].DeptAmnt;
                        rowdataxls["LastPayDate3"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LastPayDate ?? string.Empty);
                        rowdataxls["AccCloseAmnt3"] = datacpslist[i].AccCloseAmnt;
                        rowdataxls["AccClose6Amnt3"] = datacpslist[i].AccClose6Amnt;
                        rowdataxls["AccClose12Amnt3"] = datacpslist[i].AccClose12Amnt;
                        rowdataxls["AccClose24Amnt3"] = datacpslist[i].AccClose24Amnt;
                        break;
                    case 3:
                        rowdataxls["CardNo4"] = datacpslist[i].CardNo;
                        rowdataxls["JudgmentAmnt4"] = datacpslist[i].JudgmentAmnt;
                        rowdataxls["PrincipleAmnt4"] = datacpslist[i].PrincipleAmnt;
                        rowdataxls["PayAfterJudgAmt4"] = datacpslist[i].PayAfterJudgAmt;
                        rowdataxls["DeptAmnt4"] = datacpslist[i].DeptAmnt;
                        rowdataxls["LastPayDate4"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LastPayDate ?? string.Empty);
                        rowdataxls["AccCloseAmnt4"] = datacpslist[i].AccCloseAmnt;
                        rowdataxls["AccClose6Amnt4"] = datacpslist[i].AccClose6Amnt;
                        rowdataxls["AccClose12Amnt4"] = datacpslist[i].AccClose12Amnt;
                        rowdataxls["AccClose24Amnt4"] = datacpslist[i].AccClose24Amnt;
                        break;
                    case 4:
                        rowdataxls["CardNo5"] = datacpslist[i].CardNo;
                        rowdataxls["JudgmentAmnt5"] = datacpslist[i].JudgmentAmnt;
                        rowdataxls["PrincipleAmnt5"] = datacpslist[i].PrincipleAmnt;
                        rowdataxls["PayAfterJudgAmt5"] = datacpslist[i].PayAfterJudgAmt;
                        rowdataxls["DeptAmnt5"] = datacpslist[i].DeptAmnt;
                        rowdataxls["LastPayDate5"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LastPayDate ?? string.Empty);
                        rowdataxls["AccCloseAmnt5"] = datacpslist[i].AccCloseAmnt;
                        rowdataxls["AccClose6Amnt5"] = datacpslist[i].AccClose6Amnt;
                        rowdataxls["AccClose12Amnt5"] = datacpslist[i].AccClose12Amnt;
                        rowdataxls["AccClose24Amnt5"] = datacpslist[i].AccClose24Amnt;
                        break;
                    case 5:
                        rowdataxls["CardNo6"] = datacpslist[i].CardNo;
                        rowdataxls["JudgmentAmnt6"] = datacpslist[i].JudgmentAmnt;
                        rowdataxls["PrincipleAmnt6"] = datacpslist[i].PrincipleAmnt;
                        rowdataxls["PayAfterJudgAmt6"] = datacpslist[i].PayAfterJudgAmt;
                        rowdataxls["DeptAmnt6"] = datacpslist[i].DeptAmnt;
                        rowdataxls["LastPayDate6"] = dateHelper.doGetShortDateTHFromDBToPDF(datacpslist[i].LastPayDate ?? string.Empty);
                        rowdataxls["AccCloseAmnt6"] = datacpslist[i].AccCloseAmnt;
                        rowdataxls["AccClose6Amnt6"] = datacpslist[i].AccClose6Amnt;
                        rowdataxls["AccClose12Amnt6"] = datacpslist[i].AccClose12Amnt;
                        rowdataxls["AccClose24Amnt6"] = datacpslist[i].AccClose24Amnt;
                        break;
                }
            }
            datatablexls.Rows.Add(rowdataxls);
        }
        private void doCreateTableExport(ArrayList arrdataexport)
        {
            List<DataCPSCard> CPSCardList = new List<DataCPSCard>();

            if (arrdataexport != null)
            {
                datatablexls = dtService.doCreateMasterCPSDataAfterCalculate();
                for (int i = 0; i < arrdataexport.Count; i++)
                {
                    var datacard = arrdataexport[i];
                    if (datacard != null)
                    {
                        CPSCardList = (List<DataCPSCard>)datacard;
                        for (int n = 0; n < CPSCardList.Count; n++)
                        {
                            DataCPSCard cardCPS = CPSCardList[n];
                            if (CPSCardList[n].CustomFlag == "Y")
                            {
                                if (i == 0) customData = sqlitedsrv.doGetDataCustomWithID(cardCPS.CustomerID ?? "", cardCPS.CaseID ?? "");
                                if (customData.Count > 0) calcsrv.CaluLateData6CardCustom(ref cardCPS, customData, i);
                            }
                            else
                            {
                                calcsrv.CaluLateData6Card(ref cardCPS, setdata);
                            }
                        }
                    }
                    doCreateExportDataTableXLS(CPSCardList);
                }
            }
        }
        private ArrayList doConvertDataForExport(DataRow[] dataselected)
        {
            ArrayList datacpsarrlist = new ArrayList();
            if (dataselected == null) return datacpsarrlist;
            for (int i = 0; i < dataselected.Length; i++)
            {
               
                List<DataCPSCard> datacardlist = new List<DataCPSCard>();

                string? customerid = Convert.ToString(dataselected[i]["CustomerID"] is DBNull ? "" : dataselected[i]["CustomerID"]);
                string? case_id = Convert.ToString(dataselected[i]["CaseID"] is DBNull ? "" : dataselected[i]["CaseID"]);
                if (!(string.IsNullOrEmpty(customerid) || (string.IsNullOrEmpty(case_id))))
                {
                    List<DataCPSMaster> datamasterlist = sqlitedsrv.doGetDataCPSMasterAllByCustomerID(customerid, case_id);
                    if (datamasterlist.Count > 0)
                    {
                        List<DataCPSPerson> datapersonlist = doConvertDataMasterToCPSPerson(datamasterlist);
                        datacardlist = doConvertToCardDataCPS(datapersonlist[0]);
                        datacpsarrlist.Add(datacardlist);
                    }
                }
            }
            return datacpsarrlist;
        }
        #endregion
        #region Print PDF
        private bool doPrintReportC2WithRange(bool showmsg = true)
        {
            bool result = false;
            if (datatableshow != null)
            {
                Cursor = Cursors.WaitCursor;
                DataRow[] datafind = datatableshow.Select("IsSelect = true");
                for (int i = 0; i < datafind.Length; i++)
                {
                    string? customerid = Convert.ToString(datafind[i]["CustomerID"] is DBNull ? "" : datafind[i]["CustomerID"]);
                    string? case_id = Convert.ToString(datafind[i]["CaseID"] is DBNull ? "" : datafind[i]["CaseID"]);
                    if (!(string.IsNullOrEmpty(customerid)||(string.IsNullOrEmpty(case_id))))
                    {
                       List<DataCPSCard> dataCardlist = doConvertDataForReportCustIDCaseID(customerid,case_id);
                        if (dataCardlist.Count > 0)
                        {
                            result = doCreateC2DataReport(dataCardlist);
                        }
                        else
                        {
                            result = false;
                        }
                        if (!result) break;
                    }
                }
                Cursor = Cursors.Default;
                if (result)
                {
                    if (showmsg) MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            return result;
        }
        private bool doPrintReportC2WithRangeMerge(bool showmsg = true)
        {
            bool result = false;
            DataRow[]? datafind;
            if (datatableshow != null)
            {
                Cursor = Cursors.WaitCursor;
                if (typerange == "W")
                {
                    datafind = datatableshow.Select("IsSelect = true", "WorkNo");
                    ArrayList arrcardcpsdata = doConvertDataForReportMerge(datafind);
                    result = doCreateC2DataReportMerge(arrcardcpsdata);
                }
                if (typerange == "L")
                {
                    datafind = datatableshow.Select("IsSelect = true", "LedNumber");
                    ArrayList arrcardcpsdata = doConvertDataForReportMerge(datafind);
                    result = doCreateC2DataReportMerge(arrcardcpsdata);
                }
                Cursor = Cursors.Default;
            }
            if (result)
            {
                if (showmsg) MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return result;
        }
        private bool doPrintReportTableWithRangeMerge(bool showmsg = true)
        {
            bool result = false;
            DataRow[]? datafind;
            if (datatableshow != null)
            {
                Cursor = Cursors.WaitCursor;
                if (typerange == "W")
                {
                    datafind = datatableshow.Select("IsSelect = true", "WorkNo");
                    ArrayList arrcardcpsdata = doConvertDataForReportMerge(datafind);
                    result = doCreateDataTablePDFMerge(arrcardcpsdata);
                }
                if (typerange == "L")
                {
                    datafind = datatableshow.Select("IsSelect = true", "LedNumber");
                    ArrayList arrcardcpsdata = doConvertDataForReportMerge(datafind);
                    result = doCreateDataTablePDFMerge(arrcardcpsdata);
                }
                Cursor = Cursors.Default;
            }
            if (result)
            {
                if (showmsg) MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return result;
        }
        private bool doPrintReportTableWithRange(bool showmsg = true)
        {
            bool result = false;
            if (datatableshow != null)
            {
                Cursor = Cursors.WaitCursor;
                DataRow[] datafind = datatableshow.Select("IsSelect = true");
                for (int i = 0; i < datafind.Length; i++)
                {
                    string? customerid = Convert.ToString(datafind[i]["CustomerID"] is DBNull ? "" : datafind[i]["CustomerID"]);
                    string? case_id = Convert.ToString(datafind[i]["CaseID"] is DBNull ? "" : datafind[i]["CaseID"]);
                    if (!(string.IsNullOrEmpty(customerid) || (string.IsNullOrEmpty(case_id))))
                    {
                        List<DataCPSCard> dataCardlist = doConvertDataForReportCustIDCaseID(customerid, case_id);
                        if (dataCardlist.Count > 0)
                        {
                            result = doCreateDataTablePDF(dataCardlist);
                        }
                        else
                        {
                            result = false;
                        }
                    }
                }
                Cursor = Cursors.Default;
            }
            if (result)
            {
                if (showmsg) MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return result;
        }

        #endregion
        #region Control Event
        private void btn_export_Excel_Click(object sender, EventArgs e)
        {
            string filexlspath = selectSavefile();
            if (string.IsNullOrEmpty(filexlspath)) return;
            Cursor = Cursors.WaitCursor;
            doCrreateTableExportWithRange();
            if (datatablexls != null)
            {
                if (datatablexls.Rows.Count > 0)
                {
                    excelDataService xlsserv = new excelDataService();
                    xlsserv.ExportToExcelAll(datatablexls, filexlspath);
                    MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filexlspath), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            Cursor = Cursors.Default;
        }
        private string selectSavefile()
        {
            string path = string.Empty;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Save a Excel File";
            string s_no = string.IsNullOrEmpty(txt_startlednumber.Text) ? txt_startworkno.Text : txt_startlednumber.Text;
            string e_no = string.IsNullOrEmpty(txt_endlednumber.Text) ? txt_endworkno.Text : txt_endlednumber.Text;
            if (string.IsNullOrEmpty(s_no) || string.IsNullOrEmpty(e_no))
            {
                saveFileDialog.FileName = "datacps_export.xlsx";
            }
            else
            {
                saveFileDialog.FileName = string.Format("datacps_export_{0}_to_{1}.xlsx", s_no, e_no);
            }
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = saveFileDialog.FileName;
            }
            return path;
        }        
        private void btn_save_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doSaveDataSetting();
            Cursor = Cursors.Default;
            MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void brn_getwithrange_Click(object sender, EventArgs e)
        {
            string start_workno = txt_startworkno.Text;
            string end_workno = txt_endworkno.Text;
            string start_ledno = txt_startlednumber.Text;
            string end_Ledno = txt_endlednumber.Text;
            ARRdataForPrint = new ArrayList();
            Cursor = Cursors.Default;
            if (!(string.IsNullOrEmpty(start_workno) || string.IsNullOrEmpty(end_workno)))
            {
                doLodeDataWithRange("WORKNO", txt_startworkno.Text, txt_endworkno.Text);
                typerange = "W";
            }
            if (!(string.IsNullOrEmpty(start_ledno) || string.IsNullOrEmpty(end_Ledno)))
            {
                doLodeDataWithRange("LEDNO", txt_startlednumber.Text, txt_endlednumber.Text);
                typerange = "L";
            }
            dataGridShow.DataSource = datatableshow;
            dataGridShow.ClearSelection();
            Cursor = Cursors.Default;
        }
        private void textClearOtherRange(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string txtname = textBox.Name;
            if (string.IsNullOrEmpty(textBox.Text)) return;
            if (txtname.Contains("lednumber"))
            {
                txt_startworkno.Text = string.Empty;
                txt_endworkno.Text = string.Empty;

            }
            if (txtname.Contains("workno"))
            {
                txt_startlednumber.Text = string.Empty;
                txt_endlednumber.Text += string.Empty;
            }
        }
        private void btn_C2RangePDF_Click(object sender, EventArgs e)
        {
            if (chk_merge.Checked)
            {
                doPrintReportC2WithRangeMerge();
            }
            else
            {
                doPrintReportC2WithRange();
            }
        }
        private void btn_TableRangePDF_Click(object sender, EventArgs e)
        {
            if (chk_merge.Checked)
            {
                doPrintReportTableWithRangeMerge();
            }
            else
            {
                doPrintReportTableWithRange();
            }

        }
        private void btn_C2TableRangePDF_Click(object sender, EventArgs e)
        {
            bool result = false;
            if (chk_merge.Checked)
            {
                result = doPrintReportC2WithRangeMerge(false);
                result = doPrintReportTableWithRangeMerge(false);
                if (result)
                {
                    MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("บันทึกสำเร็จไม่สำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                result = doPrintReportC2WithRange(false);
                result = doPrintReportTableWithRange(false);
                if (result)
                {
                    MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("บันทึกสำเร็จไม่สำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void btn_path_c2_Click(object sender, EventArgs e)
        {
            txt_path_c2.Text = browFolder();
        }
        private void btn_table_cal_Click(object sender, EventArgs e)
        {
            txt_path_table.Text = browFolder();
        }

        #endregion
        #region Other Method       
        private string browFolder()
        {
            string ls_path = string.Empty;
            FolderBrowserDialog Fld = new FolderBrowserDialog();
            Fld.SelectedPath = ls_path;
            Fld.ShowNewFolderButton = true;


            if (Fld.ShowDialog() == DialogResult.OK)
            {
                ls_path = Fld.SelectedPath;
            }
            return ls_path;
        }       
        private void doSaveDataSetting()
        {
            if (setdata != null)
            {
                setdata.C2PathPdf = txt_path_c2.Text;
                setdata.TablePathPdf = txt_path_table.Text;
                setdata.FestNo = txt_festNo.Text;
                setdata.FestDate = dateHelper.doGetDateUIToDB(txt_datefest.Text);
                setdata.DateAtCalulate = dateHelper.doGetDateUIToDB(txt_dateatcalculate.Text);
                setdata.FestName = txt_festName.Text;
                setdata.FistDateInstall = dateHelper.doGetDateUIToDB(txt_firstInstalldate.Text);
                setdata.MaxMonth = Convert.ToInt32(string.IsNullOrEmpty(txt_maxmonth.Text) ? "0" : txt_maxmonth.Text);
                sqlitedsrv.doUpdateSetingData(setdata);
            }
        }
        private void doSetShowDataTable(List<DataCPSPerson> datapersonist)
        {
            datatableshow = dtService.doCreateDataShowTable();
            for (int i = 0; i < datapersonist.Count; i++)
            {
                DataRow datarow = datatableshow.NewRow();
                datarow["IsSelect"] = true;
                datarow["LedNumber"] = datapersonist[i].LedNumber;
                datarow["WorkNo"] = datapersonist[i].WorkNo;
                datarow["CustomerID"] = datapersonist[i].CustomerID;
                datarow["CustomerName"] = datapersonist[i].CustomerName;
                datarow["LegalStatus"] = datapersonist[i].LegalStatus;
                datarow["CardStatus"] = datapersonist[i].CardStatus;
                datarow["LegalExecRemark"] = datapersonist[i].LegalExecRemark;
                datarow["CaseID"] = datapersonist[i].CaseID;
                datatableshow.Rows.Add(datarow);
            }
        }
        #endregion
    }
}