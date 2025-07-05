using CPSAppData.Models;
using CPSAppData.Services;
using CPSAppData.UI.BaseForm;
using CpsDataApp.Models;
using CpsDataApp.Services;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using QueueAppManager.Service;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Runtime.CompilerServices;
namespace CPSAppData.UI.Report
{
    public partial class frmAdvUserReport : frmBaseWindow
    {
        #region Member
        SettingData setdata = new SettingData();
        dataTableService dtService = new dataTableService();
        reportService reportsrv = new reportService();
        calculateDataService calcsrv = new calculateDataService();
        sqliteDataService sqlitedsrv = new sqliteDataService();
        dateTimeHelper dateHelper = new dateTimeHelper();
        DataTable datatabledetailshow;
        List<DataCPSCard> CARDPRINT = new List<DataCPSCard>();
        bool is_kzas;

        string typerange = string.Empty;
        #endregion
        #region Constructor
        public frmAdvUserReport()
        {
            InitializeComponent();
            doLoadSettingData();
            datatabledetailshow = dtService.doCreateDataDatailShowTable();
            dataGridDetailCard.DataSource = datatabledetailshow;
            doSettingGridDataCardDetail();
        }
        #endregion
        #region Inittail 
        private void doSettingGridDataCardDetail()
        {
            dataGridDetailCard.Columns["IsSelect"].HeaderText = "#";
            dataGridDetailCard.Columns["CustomerID"].HeaderText = "เลขที่บัตรประชาชน";
            dataGridDetailCard.Columns["CustomerName"].HeaderText = "ชื่อ-นามสกุล";
            dataGridDetailCard.Columns["CardStatus"].HeaderText = "สถานะ(หลัก/เสริม)";
            dataGridDetailCard.Columns["LegalStatus"].HeaderText = "สถานะทางคดี";
            dataGridDetailCard.Columns["LegalExecDate"].HeaderText = "วันที่ยึด";
            dataGridDetailCard.Columns["LegalExecRemark"].HeaderText = "หมายเหตุบังคับคดี";

            dataGridDetailCard.Columns["IsSelect"].Width = 20;
            dataGridDetailCard.Columns["CustomerID"].Width = 180;
            dataGridDetailCard.Columns["CustomerName"].Width = 200;
            dataGridDetailCard.Columns["CardStatus"].Width = 120;
            dataGridDetailCard.Columns["LegalStatus"].Width = 100;
            dataGridDetailCard.Columns["LegalExecDate"].Width = 90;
            dataGridDetailCard.Columns["LegalExecRemark"].Width = 180;

            dataGridDetailCard.Columns["IsSelect"].ReadOnly = false;
            dataGridDetailCard.Columns["CustomerID"].ReadOnly = true;
            dataGridDetailCard.Columns["CustomerName"].ReadOnly = true;
            dataGridDetailCard.Columns["CardStatus"].ReadOnly = true;
            dataGridDetailCard.Columns["LegalStatus"].ReadOnly = true;
            dataGridDetailCard.Columns["LegalExecDate"].ReadOnly = true;
            dataGridDetailCard.Columns["LegalExecRemark"].ReadOnly = true;
            dataGridDetailCard.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
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
                    string caseid = dataperson.CaseID ?? "";
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

        private List<DataCPSPerson> doConvertDataMasterToCPSPerson(List<DataCPSMaster> datamasterlist)
        {
            List<DataCPSPerson> dataCPSperson = new List<DataCPSPerson>();
            DataCPSPerson dataperson = new DataCPSPerson();
            for (int i = 0; i < datamasterlist.Count; i++)
            {
                if (i == 0)
                {
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
                }
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
                    case 5:
                        dataperson.CardNo6 = datamasterlist[i].CardNo;
                        dataperson.JudgmentAmnt6 = datamasterlist[i].JudgmentAmnt;
                        dataperson.PrincipleAmnt6 = datamasterlist[i].PrincipleAmnt;
                        dataperson.PayAfterJudgAmt6 = datamasterlist[i].PayAfterJudgAmt;
                        dataperson.DeptAmnt6 = datamasterlist[i].DeptAmnt;
                        dataperson.LastPayDate6 = datamasterlist[i].LastPayDate;
                        break;
                }
            }
            dataCPSperson.Add(dataperson);
            return dataCPSperson;

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
        //New
        private void doSearhCPSDataWithCusomerID()
        {
            string customer_id = string.Empty;

            datatabledetailshow.Clear();
            dataGridDetailCard.DataSource = datatabledetailshow;
            List<DataCPSMaster> dataMasterIDList = new List<DataCPSMaster>();
            string search_text_id = txt_search_custid_shdt.Text.Replace(" ", "");
            string search_text_lednumber = txt_queueno_shdt.Text.Replace(" ", "");
            if (!string.IsNullOrEmpty(search_text_id))
            {
                doClearControl(false);
                is_kzas = false;
                dataMasterIDList = sqlitedsrv.doGetDataCPSMasterByCustomerID(search_text_id);
                doGetDataMasterCPSByCaseID(ref dataMasterIDList);
                if (dataMasterIDList.Count > 0)
                {
                    if (dataMasterIDList != null)
                    {
                        for (int i = 0; i < dataMasterIDList.Count; i++)
                        {
                            string legalstatus = dataMasterIDList[i].LegalStatus ?? string.Empty;
                            if (legalstatus.Contains("KZAS")) { is_kzas = true; }
                            DataRow datadt = datatabledetailshow.NewRow();
                            datadt["IsSelect"] = false;
                            datadt["CustomerID"] = dataMasterIDList[i].CustomerID;
                            datadt["CustomerName"] = dataMasterIDList[i].CustomerName;
                            datadt["CardStatus"] = dataMasterIDList[i].CardStatus;
                            datadt["LegalStatus"] = dataMasterIDList[i].LegalStatus;
                            datadt["LegalExecDate"] = dataMasterIDList[i].LegalExecDate;
                            datadt["LegalExecRemark"] = dataMasterIDList[i].LegalExecRemark;
                            datadt["CaseID"] = dataMasterIDList[i].CaseID;
                            datatabledetailshow.Rows.Add(datadt);
                        }
                     
                    }
                    doSortDataTable(ref datatabledetailshow, "CustomerID ASC, CaseID ASC");
                }
                else
                {
                    MessageBox.Show("ไม่พบข้อมูลลูกค้า ในระบบ", "ไม่พบข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    doClearControl(false);
                    dataGridDetailCard.DataSource = datatabledetailshow;
                }
            }
            else
            {
                doClearControl(false);
                dataGridDetailCard.DataSource = datatabledetailshow;
            }
        }
        private void doSortDataTable(ref DataTable datasoure, string sort_str)
        {
            DataView view = datasoure.DefaultView;
            view.Sort = sort_str;

            datasoure = view.ToTable();
        }
        private void doGetDataMasterCPSByCaseID(ref List<DataCPSMaster> datamasterlist)
        {
            if (datamasterlist != null)
            {
                int countlist = datamasterlist.Count;
                for (int i = 0; i < countlist; i++)
                {
                    string? cus_id = Convert.ToString(datamasterlist[i].CustomerID is null ? string.Empty : datamasterlist[i].CustomerID);
                    string? case_id = Convert.ToString(datamasterlist[i].CaseID is null ? string.Empty : datamasterlist[i].CaseID);
                    if (!(string.IsNullOrEmpty(cus_id) || string.IsNullOrEmpty(case_id)))
                    {
                        sqlitedsrv.doGetDataCPSMasterByCaseID(ref datamasterlist, cus_id, case_id);
                    }
                }
            }
        }
        private void doClearControl(bool isclearsearch)
        {
            txt_custid_shdt.Text = string.Empty;
            txt_custname_shdt.Text = string.Empty;
            txt_queueno_shdt.Text = string.Empty;

            txt_cardnoshdt_1.Text = string.Empty;
            txt_judgmentamntshdt_1.Text = string.Empty;
            txt_principleamntshdt_1.Text = string.Empty;
            txt_payafterjudgamtshdt_1.Text = string.Empty;
            txt_deptamntshdt_1.Text = string.Empty;
            txt_lastpaydateshdt_1.Text = string.Empty;
            txt_closeamntshdt_1.Text = string.Empty;
            txt_close6amntshdt_1.Text = string.Empty;
            txt_install6amntshdt_1.Text = string.Empty;
            txt_close12amntshdt_1.Text = string.Empty;
            txt_install12amntshdt_1.Text = string.Empty;
            txt_close24amntshdt_1.Text = string.Empty;
            txt_install24amntshdt_1.Text = string.Empty;


            txt_cardnoshdt_2.Text = string.Empty;
            txt_judgmentamntshdt_2.Text = string.Empty;
            txt_principleamntshdt_2.Text = string.Empty;
            txt_payafterjudgamtshdt_2.Text = string.Empty;
            txt_deptamntshdt_2.Text = string.Empty;
            txt_lastpaydateshdt_2.Text = string.Empty;
            txt_closeamntshdt_2.Text = string.Empty;
            txt_close6amntshdt_2.Text = string.Empty;
            txt_install6amntshdt_2.Text = string.Empty;
            txt_close12amntshdt_2.Text = string.Empty;
            txt_install12amntshdt_2.Text = string.Empty;
            txt_close24amntshdt_2.Text = string.Empty;
            txt_install24amntshdt_2.Text = string.Empty;

            txt_cardnoshdt_3.Text = string.Empty;
            txt_judgmentamntshdt_3.Text = string.Empty;
            txt_principleamntshdt_3.Text = string.Empty;
            txt_payafterjudgamtshdt_3.Text = string.Empty;
            txt_deptamntshdt_3.Text = string.Empty;
            txt_lastpaydateshdt_3.Text = string.Empty;
            txt_closeamntshdt_3.Text = string.Empty;
            txt_close6amntshdt_3.Text = string.Empty;
            txt_install6amntshdt_3.Text = string.Empty;
            txt_close12amntshdt_3.Text = string.Empty;
            txt_install12amntshdt_3.Text = string.Empty;
            txt_close24amntshdt_3.Text = string.Empty;
            txt_install24amntshdt_3.Text = string.Empty;


            txt_cardnoshdt_4.Text = string.Empty;
            txt_judgmentamntshdt_4.Text = string.Empty;
            txt_principleamntshdt_4.Text = string.Empty;
            txt_payafterjudgamtshdt_4.Text = string.Empty;
            txt_deptamntshdt_4.Text = string.Empty;
            txt_lastpaydateshdt_4.Text = string.Empty;
            txt_closeamntshdt_4.Text = string.Empty;
            txt_close6amntshdt_4.Text = string.Empty;
            txt_install6amntshdt_4.Text = string.Empty;
            txt_close12amntshdt_4.Text = string.Empty;
            txt_install12amntshdt_4.Text = string.Empty;
            txt_close24amntshdt_4.Text = string.Empty;
            txt_install24amntshdt_4.Text = string.Empty;


            txt_cardnoshdt_5.Text = string.Empty;
            txt_judgmentamntshdt_5.Text = string.Empty;
            txt_principleamntshdt_5.Text = string.Empty;
            txt_payafterjudgamtshdt_5.Text = string.Empty;
            txt_deptamntshdt_5.Text = string.Empty;
            txt_lastpaydateshdt_5.Text = string.Empty;

            txt_closeamntshdt_5.Text = string.Empty;
            txt_close6amntshdt_5.Text = string.Empty;
            txt_install6amntshdt_5.Text = string.Empty;
            txt_close12amntshdt_5.Text = string.Empty;
            txt_install12amntshdt_5.Text = string.Empty;
            txt_close24amntshdt_5.Text =
            txt_install24amntshdt_5.Text = string.Empty;



            txt_cardnoshdt_6.Text = string.Empty;
            txt_judgmentamntshdt_6.Text = string.Empty;
            txt_principleamntshdt_6.Text = string.Empty;
            txt_payafterjudgamtshdt_6.Text = string.Empty;
            txt_deptamntshdt_6.Text = string.Empty;
            txt_lastpaydateshdt_6.Text = string.Empty;
            txt_closeamntshdt_6.Text =
            txt_close6amntshdt_6.Text =
            txt_install6amntshdt_6.Text =
            txt_close12amntshdt_6.Text =
            txt_install12amntshdt_6.Text =
            txt_close24amntshdt_6.Text =
            txt_install24amntshdt_6.Text =


            txt_queueno_shdt.Text = string.Empty;
            linkfileopen.Text = string.Empty;
            linkfileopen.Visible = false;

            if (isclearsearch) txt_search_custid_shdt.Text = string.Empty;
        }
        #endregion
        #region Print PDF

        private void doPrintC2TableMeargeFile()
        {
            if (datatabledetailshow == null) { return; }
            if (Path.Exists(txt_path_c2.Text))
            {
                DataRow[] dataselect = datatabledetailshow.Select("IsSelect = 1");
                GlobalFontSettings.FontResolver = new FileFontResolver();
                PdfDocument doc = new PdfDocument();
                List<FestCustom> customDataList = new List<FestCustom>();
                string? customeridfilename = string.Empty;
                if (dataselect.Length > 0)
                {
                    for (int i = 0; i < dataselect.Length; i++)
                    {
                        string? customerid = dataselect[i]["CustomerID"].ToString();
                        string? caseid = dataselect[i]["CaseID"].ToString();
                        if (i == 0) customeridfilename = customerid;
                        if (!(string.IsNullOrEmpty(customerid) || string.IsNullOrEmpty(caseid)))
                        {
                            List<DataCPSMaster> datamasterlist = sqlitedsrv.doGetDataCPSMasterAllByCustomerID(customerid, caseid);
                            List<DataCPSPerson> dataPerson = doConvertDataMasterToCPSPerson(datamasterlist);
                            if (dataPerson.Count > 0)
                            {
                                List<DataCPSCard> cardDataList = doConvertToCardDataCPS(dataPerson[0]);
                                for (int m = 0; m < cardDataList.Count; m++)
                                {
                                    DataCPSCard CPSCall = cardDataList[m];
                                    if (string.IsNullOrEmpty(CPSCall.LedNumber)) CPSCall.LedNumber = txt_queueno_shdt.Text;
                                    if (CPSCall.CustomFlag == "Y")
                                    {
                                        if (m == 0) customDataList = sqlitedsrv.doGetDataCustomWithID(CPSCall.CustomerID ?? "", CPSCall.CaseID ?? "");
                                        if (customDataList.Count > 0) calcsrv.CaluLateData6CardCustom(ref CPSCall, customDataList, m);
                                    }
                                    else
                                    {
                                        calcsrv.CaluLateData6Card(ref CPSCall, setdata);
                                    }
                                    cardDataList[m] = CPSCall;
                                }
                                if (Path.Exists(txt_path_c2.Text))
                                {
                                    reportsrv.doCreteC2TableReport1File(ref doc, cardDataList, setdata, 2);
                                }
                            }

                        }
                    }
                    string file_name = string.Format("C2Table_{0}_{1}.pdf", txt_queueno_shdt.Text, customeridfilename);
                    file_name = file_name.Replace("\n", "").Replace("\r", "").Replace("/", "").Replace(" ", "");
                    string fullfilename = Path.Combine(txt_path_c2.Text, file_name);
                    if (Path.Exists(txt_path_c2.Text))
                    {
                        doc.Save(fullfilename);
                        if (openaftersave.Checked) doOpenPdfFile(fullfilename);
                        linkfileopen.Text = file_name;
                        linkfileopen.Visible = true;
                    }
                }
            }
            else
            {
                MessageBox.Show("กรุณาตรวจสอบ Path เก็บข้อมูล", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion
        #region Control Event
        private void btn_save_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doSaveDataSetting();
            Cursor = Cursors.Default;
            MessageBox.Show("บันทึกสำเร็จ", "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_path_c2_Click(object sender, EventArgs e)
        {
            txt_path_c2.Text = browFolder();
        }
        private void btn_table_cal_Click(object sender, EventArgs e)
        {
            txt_path_table.Text = browFolder();
        }
        private void btn_custid_shdt_Click(object sender, EventArgs e)
        {
            doSearhCPSDataWithCusomerID();
        }
        private void btn_print_c2table_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (!string.IsNullOrEmpty(txt_queueno_shdt.Text))
            {
               doPrintC2TableMeargeFile();

            }
            else
            {
                MessageBox.Show("หมายเลขคิวห้ามว่าง", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txt_queueno_shdt.Focus();
            }

            Cursor = Cursors.Default;
        }

        
        private void linkfileopen_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!string.IsNullOrEmpty(linkfileopen.Text))
            {
                string fllpath = Path.Combine(txt_path_c2.Text, linkfileopen.Text);
                doOpenPdfFile(fllpath);
            }
        }

        private void btn_cleardata_Click(object sender, EventArgs e)
        {
            doClearControl(true);
            datatabledetailshow.Clear();
            dataGridDetailCard.DataSource = datatabledetailshow;
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

        private void doOpenPdfFile(string fullfilename)
        {
            if (Path.Exists(fullfilename))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = fullfilename,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show(string.Format("{0} \r\n ไม่ถูกต้องกรุณาตรวจสอบ", fullfilename), "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void btn_caculate_Click(object sender, EventArgs e)
        {
            doClearControl(false);
            DataRow[] dataselect = datatabledetailshow.Select("IsSelect = 1");
            
            if (dataselect.Length == 1)
            {
                int indexrow = datatabledetailshow.Rows.IndexOf(dataselect[0]);
                if (indexrow >= 0)
                {
                    doCalculateDataShowCalTab(indexrow);
                    tabControl1.SelectedIndex = 1;
                }
            }
            else
            {
                MessageBox.Show("กรูณาเลือกครั้งละ 1 ราย", "เลือก 1 ราย", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void doCalculateDataShowCalTab(int indexrow)
        {
            CARDPRINT = new List<DataCPSCard>();
            string? customerid = Convert.ToString(datatabledetailshow.Rows[indexrow]["CustomerID"] is DBNull ? "" : datatabledetailshow.Rows[indexrow]["CustomerID"]);
            string? caseid = Convert.ToString(datatabledetailshow.Rows[indexrow]["CaseID"] is DBNull ? "" : datatabledetailshow.Rows[indexrow]["CaseID"]);
            if (!(string.IsNullOrEmpty(customerid) || string.IsNullOrEmpty(caseid)))
            {
                List<DataCPSMaster> datamasterlist = sqlitedsrv.doGetDataCPSMasterAllByCustomerID(customerid, caseid);
                List<DataCPSPerson> dataPerson = doConvertDataMasterToCPSPerson(datamasterlist);
                List<FestCustom> customDataList = new List<FestCustom>();
                if (dataPerson.Count > 0)
                {
                    List<DataCPSCard> cardDataList = doConvertToCardDataCPS(dataPerson[0]);
                    for (int m = 0; m < cardDataList.Count; m++)
                    {
                        DataCPSCard CPSCall = cardDataList[m];
                        if (string.IsNullOrEmpty(CPSCall.LedNumber)) CPSCall.LedNumber = txt_queueno_shdt.Text;
                        if (CPSCall.CustomFlag == "Y")
                        {
                            if (m == 0) customDataList = sqlitedsrv.doGetDataCustomWithID(CPSCall.CustomerID ?? "", CPSCall.CaseID ?? "");
                            if (customDataList.Count > 0) calcsrv.CaluLateData6CardCustom(ref CPSCall, customDataList, m);
                        }
                        else
                        {
                            calcsrv.CaluLateData6Card(ref CPSCall, setdata);
                        }
                        cardDataList[m] = CPSCall;
                    }
                    doSetDataShowText(cardDataList);
                    CARDPRINT = cardDataList;
                }
            }
        }
        private void doSetDataShowText(List<DataCPSCard> cardDataList)
        {
            for (int i = 0; i < cardDataList.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        txt_custid_shdt.Text = cardDataList[i].CustomerID;
                        txt_custname_shdt.Text = cardDataList[i].CustomerName;

                        txt_cardnoshdt_1.Text = cardDataList[i].CardNo;
                        txt_judgmentamntshdt_1.Text = cardDataList[i].JudgmentAmnt.ToString("N");
                        txt_principleamntshdt_1.Text = cardDataList[i].PrincipleAmnt == 0 ? "-" : cardDataList[i].PrincipleAmnt.ToString("N");
                        txt_payafterjudgamtshdt_1.Text = cardDataList[i].PayAfterJudgAmt == 0 ? "-" : cardDataList[i].PayAfterJudgAmt.ToString("N");
                        txt_deptamntshdt_1.Text = cardDataList[i].DeptAmnt == 0 ? "-" : cardDataList[i].DeptAmnt.ToString("N");
                        txt_lastpaydateshdt_1.Text = dateHelper.doGetShortDateTHFromDBToPDF(cardDataList[i].LastPayDate ?? "-");
                        txt_closeamntshdt_1.Text = cardDataList[i].AccCloseAmnt == 0 ? "-" : cardDataList[i].AccCloseAmnt.ToString("N");
                        txt_close6amntshdt_1.Text = cardDataList[i].AccClose6Amnt == 0 ? "-" : cardDataList[i].AccClose6Amnt.ToString("N");
                        txt_install6amntshdt_1.Text = cardDataList[i].Installment6Amnt == 0 ? "-" : cardDataList[i].Installment6Amnt.ToString("N");
                        txt_close12amntshdt_1.Text = cardDataList[i].AccClose12Amnt == 0 ? "-" : cardDataList[i].AccClose12Amnt.ToString("N");
                        txt_install12amntshdt_1.Text = cardDataList[i].Installment12Amnt == 0 ? "-" : cardDataList[i].Installment12Amnt.ToString("N");
                        txt_close24amntshdt_1.Text = cardDataList[i].AccClose24Amnt == 0 ? "-" : cardDataList[i].AccClose24Amnt.ToString("N");
                        txt_install24amntshdt_1.Text = cardDataList[i].Installment24Amnt == 0 ? "-" : cardDataList[i].Installment24Amnt.ToString("N");
                        break;
                    case 1:
                        txt_cardnoshdt_2.Text = cardDataList[i].CardNo;
                        txt_judgmentamntshdt_2.Text = cardDataList[i].JudgmentAmnt.ToString("N");
                        txt_principleamntshdt_2.Text = cardDataList[i].PrincipleAmnt == 0 ? "-" : cardDataList[i].PrincipleAmnt.ToString("N");
                        txt_payafterjudgamtshdt_2.Text = cardDataList[i].PayAfterJudgAmt == 0 ? "-" : cardDataList[i].PayAfterJudgAmt.ToString("N");
                        txt_deptamntshdt_2.Text = cardDataList[i].DeptAmnt == 0 ? "-" : cardDataList[i].DeptAmnt.ToString("N");
                        txt_lastpaydateshdt_2.Text = dateHelper.doGetShortDateTHFromDBToPDF(cardDataList[i].LastPayDate ?? "-");
                        txt_closeamntshdt_2.Text = cardDataList[i].AccCloseAmnt == 0 ? "-" : cardDataList[i].AccCloseAmnt.ToString("N");
                        txt_close6amntshdt_2.Text = cardDataList[i].AccClose6Amnt == 0 ? "-" : cardDataList[i].AccClose6Amnt.ToString("N");
                        txt_install6amntshdt_2.Text = cardDataList[i].Installment6Amnt == 0 ? "-" : cardDataList[i].Installment6Amnt.ToString("N");
                        txt_close12amntshdt_2.Text = cardDataList[i].AccClose12Amnt == 0 ? "-" : cardDataList[i].AccClose12Amnt.ToString("N");
                        txt_install12amntshdt_2.Text = cardDataList[i].Installment12Amnt == 0 ? "-" : cardDataList[i].Installment12Amnt.ToString("N");
                        txt_close24amntshdt_2.Text = cardDataList[i].AccClose24Amnt == 0 ? "-" : cardDataList[i].AccClose24Amnt.ToString("N");
                        txt_install24amntshdt_2.Text = cardDataList[i].Installment24Amnt == 0 ? "-" : cardDataList[i].Installment24Amnt.ToString("N");
                        break;
                    case 2:
                        txt_cardnoshdt_3.Text = cardDataList[i].CardNo;
                        txt_judgmentamntshdt_3.Text = cardDataList[i].JudgmentAmnt.ToString("N");
                        txt_principleamntshdt_3.Text = cardDataList[i].PrincipleAmnt == 0 ? "-" : cardDataList[i].PrincipleAmnt.ToString("N");
                        txt_payafterjudgamtshdt_3.Text = cardDataList[i].PayAfterJudgAmt == 0 ? "-" : cardDataList[i].PayAfterJudgAmt.ToString("N");
                        txt_deptamntshdt_3.Text = cardDataList[i].DeptAmnt == 0 ? "-" : cardDataList[i].DeptAmnt.ToString("N");
                        txt_lastpaydateshdt_3.Text = dateHelper.doGetShortDateTHFromDBToPDF(cardDataList[i].LastPayDate ?? "-");
                        txt_closeamntshdt_3.Text = cardDataList[i].AccCloseAmnt == 0 ? "-" : cardDataList[i].AccCloseAmnt.ToString("N");
                        txt_close6amntshdt_3.Text = cardDataList[i].AccClose6Amnt == 0 ? "-" : cardDataList[i].AccClose6Amnt.ToString("N");
                        txt_install6amntshdt_3.Text = cardDataList[i].Installment6Amnt == 0 ? "-" : cardDataList[i].Installment6Amnt.ToString("N");
                        txt_close12amntshdt_3.Text = cardDataList[i].AccClose12Amnt == 0 ? "-" : cardDataList[i].AccClose12Amnt.ToString("N");
                        txt_install12amntshdt_3.Text = cardDataList[i].Installment12Amnt == 0 ? "-" : cardDataList[i].Installment12Amnt.ToString("N");
                        txt_close24amntshdt_3.Text = cardDataList[i].AccClose24Amnt == 0 ? "-" : cardDataList[i].AccClose24Amnt.ToString("N");
                        txt_install24amntshdt_3.Text = cardDataList[i].Installment24Amnt == 0 ? "-" : cardDataList[i].Installment24Amnt.ToString("N");
                        break;
                    case 3:
                        txt_cardnoshdt_4.Text = cardDataList[i].CardNo;
                        txt_judgmentamntshdt_4.Text = cardDataList[i].JudgmentAmnt.ToString("N");
                        txt_principleamntshdt_4.Text = cardDataList[i].PrincipleAmnt == 0 ? "-" : cardDataList[i].PrincipleAmnt.ToString("N");
                        txt_payafterjudgamtshdt_4.Text = cardDataList[i].PayAfterJudgAmt == 0 ? "-" : cardDataList[i].PayAfterJudgAmt.ToString("N");
                        txt_deptamntshdt_4.Text = cardDataList[i].DeptAmnt == 0 ? "-" : cardDataList[i].DeptAmnt.ToString("N");
                        txt_lastpaydateshdt_4.Text = dateHelper.doGetShortDateTHFromDBToPDF(cardDataList[i].LastPayDate ?? "-");
                        txt_closeamntshdt_4.Text = cardDataList[i].AccCloseAmnt == 0 ? "-" : cardDataList[i].AccCloseAmnt.ToString("N");
                        txt_close6amntshdt_4.Text = cardDataList[i].AccClose6Amnt == 0 ? "-" : cardDataList[i].AccClose6Amnt.ToString("N");
                        txt_install6amntshdt_4.Text = cardDataList[i].Installment6Amnt == 0 ? "-" : cardDataList[i].Installment6Amnt.ToString("N");
                        txt_close12amntshdt_4.Text = cardDataList[i].AccClose12Amnt == 0 ? "-" : cardDataList[i].AccClose12Amnt.ToString("N");
                        txt_install12amntshdt_4.Text = cardDataList[i].Installment12Amnt == 0 ? "-" : cardDataList[i].Installment12Amnt.ToString("N");
                        txt_close24amntshdt_4.Text = cardDataList[i].AccClose24Amnt == 0 ? "-" : cardDataList[i].AccClose24Amnt.ToString("N");
                        txt_install24amntshdt_4.Text = cardDataList[i].Installment24Amnt == 0 ? "-" : cardDataList[i].Installment24Amnt.ToString("N");
                        break;
                    case 4:
                        txt_cardnoshdt_5.Text = cardDataList[i].CardNo;
                        txt_judgmentamntshdt_5.Text = cardDataList[i].JudgmentAmnt.ToString("N");
                        txt_principleamntshdt_5.Text = cardDataList[i].PrincipleAmnt == 0 ? "-" : cardDataList[i].PrincipleAmnt.ToString("N");
                        txt_payafterjudgamtshdt_5.Text = cardDataList[i].PayAfterJudgAmt == 0 ? "-" : cardDataList[i].PayAfterJudgAmt.ToString("N");
                        txt_deptamntshdt_5.Text = cardDataList[i].DeptAmnt == 0 ? "-" : cardDataList[i].DeptAmnt.ToString("N");
                        txt_lastpaydateshdt_5.Text = dateHelper.doGetShortDateTHFromDBToPDF(cardDataList[i].LastPayDate ?? "-");
                        txt_closeamntshdt_5.Text = cardDataList[i].AccCloseAmnt == 0 ? "-" : cardDataList[i].AccCloseAmnt.ToString("N");
                        txt_close6amntshdt_5.Text = cardDataList[i].AccClose6Amnt == 0 ? "-" : cardDataList[i].AccClose6Amnt.ToString("N");
                        txt_install6amntshdt_5.Text = cardDataList[i].Installment6Amnt == 0 ? "-" : cardDataList[i].Installment6Amnt.ToString("N");
                        txt_close12amntshdt_5.Text = cardDataList[i].AccClose12Amnt == 0 ? "-" : cardDataList[i].AccClose12Amnt.ToString("N");
                        txt_install12amntshdt_5.Text = cardDataList[i].Installment12Amnt == 0 ? "-" : cardDataList[i].Installment12Amnt.ToString("N");
                        txt_close24amntshdt_5.Text = cardDataList[i].AccClose24Amnt == 0 ? "-" : cardDataList[i].AccClose24Amnt.ToString("N");
                        txt_install24amntshdt_5.Text = cardDataList[i].Installment24Amnt == 0 ? "-" : cardDataList[i].Installment24Amnt.ToString("N");
                        break;
                    case 5:
                        txt_cardnoshdt_6.Text = cardDataList[i].CardNo;
                        txt_judgmentamntshdt_6.Text = cardDataList[i].JudgmentAmnt.ToString("N");
                        txt_principleamntshdt_6.Text = cardDataList[i].PrincipleAmnt == 0 ? "-" : cardDataList[i].PrincipleAmnt.ToString("N");
                        txt_payafterjudgamtshdt_6.Text = cardDataList[i].PayAfterJudgAmt == 0 ? "-" : cardDataList[i].PayAfterJudgAmt.ToString("N");
                        txt_deptamntshdt_6.Text = cardDataList[i].DeptAmnt == 0 ? "-" : cardDataList[i].DeptAmnt.ToString("N");
                        txt_lastpaydateshdt_6.Text = dateHelper.doGetShortDateTHFromDBToPDF(cardDataList[i].LastPayDate ?? "-");
                        txt_closeamntshdt_6.Text = cardDataList[i].AccCloseAmnt == 0 ? "-" : cardDataList[i].AccCloseAmnt.ToString("N");
                        txt_close6amntshdt_6.Text = cardDataList[i].AccClose6Amnt == 0 ? "-" : cardDataList[i].AccClose6Amnt.ToString("N");
                        txt_install6amntshdt_6.Text = cardDataList[i].Installment6Amnt == 0 ? "-" : cardDataList[i].Installment6Amnt.ToString("N");
                        txt_close12amntshdt_6.Text = cardDataList[i].AccClose12Amnt == 0 ? "-" : cardDataList[i].AccClose12Amnt.ToString("N");
                        txt_install12amntshdt_6.Text = cardDataList[i].Installment12Amnt == 0 ? "-" : cardDataList[i].Installment12Amnt.ToString("N");
                        txt_close24amntshdt_6.Text = cardDataList[i].AccClose24Amnt == 0 ? "-" : cardDataList[i].AccClose24Amnt.ToString("N");
                        txt_install24amntshdt_6.Text = cardDataList[i].Installment24Amnt == 0 ? "-" : cardDataList[i].Installment24Amnt.ToString("N");
                        break;

                }

            }
        }
        private void btn_print_c2table_Click_1(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (!string.IsNullOrEmpty(txt_queueno_shdt.Text))
            {
                string customeridfilename = CARDPRINT[0].CustomerID ?? "";
                GlobalFontSettings.FontResolver = new FileFontResolver();
                PdfDocument doc = new PdfDocument();
                if (Path.Exists(txt_path_c2.Text))
                {
                    doCreatePrintDataCard();
                    reportsrv.doCreteC2TableReport1File(ref doc, CARDPRINT, setdata, 2);
                }
                string file_name = string.Format("C2Table_{0}_{1}.pdf", txt_queueno_shdt.Text, customeridfilename);
                file_name = file_name.Replace("\n", "").Replace("\r", "").Replace("/", "").Replace(" ", "");
                string fullfilename = Path.Combine(txt_path_c2.Text, file_name);
                if (Path.Exists(txt_path_c2.Text))
                {
                    doc.Save(fullfilename);
                    if (openaftersave.Checked) doOpenPdfFile(fullfilename);
                    linkfileopen.Text = file_name;
                    linkfileopen.Visible = true;
                }
                doClearControl(false);
                tabControl1.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("หมายเลขคิวห้ามว่าง", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txt_queueno_shdt.Focus();
            }

            Cursor = Cursors.Default;
            
        }
        private void doCreatePrintDataCard()
        {
            for (int i = 0; i < CARDPRINT.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        CARDPRINT[i].LedNumber = string.IsNullOrEmpty(CARDPRINT[i].LedNumber) ? txt_queueno_shdt.Text: CARDPRINT[i].LedNumber;
                        CARDPRINT[i].AccCloseAmnt = Convert.ToDouble(txt_closeamntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_closeamntshdt_1.Text) ? "0.0" : txt_closeamntshdt_1.Text);
                        CARDPRINT[i].AccClose6Amnt = Convert.ToDouble(txt_close6amntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_close6amntshdt_1.Text) ? "0.0" : txt_close6amntshdt_1.Text);
                        CARDPRINT[i].Installment6Amnt = Convert.ToDouble(txt_install6amntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_install6amntshdt_1.Text) ? "0.0" : txt_install6amntshdt_1.Text);
                        CARDPRINT[i].AccClose12Amnt = Convert.ToDouble(txt_close12amntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_close12amntshdt_1.Text) ? "0.0" : txt_close12amntshdt_1.Text);
                        CARDPRINT[i].Installment12Amnt = Convert.ToDouble(txt_install12amntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_install12amntshdt_1.Text) ? "0.0" : txt_install12amntshdt_1.Text);
                        CARDPRINT[i].AccClose24Amnt = Convert.ToDouble(txt_close24amntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_close24amntshdt_1.Text) ? "0.0" : txt_close24amntshdt_1.Text);
                        CARDPRINT[i].Installment24Amnt = Convert.ToDouble(txt_install24amntshdt_1.Text == "_" || string.IsNullOrEmpty(txt_install24amntshdt_1.Text) ? "0.0" : txt_install24amntshdt_1.Text);
                        break;
                    case 1:
                        CARDPRINT[i].AccCloseAmnt = Convert.ToDouble(txt_closeamntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_closeamntshdt_2.Text) ? "0.0" : txt_closeamntshdt_2.Text);
                        CARDPRINT[i].AccClose6Amnt = Convert.ToDouble(txt_close6amntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_close6amntshdt_2.Text) ? "0.0" : txt_close6amntshdt_2.Text);
                        CARDPRINT[i].Installment6Amnt = Convert.ToDouble(txt_install6amntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_install6amntshdt_2.Text) ? "0.0" : txt_install6amntshdt_2.Text);
                        CARDPRINT[i].AccClose12Amnt = Convert.ToDouble(txt_close12amntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_close12amntshdt_2.Text) ? "0.0" : txt_close12amntshdt_2.Text);
                        CARDPRINT[i].Installment12Amnt = Convert.ToDouble(txt_install12amntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_install12amntshdt_2.Text) ? "0.0" : txt_install12amntshdt_2.Text);
                        CARDPRINT[i].AccClose24Amnt = Convert.ToDouble(txt_close24amntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_close24amntshdt_2.Text) ? "0.0" : txt_close24amntshdt_2.Text);
                        CARDPRINT[i].Installment24Amnt = Convert.ToDouble(txt_install24amntshdt_2.Text == "_" || string.IsNullOrEmpty(txt_install24amntshdt_2.Text) ? "0.0" : txt_install24amntshdt_2.Text);
                        break;
                    case 2:
                        CARDPRINT[i].AccCloseAmnt = Convert.ToDouble(txt_closeamntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_closeamntshdt_3.Text) ? "0.0" : txt_closeamntshdt_3.Text);
                        CARDPRINT[i].AccClose6Amnt = Convert.ToDouble(txt_close6amntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_close6amntshdt_3.Text) ? "0.0" : txt_close6amntshdt_3.Text);
                        CARDPRINT[i].Installment6Amnt = Convert.ToDouble(txt_install6amntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_install6amntshdt_3.Text) ? "0.0" : txt_install6amntshdt_3.Text);
                        CARDPRINT[i].AccClose12Amnt = Convert.ToDouble(txt_close12amntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_close12amntshdt_3.Text) ? "0.0" : txt_close12amntshdt_3.Text);
                        CARDPRINT[i].Installment12Amnt = Convert.ToDouble(txt_install12amntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_install12amntshdt_3.Text) ? "0.0" : txt_install12amntshdt_3.Text);
                        CARDPRINT[i].AccClose24Amnt = Convert.ToDouble(txt_close24amntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_close24amntshdt_3.Text) ? "0.0" : txt_close24amntshdt_3.Text);
                        CARDPRINT[i].Installment24Amnt = Convert.ToDouble(txt_install24amntshdt_3.Text == "_" || string.IsNullOrEmpty(txt_install24amntshdt_3.Text) ? "0.0" : txt_install24amntshdt_3.Text);
                        break;
                    case 3:
                        CARDPRINT[i].AccCloseAmnt = Convert.ToDouble(txt_closeamntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_closeamntshdt_4.Text) ? "0.0" : txt_closeamntshdt_4.Text);
                        CARDPRINT[i].AccClose6Amnt = Convert.ToDouble(txt_close6amntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_close6amntshdt_4.Text) ? "0.0" : txt_close6amntshdt_4.Text);
                        CARDPRINT[i].Installment6Amnt = Convert.ToDouble(txt_install6amntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_install6amntshdt_4.Text) ? "0.0" : txt_install6amntshdt_4.Text);
                        CARDPRINT[i].AccClose12Amnt = Convert.ToDouble(txt_close12amntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_close12amntshdt_4.Text) ? "0.0" : txt_close12amntshdt_4.Text);
                        CARDPRINT[i].Installment12Amnt = Convert.ToDouble(txt_install12amntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_install12amntshdt_4.Text) ? "0.0" : txt_install12amntshdt_4.Text);
                        CARDPRINT[i].AccClose24Amnt = Convert.ToDouble(txt_close24amntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_close24amntshdt_4.Text) ? "0.0" : txt_close24amntshdt_4.Text);
                        CARDPRINT[i].Installment24Amnt = Convert.ToDouble(txt_install24amntshdt_4.Text == "_" || string.IsNullOrEmpty(txt_install24amntshdt_4.Text) ? "0.0" : txt_install24amntshdt_4.Text);
                        break;
                    case 4:
                        CARDPRINT[i].AccCloseAmnt = Convert.ToDouble(txt_closeamntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_closeamntshdt_5.Text) ? "0.0" : txt_closeamntshdt_5.Text);
                        CARDPRINT[i].AccClose6Amnt = Convert.ToDouble(txt_close6amntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_close6amntshdt_5.Text) ? "0.0" : txt_close6amntshdt_5.Text);
                        CARDPRINT[i].Installment6Amnt = Convert.ToDouble(txt_install6amntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_install6amntshdt_5.Text) ? "0.0" : txt_install6amntshdt_5.Text);
                        CARDPRINT[i].AccClose12Amnt = Convert.ToDouble(txt_close12amntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_close12amntshdt_5.Text) ? "0.0" : txt_close12amntshdt_5.Text);
                        CARDPRINT[i].Installment12Amnt = Convert.ToDouble(txt_install12amntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_install12amntshdt_5.Text) ? "0.0" : txt_install12amntshdt_5.Text);
                        CARDPRINT[i].AccClose24Amnt = Convert.ToDouble(txt_close24amntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_close24amntshdt_5.Text) ? "0.0" : txt_close24amntshdt_5.Text);
                        CARDPRINT[i].Installment24Amnt = Convert.ToDouble(txt_install24amntshdt_5.Text == "_" || string.IsNullOrEmpty(txt_install24amntshdt_5.Text) ? "0.0" : txt_install24amntshdt_5.Text);
                        break;
                    case 5:
                        CARDPRINT[i].AccCloseAmnt = Convert.ToDouble(txt_closeamntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_closeamntshdt_6.Text) ? "0.0" : txt_closeamntshdt_6.Text);
                        CARDPRINT[i].AccClose6Amnt = Convert.ToDouble(txt_close6amntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_close6amntshdt_6.Text) ? "0.0" : txt_close6amntshdt_6.Text);
                        CARDPRINT[i].Installment6Amnt = Convert.ToDouble(txt_install6amntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_install6amntshdt_6.Text) ? "0.0" : txt_install6amntshdt_6.Text);
                        CARDPRINT[i].AccClose12Amnt = Convert.ToDouble(txt_close12amntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_close12amntshdt_6.Text) ? "0.0" : txt_close12amntshdt_6.Text);
                        CARDPRINT[i].Installment12Amnt = Convert.ToDouble(txt_install12amntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_install12amntshdt_6.Text) ? "0.0" : txt_install12amntshdt_6.Text);
                        CARDPRINT[i].AccClose24Amnt = Convert.ToDouble(txt_close24amntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_close24amntshdt_6.Text) ? "0.0" : txt_close24amntshdt_6.Text);
                        CARDPRINT[i].Installment24Amnt = Convert.ToDouble(txt_install24amntshdt_6.Text == "_" || string.IsNullOrEmpty(txt_install24amntshdt_6.Text) ? "0.0" : txt_install24amntshdt_6.Text);
                        break;

                }

            }
        }
        private void dataGridDetailCard_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) { return; }
            string columnName = dataGridDetailCard.Columns[e.ColumnIndex].Name;
            if (columnName == "IsSelect")
            {
                bool isselect = Convert.ToBoolean(datatabledetailshow.Rows[e.RowIndex]["IsSelect"] is DBNull ? false : datatabledetailshow.Rows[e.RowIndex]["IsSelect"]);
                datatabledetailshow.Rows[e.RowIndex]["IsSelect"] = !isselect;
            }
        }
    }
}