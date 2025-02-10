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
    public partial class frmUserReport : frmBaseWindow
    {
        #region Member
        SettingData setdata = new SettingData();
        dataTableService dtService = new dataTableService();
        reportService reportsrv = new reportService();
        calculateDataService calcsrv = new calculateDataService();
        sqliteDataService sqlitedsrv = new sqliteDataService();
        dateTimeHelper dateHelper = new dateTimeHelper();
        DataTable datatabledetailshow;
        bool is_kzas;

        string typerange = string.Empty;
        #endregion
        #region Constructor
        public frmUserReport()
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

            dataGridDetailCard.CellClick += new DataGridViewCellEventHandler(DataGridView_CellClick);
        }
        private void DataGridView_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            int rowindex = e.RowIndex;
            if (rowindex >= 0) // Ensure that the click is on a valid row
            {
                if (sender != null)
                {
                    if (rowindex >= 0)
                    {
                        DataGridView dataGridView = (DataGridView)sender;
                        DataGridViewRow selectedRow = dataGridView.Rows[rowindex];
                        string? customerid = selectedRow.Cells["CustomerID"].Value is null ? "" : selectedRow.Cells["CustomerID"].Value.ToString();
                        string? caseid = selectedRow.Cells["CaseID"].Value is null ? "" : selectedRow.Cells["CaseID"].Value.ToString();
                        if (!(string.IsNullOrEmpty(customerid)||string.IsNullOrEmpty(caseid)))
                        {
                            List<DataCPSMaster> datamasterlist = sqlitedsrv.doGetDataCPSMasterAllByCustomerID(customerid, caseid);
                            if (datamasterlist != null)
                            {
                                List<DataCPSPerson> datacpsperson = doConvertDataMasterToCPSPerson(datamasterlist);
                                doSetdataCPSShow(datacpsperson);
                            }
                        }

                    }
                }

            }
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
                    string customid = dataperson.CustomerID??"";
                    string caseid = dataperson.CaseID ?? "";
                    if (!string.IsNullOrEmpty(customid))
                    {
                        List<FestCustom> datacustom = sqlitedsrv.doGetDataCustomWithID(customid, caseid);
                        if(datacustom.Count > 0)
                        {
                            
                            string cardno1 = datacustom[0].CardNo1??"";
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
        private void doSetdataCPSShow(List<DataCPSPerson> data6CardList)
        {
            string? customstatus = Convert.ToString(data6CardList[0].CustomFlag is null ? "N" : data6CardList[0].CustomFlag);           
            if (customstatus == "Y")
            {
                #region Set Data 1 Account
                    string customerid = data6CardList[0].CustomerID ?? "";
                    string caseid = data6CardList[0].CaseID ?? "";

                doSetDispalyCustom(customstatus);
                    txt_custid_shdt.Text = data6CardList[0].CustomerID;
                    txt_custname_shdt.Text = data6CardList[0].CustomerName;

                    string? lednumber = Convert.ToString(data6CardList[0].LedNumber is null ? "" : data6CardList[0].LedNumber);

                    List<FestCustom> data6CardListCustom = sqlitedsrv.doGetDataCustomWithID(customerid, caseid);
                    if (data6CardListCustom == null) return;
                    if (data6CardListCustom.Count > 0) 
                    {
                        string cardno1 = data6CardListCustom[0].CardNo1 ?? "";
                        if (!string.IsNullOrEmpty(cardno1))
                        {
                            int indexcard1 = findindexCardNo(data6CardList[0],cardno1);
                            if (indexcard1 >= 0)
                            {
                                DataCPSPerson datartn = doGetDataCardWithIndex(indexcard1, data6CardList[0]);
                                txt_cardnoshdt_1.Text = cardno1;
                                txt_judgmentamntshdt_1.Text = datartn.JudgmentAmnt1.ToString("N");
                                txt_principleamntshdt_1.Text = datartn.PrincipleAmnt1 == 0 ? "-" : datartn.PrincipleAmnt1.ToString("N");
                                txt_payafterjudgamtshdt_1.Text = datartn.PayAfterJudgAmt1 == 0 ? "-" : datartn.PayAfterJudgAmt1.ToString("N");
                                txt_deptamntshdt_1.Text = datartn.DeptAmnt1 == 0 ? "-" : datartn.DeptAmnt1.ToString("N");
                                txt_lastpaydateshdt_1.Text = dateHelper.doGetShortDateTHFromDBToPDF(datartn.LastPayDate1 ?? "-");
                            }
                        }

                        string cardno2 = data6CardListCustom[0].CardNo2 ?? "";
                        if (!string.IsNullOrEmpty(cardno2))
                        {
                            int indexcard2 = findindexCardNo(data6CardList[0], cardno2);
                            if (indexcard2 >= 0)
                            {
                                DataCPSPerson datartn = doGetDataCardWithIndex(indexcard2, data6CardList[0]);
                                txt_cardnoshdt_2.Text = cardno2;
                                txt_judgmentamntshdt_2.Text = datartn.JudgmentAmnt1.ToString("N");
                                txt_principleamntshdt_2.Text = datartn.PrincipleAmnt1 == 0 ? "-" : datartn.PrincipleAmnt1.ToString("N");
                                txt_payafterjudgamtshdt_2.Text = datartn.PayAfterJudgAmt1 == 0 ? "-" : datartn.PayAfterJudgAmt1.ToString("N");
                                txt_deptamntshdt_2.Text = datartn.DeptAmnt1 == 0 ? "-" : datartn.DeptAmnt1.ToString("N");
                                txt_lastpaydateshdt_2.Text = dateHelper.doGetShortDateTHFromDBToPDF(datartn.LastPayDate1 ?? "-");
                            }
                        }

                        string cardno3 = data6CardListCustom[0].CardNo3 ?? "";
                        if (!string.IsNullOrEmpty(cardno3))
                        {
                            int indexcard3 = findindexCardNo(data6CardList[0], cardno3);
                            if (indexcard3 >= 0)
                            {
                                DataCPSPerson datartn = doGetDataCardWithIndex(indexcard3, data6CardList[0]);
                                txt_cardnoshdt_3.Text = cardno3;
                                txt_judgmentamntshdt_3.Text = datartn.JudgmentAmnt1.ToString("N");
                                txt_principleamntshdt_3.Text = datartn.PrincipleAmnt1 == 0 ? "-" : datartn.PrincipleAmnt1.ToString("N");
                                txt_payafterjudgamtshdt_3.Text = datartn.PayAfterJudgAmt1 == 0 ? "-" : datartn.PayAfterJudgAmt1.ToString("N");
                                txt_deptamntshdt_3.Text = datartn.DeptAmnt1 == 0 ? "-" : datartn.DeptAmnt1.ToString("N");
                                txt_lastpaydateshdt_3.Text = dateHelper.doGetShortDateTHFromDBToPDF(datartn.LastPayDate1 ?? "-");
                            }
                        }

                        string cardno4 = data6CardListCustom[0].CardNo4 ?? "";
                        if (!string.IsNullOrEmpty(cardno4))
                        {
                            int indexcard4 = findindexCardNo(data6CardList[0], cardno4);
                            if (indexcard4 >= 0)
                            {
                                DataCPSPerson datartn = doGetDataCardWithIndex(indexcard4, data6CardList[0]);
                                txt_cardnoshdt_4.Text = cardno4;
                                txt_judgmentamntshdt_4.Text = datartn.JudgmentAmnt1.ToString("N");
                                txt_principleamntshdt_4.Text = datartn.PrincipleAmnt1 == 0 ? "-" : datartn.PrincipleAmnt1.ToString("N");
                                txt_payafterjudgamtshdt_4.Text = datartn.PayAfterJudgAmt1 == 0 ? "-" : datartn.PayAfterJudgAmt1.ToString("N");
                                txt_deptamntshdt_4.Text = datartn.DeptAmnt1 == 0 ? "-" : datartn.DeptAmnt1.ToString("N");
                                txt_lastpaydateshdt_4.Text = dateHelper.doGetShortDateTHFromDBToPDF(datartn.LastPayDate1 ?? "-");
                            }
                        }

                        string cardno5 = data6CardListCustom[0].CardNo5 ?? "";
                        if (!string.IsNullOrEmpty(cardno5))
                        {
                            int indexcard5 = findindexCardNo(data6CardList[0], cardno5);
                            if (indexcard5 >= 0)
                            {
                                DataCPSPerson datartn = doGetDataCardWithIndex(indexcard5, data6CardList[0]);
                                txt_cardnoshdt_5.Text = cardno5;
                                txt_judgmentamntshdt_5.Text = datartn.JudgmentAmnt1.ToString("N");
                                txt_principleamntshdt_5.Text = datartn.PrincipleAmnt1 == 0 ? "-" : datartn.PrincipleAmnt1.ToString("N");
                                txt_payafterjudgamtshdt_5.Text = datartn.PayAfterJudgAmt1 == 0 ? "-" : datartn.PayAfterJudgAmt1.ToString("N");
                                txt_deptamntshdt_5.Text = datartn.DeptAmnt1 == 0 ? "-" : datartn.DeptAmnt1.ToString("N");
                                txt_lastpaydateshdt_5.Text = dateHelper.doGetShortDateTHFromDBToPDF(datartn.LastPayDate1 ?? "-");
                            }
                        }

                        string cardno6 = data6CardListCustom[0].CardNo6 ?? "";
                        if (!string.IsNullOrEmpty(cardno6))
                        {
                            int indexcard6 = findindexCardNo(data6CardList[0], cardno6);
                            if (indexcard6 >= 0)
                            {
                                DataCPSPerson datartn = doGetDataCardWithIndex(indexcard6, data6CardList[0]);
                                txt_cardnoshdt_6.Text = cardno6;
                                txt_judgmentamntshdt_6.Text = datartn.JudgmentAmnt1.ToString("N");
                                txt_principleamntshdt_6.Text = datartn.PrincipleAmnt1 == 0 ? "-" : datartn.PrincipleAmnt1.ToString("N");
                                txt_payafterjudgamtshdt_6.Text = datartn.PayAfterJudgAmt1 == 0 ? "-" : datartn.PayAfterJudgAmt1.ToString("N");
                                txt_deptamntshdt_6.Text = datartn.DeptAmnt1 == 0 ? "-" : datartn.DeptAmnt1.ToString("N");
                                txt_lastpaydateshdt_6.Text = dateHelper.doGetShortDateTHFromDBToPDF(datartn.LastPayDate1 ?? "-");
                            }
                        }                    
                    #endregion
                }
            }
            else
            {
                #region Set Data 1 Account
                txt_custid_shdt.Text = data6CardList[0].CustomerID;
                txt_custname_shdt.Text = data6CardList[0].CustomerName;

                string? lednumber = Convert.ToString(data6CardList[0].LedNumber is null ? "" : data6CardList[0].LedNumber);

                txt_cardnoshdt_1.Text = data6CardList[0].CardNo1;
                txt_judgmentamntshdt_1.Text = data6CardList[0].JudgmentAmnt1.ToString("N");
                txt_principleamntshdt_1.Text = data6CardList[0].PrincipleAmnt1 == 0 ? "-" : data6CardList[0].PrincipleAmnt1.ToString("N");
                txt_payafterjudgamtshdt_1.Text = data6CardList[0].PayAfterJudgAmt1 == 0 ? "-" : data6CardList[0].PayAfterJudgAmt1.ToString("N");
                txt_deptamntshdt_1.Text = data6CardList[0].DeptAmnt1 == 0 ? "-" : data6CardList[0].DeptAmnt1.ToString("N");
                txt_lastpaydateshdt_1.Text = dateHelper.doGetShortDateTHFromDBToPDF(data6CardList[0].LastPayDate1 ?? "-");

                txt_cardnoshdt_2.Text = data6CardList[0].CardNo2;
                txt_judgmentamntshdt_2.Text = data6CardList[0].JudgmentAmnt2.ToString("N");
                txt_principleamntshdt_2.Text = data6CardList[0].PrincipleAmnt2 == 0 ? "-" : data6CardList[0].PrincipleAmnt2.ToString("N");
                txt_payafterjudgamtshdt_2.Text = data6CardList[0].PayAfterJudgAmt2 == 0 ? "-" : data6CardList[0].PayAfterJudgAmt2.ToString("N");
                txt_deptamntshdt_2.Text = data6CardList[0].DeptAmnt2 == 0 ? "-" : data6CardList[0].DeptAmnt2.ToString("N");
                txt_lastpaydateshdt_2.Text = dateHelper.doGetShortDateTHFromDBToPDF(data6CardList[0].LastPayDate2 ?? "-");

                txt_cardnoshdt_3.Text = data6CardList[0].CardNo3;
                txt_judgmentamntshdt_3.Text = data6CardList[0].JudgmentAmnt3.ToString("N");
                txt_principleamntshdt_3.Text = data6CardList[0].PrincipleAmnt3 == 0 ? "-" : data6CardList[0].PrincipleAmnt3.ToString("N");
                txt_payafterjudgamtshdt_3.Text = data6CardList[0].PayAfterJudgAmt3 == 0 ? "-" : data6CardList[0].PayAfterJudgAmt3.ToString("N");
                txt_deptamntshdt_3.Text = data6CardList[0].DeptAmnt3 == 0 ? "-" : data6CardList[0].DeptAmnt3.ToString("N");
                txt_lastpaydateshdt_3.Text = dateHelper.doGetShortDateTHFromDBToPDF(data6CardList[0].LastPayDate3 ?? "-");


                txt_cardnoshdt_4.Text = data6CardList[0].CardNo4;
                txt_judgmentamntshdt_4.Text = data6CardList[0].JudgmentAmnt4.ToString("N");
                txt_principleamntshdt_4.Text = data6CardList[0].PrincipleAmnt4 == 0 ? "-" : data6CardList[0].PrincipleAmnt4.ToString("N");
                txt_payafterjudgamtshdt_4.Text = data6CardList[0].PayAfterJudgAmt4 == 0 ? "-" : data6CardList[0].PayAfterJudgAmt4.ToString("N");
                txt_deptamntshdt_4.Text = data6CardList[0].DeptAmnt4 == 0 ? "-" : data6CardList[0].DeptAmnt4.ToString("N");
                txt_lastpaydateshdt_4.Text = dateHelper.doGetShortDateTHFromDBToPDF(data6CardList[0].LastPayDate4 ?? "-");

                txt_cardnoshdt_5.Text = data6CardList[0].CardNo5;
                txt_judgmentamntshdt_5.Text = data6CardList[0].JudgmentAmnt5.ToString("N");
                txt_principleamntshdt_5.Text = data6CardList[0].PrincipleAmnt5 == 0 ? "-" : data6CardList[0].PrincipleAmnt5.ToString("N");
                txt_payafterjudgamtshdt_5.Text = data6CardList[0].PayAfterJudgAmt5 == 0 ? "-" : data6CardList[0].PayAfterJudgAmt5.ToString("N");
                txt_deptamntshdt_5.Text = data6CardList[0].DeptAmnt5 == 0 ? "-" : data6CardList[0].DeptAmnt5.ToString("N");
                txt_lastpaydateshdt_5.Text = dateHelper.doGetShortDateTHFromDBToPDF(data6CardList[0].LastPayDate5 ?? "-");

                txt_cardnoshdt_6.Text = data6CardList[0].CardNo6;
                txt_judgmentamntshdt_6.Text = data6CardList[0].JudgmentAmnt6.ToString("N");
                txt_principleamntshdt_6.Text = data6CardList[0].PrincipleAmnt6 == 0 ? "-" : data6CardList[0].PrincipleAmnt6.ToString("N");
                txt_payafterjudgamtshdt_6.Text = data6CardList[0].PayAfterJudgAmt6 == 0 ? "-" : data6CardList[0].PayAfterJudgAmt6.ToString("N");
                txt_deptamntshdt_6.Text = data6CardList[0].DeptAmnt6 == 0 ? "-" : data6CardList[0].DeptAmnt6.ToString("N");
                txt_lastpaydateshdt_6.Text = dateHelper.doGetShortDateTHFromDBToPDF(data6CardList[0].LastPayDate6 ?? "-");
                #endregion
            }
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
                    case 6:
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

        private int findindexCardNo(DataCPSPerson dataCard,string cardno)
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
            DataCPSPerson  datacpscard = new DataCPSPerson();
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
                is_kzas =false;
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
                            datadt["IsSelect"] = true;
                            datadt["CustomerID"] = dataMasterIDList[i].CustomerID;
                            datadt["CustomerName"] = dataMasterIDList[i].CustomerName;
                            datadt["CardStatus"] = dataMasterIDList[i].CardStatus;
                            datadt["LegalStatus"] = dataMasterIDList[i].LegalStatus;
                            datadt["LegalExecDate"] = dataMasterIDList[i].LegalExecDate;
                            datadt["LegalExecRemark"] = dataMasterIDList[i].LegalExecRemark;
                            datadt["CaseID"] = dataMasterIDList[i].CaseID;
                            datatabledetailshow.Rows.Add(datadt);
                        }
                        if (is_kzas) chk_calculate_add.Checked = true;
                    }
                    doSortDataTable(ref datatabledetailshow, "CustomerID ASC, CaseID ASC");
                }
                else
                {
                    MessageBox.Show("ไม่พบข้อมูลลูกค้า ในระบบ","ไม่พบข้อมูล",MessageBoxButtons.OK,MessageBoxIcon.Warning);                  
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

            txt_cardnoshdt_2.Text = string.Empty;
            txt_judgmentamntshdt_2.Text = string.Empty;
            txt_principleamntshdt_2.Text = string.Empty;
            txt_payafterjudgamtshdt_2.Text = string.Empty;
            txt_deptamntshdt_2.Text = string.Empty;
            txt_lastpaydateshdt_2.Text = string.Empty;

            txt_cardnoshdt_3.Text = string.Empty;
            txt_judgmentamntshdt_3.Text = string.Empty;
            txt_principleamntshdt_3.Text = string.Empty;
            txt_payafterjudgamtshdt_3.Text = string.Empty;
            txt_deptamntshdt_3.Text = string.Empty;
            txt_lastpaydateshdt_3.Text = string.Empty;

            txt_cardnoshdt_4.Text = string.Empty;
            txt_judgmentamntshdt_4.Text = string.Empty;
            txt_principleamntshdt_4.Text = string.Empty;
            txt_payafterjudgamtshdt_4.Text = string.Empty;
            txt_deptamntshdt_4.Text = string.Empty;
            txt_lastpaydateshdt_4.Text = string.Empty;

            txt_cardnoshdt_5.Text = string.Empty;
            txt_judgmentamntshdt_5.Text = string.Empty;
            txt_principleamntshdt_5.Text = string.Empty;
            txt_payafterjudgamtshdt_5.Text = string.Empty;
            txt_deptamntshdt_5.Text = string.Empty;
            txt_lastpaydateshdt_5.Text = string.Empty;

            txt_cardnoshdt_6.Text = string.Empty;
            txt_judgmentamntshdt_6.Text = string.Empty;
            txt_principleamntshdt_6.Text = string.Empty;
            txt_payafterjudgamtshdt_6.Text = string.Empty;
            txt_deptamntshdt_6.Text = string.Empty;
            txt_lastpaydateshdt_6.Text = string.Empty;
            chk_calculate_add.Checked = false;
            txt_queueno_shdt.Text = string.Empty;
            linkfileopen.Text = string.Empty;
            linkfileopen.Visible = false;

            datatabledetailshow.Clear();
            if(isclearsearch) txt_search_custid_shdt.Text = string.Empty;
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
                        if (!(string.IsNullOrEmpty(customerid)|| string.IsNullOrEmpty(caseid)))
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
                                        if (m == 0) customDataList = sqlitedsrv.doGetDataCustomWithID(CPSCall.CustomerID ?? "", CPSCall.CaseID??"");
                                        if (customDataList.Count > 0) calcsrv.CaluLateData6CardCustom(ref CPSCall, customDataList, m);
                                    }
                                    else
                                    {
                                        calcsrv.CaluLateData6Card(ref CPSCall, setdata);
                                    }
                                    cardDataList[m] = CPSCall;
                                }
                                if (chk_calculate_add.Checked) calcsrv.CaluLateData6CardAddValue(ref cardDataList);
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
                if (isCheckKZASStatus())
                {
                    doPrintC2TableMeargeFile();
                }
                
            }
            else
            {
                MessageBox.Show("หมายเลขคิวห้ามว่าง", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txt_queueno_shdt.Focus();
            }

            Cursor = Cursors.Default;
        }

        private bool isCheckKZASStatus()
        {
                if (is_kzas)
                {
                    if (!chk_calculate_add.Checked)
                    {
                        DialogResult result = MessageBox.Show("สถานะทางคดี KZAS \r\n ต้องการ คำนวนยอดเพิ่มหรือไม่ ?", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (result == DialogResult.Yes)
                        {
                            chk_calculate_add.Checked = true;
                            return true;
                        }
                        else
                        {
                            return false;
                        }

                    }
                }
                else
                {
                    if (chk_calculate_add.Checked)
                    {
                        DialogResult result = MessageBox.Show("สถานะทางคดี ไม่ใช่ KZAS \r\n ต้องการ ยกเลิก การคำนวนยอดเพิ่ม หรือไม่ ?", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (result == DialogResult.Yes)
                        {
                            chk_calculate_add.Checked = false;
                            return false;
                        }
                        else
                        {
                           return true;
                        }
                    }
                }
                return true;
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
        private void doSetDispalyCustom(string? iscustomstr)
        {
            bool iscustom = false;
            if (iscustomstr == "Y") iscustom = true;
            lbl_customstatus_shdt.Visible = iscustom;
            lbl_customstatus_shdt.Text = "** ยอดตามทีมบังคับคดี ไม่คำนวนยอดปิด **";
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

        
    }
}