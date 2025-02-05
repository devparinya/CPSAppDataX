using CPSAppData.Models;
using CPSAppData.Services;
using CPSAppData.UI.BaseForm;
using CpsDataApp.Models;
using CpsDataApp.Services;
using System.Data;

namespace CPSAppData.UI.Setting
{
    public partial class frmImportMasterData : frmBaseWindow
    {
        #region Member
        dataTableService dtService = new dataTableService();
        excelDataService excelService = new excelDataService();
        sqliteDataService sqlitesrv = new sqliteDataService();
        ComboBox[] CPSMasterCtrl = [];
        ComboBox[] CPSFestCtrl = [];
        ComboBox[] CPSCustomCtrl = [];
        DataTable datatablefest = new DataTable();
        DataTable datatablecustom = new DataTable();
        #endregion
        #region   Constructor
        public frmImportMasterData()
        {
            InitializeComponent();
            sqlitesrv.doinitTableDataCPSMaster();
            sqlitesrv.doinitTableMapData();
            sqlitesrv.doinitTableSettingData();
            sqlitesrv.initTableCPSPayData();
            sqlitesrv.doinitTableCPSFestData();
            doInitCtrlDataMasterCPS();
            doInitCtrlCPSFest();
            doInitCtrlCPSCustom();
            sqlitesrv.doLoadMappingDate(CPSMasterCtrl, "DATAMASTERMAP");
            sqlitesrv.doLoadMappingDate(CPSFestCtrl, "FESTMAP");
            sqlitesrv.doLoadMappingDate(CPSCustomCtrl, "FESTCUSTOMMAP");
        }
        #endregion
        #region Initial control
        private void doSettingGridData()
        {
            datagridfest.Columns[CPSFestCtrl[2].Text.Replace("cmb_fest_", "")].HeaderText = "ลำดับที่";
            datagridfest.Columns[CPSFestCtrl[3].Text.Replace("cmb_fest_", "")].HeaderText = "ลำดับกรม";
            datagridfest.Columns[CPSFestCtrl[0].Text.Replace("cmb_fest_", "")].HeaderText = "เลขที่บัตรประชาชน";
            datagridfest.Columns[CPSFestCtrl[1].Text.Replace("cmb_fest_", "")].HeaderText = "ชื่อ-นามสกุล";
            datagridfest.Columns[CPSFestCtrl[4].Text.Replace("cmb_fest_", "")].HeaderText = "หมาเหตุบังคับคดี";

            datagridfest.Columns[CPSFestCtrl[2].Text.Replace("cmb_fest_", "")].Width = 120;
            datagridfest.Columns[CPSFestCtrl[3].Text.Replace("cmb_fest_", "")].Width = 120;
            datagridfest.Columns[CPSFestCtrl[0].Text.Replace("cmb_fest_", "")].Width = 180;
            datagridfest.Columns[CPSFestCtrl[1].Text.Replace("cmb_fest_", "")].Width = 200;
            datagridfest.Columns[CPSFestCtrl[4].Text.Replace("cmb_fest_", "")].Width = 250;
            datagridfest.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }
        private void doInitCtrlDataMasterCPS()
        {
            CPSMasterCtrl = new ComboBox[]
            {
                cmb_CardNo,
                cmb_JudgmentAmnt,
                cmb_PrincipleAmnt,
                cmb_PayAfterJudgAmt,
                cmb_DeptAmnt,
                cmb_LastPayDate,
                cmb_CustomerName,
                cmb_CustomerID,
                cmb_CustomerTel,
                cmb_LegalStatus,
                cmb_BlackNo,
                cmb_RedNo,
                cmb_JudgeDate,
                cmb_CourtName,
                cmb_CollectorName,
                cmb_CollectorTeam,
                cmb_CollectorTel,
                cmb_LegalExecDate,
                cmb_LegalExecRemark,
                cmb_CaseID,
                cmb_CardStatus
            };
        }
        private void doInitCtrlCPSFest()
        {
            CPSFestCtrl = new ComboBox[]
            {
             cmb_fest_CustomerID,
             cmb_fest_CustomerName,
             cmb_fest_WorkNo,
             cmb_fest_LedNumber,
             cmb_fest_LegalExecRemark
            };
        }
        private void doInitCtrlCPSCustom()
        {
            CPSCustomCtrl = new ComboBox[]
            {
                cmb_custom_CustomerID,
                cmb_custom_CustomerName,
                cmb_custom_LedNumber,
                cmb_custom_WorkNo,
                cmb_custom_CardNo1,
                cmb_custom_CardNo2,
                cmb_custom_CardNo3,
                cmb_custom_CardNo4,
                cmb_custom_CardNo5,
                cmb_custom_CardNo6,
                cmb_custom_AccCloseAmnt1,
                cmb_custom_AccCloseAmnt2,
                cmb_custom_AccCloseAmnt3,
                cmb_custom_AccCloseAmnt4,
                cmb_custom_AccCloseAmnt5,
                cmb_custom_AccCloseAmnt6,
                cmb_custom_AccClose6Amnt1,
                cmb_custom_AccClose6Amnt2,
                cmb_custom_AccClose6Amnt3,
                cmb_custom_AccClose6Amnt4,
                cmb_custom_AccClose6Amnt5,
                cmb_custom_AccClose6Amnt6,
                cmb_custom_AccClose12Amnt1,
                cmb_custom_AccClose12Amnt2,
                cmb_custom_AccClose12Amnt3,
                cmb_custom_AccClose12Amnt4,
                cmb_custom_AccClose12Amnt5,
                cmb_custom_AccClose12Amnt6,
                cmb_custom_AccClose24Amnt1,
                cmb_custom_AccClose24Amnt2,
                cmb_custom_AccClose24Amnt3,
                cmb_custom_AccClose24Amnt4,
                cmb_custom_AccClose24Amnt5,
                cmb_custom_AccClose24Amnt6,
                cmb_custom_LegalExecRemark,
                cmb_custom_Installment6Amnt1,
                cmb_custom_Installment6Amnt2,
                cmb_custom_Installment6Amnt3,
                cmb_custom_Installment6Amnt4,
                cmb_custom_Installment6Amnt5,
                cmb_custom_Installment6Amnt6,
                cmb_custom_Installment12Amnt1,
                cmb_custom_Installment12Amnt2,
                cmb_custom_Installment12Amnt3,
                cmb_custom_Installment12Amnt4,
                cmb_custom_Installment12Amnt5,
                cmb_custom_Installment12Amnt6,
                cmb_custom_Installment24Amnt1,
                cmb_custom_Installment24Amnt2,
                cmb_custom_Installment24Amnt3,
                cmb_custom_Installment24Amnt4,
                cmb_custom_Installment24Amnt5,
                cmb_custom_Installment24Amnt6,
                cmb_custom_CaseID
            };
        }
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
        #endregion
        #region Mapping Data
        private void doAddHeaderDataFileMap(Dictionary<string, string>? dicHeader)
        {
            // Dictionary<string, string>? dicHeader = excelService.doGetColumnHDFromExcel(@"D:\template_master.xlsx");
            if (dicHeader != null)
            {

                cmb_CardNo.Items.Clear();
                cmb_CardNo.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_JudgmentAmnt.Items.Clear(); //ยอดพิพากษา
                cmb_JudgmentAmnt.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_PrincipleAmnt.Items.Clear(); //ต้นเงิน
                cmb_PrincipleAmnt.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_PayAfterJudgAmt.Items.Clear();//ชำระหลังพิพากษา
                cmb_PayAfterJudgAmt.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_DeptAmnt.Items.Clear();//ภาระหนี้ปัจจุบัน
                cmb_DeptAmnt.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_LastPayDate.Items.Clear();//ชำระครั้งล่าสุด
                cmb_LastPayDate.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CustomerName.Items.Clear(); //ชื่อ สกุล
                cmb_CustomerName.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CustomerID.Items.Clear();//รหัสบัตรประชาชน
                cmb_CustomerID.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CustomerTel.Items.Clear();//เบอร์โทรลูกค้า
                cmb_CustomerTel.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_LegalStatus.Items.Clear();//สถานะคดี
                cmb_LegalStatus.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_BlackNo.Items.Clear();//เลขคดีดำ
                cmb_BlackNo.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_RedNo.Items.Clear();//เลขคดีแดง
                cmb_RedNo.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_JudgeDate.Items.Clear();//วันพิพากษา        
                cmb_JudgeDate.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CourtName.Items.Clear();
                cmb_CourtName.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CollectorName.Items.Clear(); // ชื่อ collcctor
                cmb_CollectorName.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CollectorTeam.Items.Clear();//ทีม
                cmb_CollectorTeam.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CollectorTel.Items.Clear(); //โทร
                cmb_CollectorTel.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_LegalExecDate.Items.Clear();//วันที่ยึดฝอายัด
                cmb_LegalExecDate.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_LegalExecRemark.Items.Clear(); //หมายเหตุบังคับคดี
                cmb_LegalExecRemark.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CaseID.Items.Clear();
                cmb_CaseID.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_CardStatus.Items.Clear();
                cmb_CardStatus.Items.AddRange(dicHeader.Keys.ToArray());
            }
        }
        private void doAddHeaderDataFestMap(Dictionary<string, string>? dicHeader)
        {
            if (dicHeader != null)
            {
                cmb_fest_CustomerID.Items.Clear();
                cmb_fest_CustomerID.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_fest_CustomerName.Items.Clear();
                cmb_fest_CustomerName.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_fest_WorkNo.Items.Clear();
                cmb_fest_WorkNo.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_fest_LedNumber.Items.Clear();
                cmb_fest_LedNumber.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_fest_LegalExecRemark.Items.Clear();
                cmb_fest_LegalExecRemark.Items.AddRange(dicHeader.Keys.ToArray());
            }
        }
        private void doAddHeaderDataCustomMap(Dictionary<string, string>? dicHeader)
        {
            if (dicHeader != null)
            {
                Cursor = Cursors.WaitCursor;
                cmb_custom_CustomerID.Items.Clear();
                cmb_custom_CustomerID.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CustomerName.Items.Clear();
                cmb_custom_CustomerName.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_LedNumber.Items.Clear();
                cmb_custom_LedNumber.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_WorkNo.Items.Clear();
                cmb_custom_WorkNo.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CardNo1.Items.Clear();
                cmb_custom_CardNo1.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CardNo2.Items.Clear();
                cmb_custom_CardNo2.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CardNo3.Items.Clear();
                cmb_custom_CardNo3.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CardNo4.Items.Clear();
                cmb_custom_CardNo4.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CardNo5.Items.Clear();
                cmb_custom_CardNo5.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_CardNo6.Items.Clear();
                cmb_custom_CardNo6.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccCloseAmnt1.Items.Clear();
                cmb_custom_AccCloseAmnt1.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccCloseAmnt2.Items.Clear();
                cmb_custom_AccCloseAmnt2.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccCloseAmnt3.Items.Clear();
                cmb_custom_AccCloseAmnt3.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccCloseAmnt4.Items.Clear();
                cmb_custom_AccCloseAmnt4.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccCloseAmnt5.Items.Clear();
                cmb_custom_AccCloseAmnt5.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccCloseAmnt6.Items.Clear();
                cmb_custom_AccCloseAmnt6.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose6Amnt1.Items.Clear();
                cmb_custom_AccClose6Amnt1.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose6Amnt2.Items.Clear();
                cmb_custom_AccClose6Amnt2.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose6Amnt3.Items.Clear();
                cmb_custom_AccClose6Amnt3.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose6Amnt4.Items.Clear();
                cmb_custom_AccClose6Amnt4.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose6Amnt5.Items.Clear();
                cmb_custom_AccClose6Amnt5.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose6Amnt6.Items.Clear();
                cmb_custom_AccClose6Amnt6.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose12Amnt1.Items.Clear();
                cmb_custom_AccClose12Amnt1.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose12Amnt2.Items.Clear();
                cmb_custom_AccClose12Amnt2.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose12Amnt3.Items.Clear();
                cmb_custom_AccClose12Amnt3.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose12Amnt4.Items.Clear();
                cmb_custom_AccClose12Amnt4.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose12Amnt5.Items.Clear();
                cmb_custom_AccClose12Amnt5.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose12Amnt6.Items.Clear();
                cmb_custom_AccClose12Amnt6.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose24Amnt1.Items.Clear();
                cmb_custom_AccClose24Amnt1.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose24Amnt2.Items.Clear();
                cmb_custom_AccClose24Amnt2.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose24Amnt3.Items.Clear();
                cmb_custom_AccClose24Amnt3.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose24Amnt4.Items.Clear();
                cmb_custom_AccClose24Amnt4.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose24Amnt5.Items.Clear();
                cmb_custom_AccClose24Amnt5.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_AccClose24Amnt6.Items.Clear();
                cmb_custom_AccClose24Amnt6.Items.AddRange(dicHeader.Keys.ToArray());

                cmb_custom_LegalExecRemark.Items.Clear();
                cmb_custom_LegalExecRemark.Items.AddRange(dicHeader.Keys.ToArray());

                Cursor = Cursors.Default;
            }
        }
        private void doSaveDataMapMaster()
        {
            Cursor = Cursors.WaitCursor;
            bool result = sqlitesrv.doSaveMapMaterCPS("DATAMASTERMAP", CPSMasterCtrl);
            if (result) MessageBox.Show("บันทึกสำเร็จ", "บันทึกสำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Cursor = Cursors.Default;
        }
        private void doSaveDataMapFest()
        {
            Cursor = Cursors.WaitCursor;
            bool result = sqlitesrv.doSaveMapMaterCPS("FESTMAP", CPSFestCtrl);
            if (result) MessageBox.Show("บันทึกสำเร็จ", "บันทึกสำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Cursor = Cursors.Default;

        }
        private void doSaveDataMapFestCustom()
        {
            Cursor = Cursors.WaitCursor;
            bool result = sqlitesrv.doSaveMapMaterCPS("FESTCUSTOMMAP", CPSCustomCtrl);
            if (result) MessageBox.Show("บันทึกสำเร็จ", "บันทึกสำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Cursor = Cursors.Default;

        }
        #endregion
        #region Show Data
        private void doLoadDataExcelFestData()
        {
            string pathfile = txt_pathName_fest.Text;
            datatablefest = dtService.doCreateFestDataTemplate();
            if (!string.IsNullOrEmpty(pathfile))
            {
                if (File.Exists(pathfile))
                {
                    Cursor = Cursors.WaitCursor;
                    datatablefest = excelService.excelToDataTable(pathfile);
                    datagridfest.DataSource = datatablefest;
                    doSettingGridData();
                    Cursor = Cursors.Default;
                }
                else
                {
                    MessageBox.Show(string.Format("{0} \r\n ไม่มีอยู่จริง", pathfile), "File Not Exist", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        #endregion
        #region Insert & Update
        private void doSaveDataExcelFestCustom()
        {
            string pathfile = txt_festcustom_path.Text;
            datatablecustom = dtService.doCreateFestCustomTemplate();
            if (!string.IsNullOrEmpty(pathfile))
            {
                if (File.Exists(pathfile))
                {
                    Cursor = Cursors.WaitCursor;
                    datatablecustom = excelService.excelToDataTable(pathfile);
                    sqlitesrv.doInsertCustomData(datatablecustom, CPSCustomCtrl, this.progressBarCustom, chk_clear_custom.Checked);
                    excelService.ExportExcelResult(sqlitesrv.getResultData(), Path.GetDirectoryName(pathfile) ?? "", "ExecuteCPS");
                    MessageBox.Show("นำเข้าข้อมูลเรียบร้อย", "นำเข้าข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Cursor = Cursors.Default;
                }
                else
                {
                    MessageBox.Show(string.Format("{0} \r\n ไม่มีอยู่จริง", pathfile), "File Not Exist", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private void doInsetCPSFestData()
        {
            if (datatablefest.Rows.Count > 0)
            {
                Cursor = Cursors.WaitCursor;
                sqlitesrv.doInsertFestData(datatablefest, CPSFestCtrl, this.progressBarFest, chk_clear_cpsdata.Checked);
                excelService.ExportExcelResult(sqlitesrv.getResultData(), Path.GetDirectoryName(txt_pathName_fest.Text) ?? "", "CSPData");
                Cursor = Cursors.Default;
                MessageBox.Show("นำเข้าข้อมูลเรียบร้อย", "นำเข้าข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion
        #region Other Method
        private void doExportDataDuplicateID()
        {
            DataTable dtDuplicate = dtService.doCreateMasterCPSDataDuplicate();
            dateTimeHelper dateHelper = new dateTimeHelper();
            List<DataCPSPerson> dataduplicate = sqlitesrv.doGetMasterDuplicateIDDataCPSAll();
            for (int i = 0; i < dataduplicate.Count; i++)
            {
                DataRow datadup = dtDuplicate.NewRow();
                datadup["CustomerID"] = dataduplicate[i].CustomerID;
                datadup["CustomerName"] = dataduplicate[i].CustomerName;
                datadup["CustomerTel"] = dataduplicate[i].CustomerTel;
                datadup["LegalStatus"] = dataduplicate[i].LegalStatus;
                datadup["BlackNo"] = dataduplicate[i].BlackNo;
                datadup["RedNo"] = dataduplicate[i].RedNo;
                datadup["JudgeDate"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].JudgeDate ?? string.Empty);
                datadup["CourtName"] = dataduplicate[i].CourtName;
                datadup["LegalExecRemark"] = dataduplicate[i].LegalExecRemark;
                datadup["LegalExecDate"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LegalExecDate ?? string.Empty);
                datadup["CollectorName"] = dataduplicate[i].CollectorName;
                datadup["CollectorTeam"] = dataduplicate[i].CollectorTeam;
                datadup["CollectorTel"] = dataduplicate[i].CollectorTel;
                datadup["CardNo1"] = dataduplicate[i].CardNo1;
                datadup["JudgmentAmnt1"] = dataduplicate[i].JudgmentAmnt1;
                datadup["PrincipleAmnt1"] = dataduplicate[i].PrincipleAmnt1;
                datadup["PayAfterJudgAmt1"] = dataduplicate[i].PayAfterJudgAmt1;
                datadup["DeptAmnt1"] = dataduplicate[i].DeptAmnt1;
                datadup["LastPayDate1"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate1 ?? string.Empty);
                datadup["CardNo1"] = dataduplicate[i].CardNo1;
                datadup["JudgmentAmnt1"] = dataduplicate[i].JudgmentAmnt1;
                datadup["PrincipleAmnt1"] = dataduplicate[i].PrincipleAmnt1;
                datadup["PayAfterJudgAmt1"] = dataduplicate[i].PayAfterJudgAmt1;
                datadup["DeptAmnt1"] = dataduplicate[i].DeptAmnt1;
                datadup["LastPayDate1"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate1 ?? string.Empty);
                datadup["CardNo1"] = dataduplicate[i].CardNo1;
                datadup["JudgmentAmnt1"] = dataduplicate[i].JudgmentAmnt1;
                datadup["PrincipleAmnt1"] = dataduplicate[i].PrincipleAmnt1;
                datadup["PayAfterJudgAmt1"] = dataduplicate[i].PayAfterJudgAmt1;
                datadup["DeptAmnt1"] = dataduplicate[i].DeptAmnt1;
                datadup["LastPayDate1"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate1 ?? string.Empty);
                datadup["CardNo1"] = dataduplicate[i].CardNo1;
                datadup["JudgmentAmnt1"] = dataduplicate[i].JudgmentAmnt1;
                datadup["PrincipleAmnt1"] = dataduplicate[i].PrincipleAmnt1;
                datadup["PayAfterJudgAmt1"] = dataduplicate[i].PayAfterJudgAmt1;
                datadup["DeptAmnt1"] = dataduplicate[i].DeptAmnt1;
                datadup["LastPayDate1"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate1 ?? string.Empty);
                datadup["CardNo2"] = dataduplicate[i].CardNo2;
                datadup["JudgmentAmnt2"] = dataduplicate[i].JudgmentAmnt2;
                datadup["PrincipleAmnt2"] = dataduplicate[i].PrincipleAmnt2;
                datadup["PayAfterJudgAmt2"] = dataduplicate[i].PayAfterJudgAmt2;
                datadup["DeptAmnt2"] = dataduplicate[i].DeptAmnt2;
                datadup["LastPayDate2"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate2 ?? string.Empty);
                datadup["CardNo3"] = dataduplicate[i].CardNo3;
                datadup["JudgmentAmnt3"] = dataduplicate[i].JudgmentAmnt3;
                datadup["PrincipleAmnt3"] = dataduplicate[i].PrincipleAmnt3;
                datadup["PayAfterJudgAmt3"] = dataduplicate[i].PayAfterJudgAmt3;
                datadup["DeptAmnt3"] = dataduplicate[i].DeptAmnt3;
                datadup["LastPayDate3"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate3 ?? string.Empty);
                datadup["CardNo4"] = dataduplicate[i].CardNo4;
                datadup["JudgmentAmnt4"] = dataduplicate[i].JudgmentAmnt4;
                datadup["PrincipleAmnt4"] = dataduplicate[i].PrincipleAmnt4;
                datadup["PayAfterJudgAmt4"] = dataduplicate[i].PayAfterJudgAmt4;
                datadup["DeptAmnt4"] = dataduplicate[i].DeptAmnt4;
                datadup["LastPayDate4"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate4 ?? string.Empty);
                datadup["CardNo5"] = dataduplicate[i].CardNo5;
                datadup["JudgmentAmnt5"] = dataduplicate[i].JudgmentAmnt5;
                datadup["PrincipleAmnt5"] = dataduplicate[i].PrincipleAmnt5;
                datadup["PayAfterJudgAmt5"] = dataduplicate[i].PayAfterJudgAmt5;
                datadup["DeptAmnt5"] = dataduplicate[i].DeptAmnt5;
                datadup["LastPayDate5"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate5 ?? string.Empty);
                datadup["CardNo6"] = dataduplicate[i].CardNo6;
                datadup["JudgmentAmnt6"] = dataduplicate[i].JudgmentAmnt6;
                datadup["PrincipleAmnt6"] = dataduplicate[i].PrincipleAmnt6;
                datadup["PayAfterJudgAmt6"] = dataduplicate[i].PayAfterJudgAmt6;
                datadup["DeptAmnt6"] = dataduplicate[i].DeptAmnt6;
                datadup["LastPayDate6"] = dateHelper.doGetShortDateTHFromDBToPDF(dataduplicate[i].LastPayDate6 ?? string.Empty);
                dtDuplicate.Rows.Add(datadup);
            }
            if (dtDuplicate.Rows.Count > 0)
            {
                string filexlspath = selectSavefile("DataMasterDuplicate.xlsx");
                if (string.IsNullOrEmpty(filexlspath)) return;
                excelDataService xlsserv = new excelDataService();
                xlsserv.ExportToExcelAll(dtDuplicate, filexlspath);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filexlspath), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void doExportDataCustomDuplicateID()
        {
            DataTable dtDuplicateCustom = dtService.doCreateFestCustomTemplate();
            List<FestCustom> datafestCustoms = sqlitesrv.doGetCustomDuplicateIDAll();
            for (int i = 0; i < datafestCustoms.Count; i++)
            {
                DataRow datadup = dtDuplicateCustom.NewRow();
                datadup["CustomerID"] = datafestCustoms[i].CustomerID;
                datadup["CustomerName"] = datafestCustoms[i].CustomerName;
                datadup["WorkNo"] = datafestCustoms[i].WorkNo;
                datadup["LedNumber"] = datafestCustoms[i].LedNumber;
                datadup["CardNo1"] = datafestCustoms[i].CardNo1;
                datadup["AccCloseAmnt1"] = datafestCustoms[i].AccCloseAmnt1;
                datadup["AccClose6Amnt1"] = datafestCustoms[i].AccClose6Amnt1;
                datadup["AccClose12Amnt1"] = datafestCustoms[i].AccClose12Amnt1;
                datadup["AccClose24Amnt1"] = datafestCustoms[i].AccClose24Amnt1;
                datadup["CardNo2"] = datafestCustoms[i].CardNo2;
                datadup["AccCloseAmnt2"] = datafestCustoms[i].AccCloseAmnt2;
                datadup["AccClose6Amnt2"] = datafestCustoms[i].AccClose6Amnt2;
                datadup["AccClose12Amnt2"] = datafestCustoms[i].AccClose12Amnt2;
                datadup["AccClose24Amnt2"] = datafestCustoms[i].AccClose24Amnt2;
                datadup["CardNo3"] = datafestCustoms[i].CardNo3;
                datadup["AccCloseAmnt3"] = datafestCustoms[i].AccCloseAmnt3;
                datadup["AccClose6Amnt3"] = datafestCustoms[i].AccClose6Amnt3;
                datadup["AccClose12Amnt3"] = datafestCustoms[i].AccClose12Amnt3;
                datadup["AccClose24Amnt3"] = datafestCustoms[i].AccClose24Amnt3;
                datadup["CardNo4"] = datafestCustoms[i].CardNo4;
                datadup["AccCloseAmnt4"] = datafestCustoms[i].AccCloseAmnt4;
                datadup["AccClose6Amnt4"] = datafestCustoms[i].AccClose6Amnt4;
                datadup["AccClose12Amnt4"] = datafestCustoms[i].AccClose12Amnt4;
                datadup["AccClose24Amnt4"] = datafestCustoms[i].AccClose24Amnt4;
                datadup["CardNo5"] = datafestCustoms[i].CardNo5;
                datadup["AccCloseAmnt5"] = datafestCustoms[i].AccCloseAmnt5;
                datadup["AccClose6Amnt5"] = datafestCustoms[i].AccClose6Amnt5;
                datadup["AccClose12Amnt5"] = datafestCustoms[i].AccClose12Amnt5;
                datadup["AccClose24Amnt5"] = datafestCustoms[i].AccClose24Amnt5;
                datadup["CardNo6"] = datafestCustoms[i].CardNo6;
                datadup["AccCloseAmnt6"] = datafestCustoms[i].AccCloseAmnt6;
                datadup["AccClose6Amnt6"] = datafestCustoms[i].AccClose6Amnt6;
                datadup["AccClose12Amnt6"] = datafestCustoms[i].AccClose12Amnt6;
                datadup["AccClose24Amnt6"] = datafestCustoms[i].AccClose24Amnt6;
                datadup["LegalExecRemark"] = datafestCustoms[i].LegalExecRemark;
                dtDuplicateCustom.Rows.Add(datadup);
            }
            if (dtDuplicateCustom.Rows.Count > 0)
            {
                string filexlspath = selectSavefile("DataExecuteDuplicate.xlsx");
                if (string.IsNullOrEmpty(filexlspath)) return;
                excelDataService xlsserv = new excelDataService();
                xlsserv.ExportToExcelAll(dtDuplicateCustom, filexlspath);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filexlspath), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void doExportDataFestNoMaster()
        {
            DataTable dtFestNoMaster = dtService.doCreateFestDataTemplate();
            List<DataCPSPerson> datanotmaster = sqlitesrv.doCheckCPSFestDataNotInMaster();
            for (int i = 0; i < datanotmaster.Count; i++)
            {
                DataRow datanotms = dtFestNoMaster.NewRow();
                datanotms["CustomerID"] = datanotmaster[i].CustomerID;
                datanotms["CustomerName"] = datanotmaster[i].CustomerName;
                datanotms["WorkNo"] = datanotmaster[i].WorkNo;
                datanotms["LedNumber"] = datanotmaster[i].LedNumber;
                dtFestNoMaster.Rows.Add(datanotms);
            }
            if (dtFestNoMaster.Rows.Count > 0)
            {
                string filexlspath = selectSavefile("DataCPSNotMaster.xlsx");
                if (string.IsNullOrEmpty(filexlspath)) return;
                excelDataService xlsserv = new excelDataService();
                xlsserv.ExportToExcelAll(dtFestNoMaster, filexlspath);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filexlspath), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("ไม่พบข้อมูลผิดปกติ", "ตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void doExportDataCustomNoMaster()
        {
            DataTable dtCustomNoMaster = dtService.doCreateFestDataTemplate();
            List<DataCPSPerson> datanotmaster = sqlitesrv.doCheckCPSCustomNotInMaster();
            for (int i = 0; i < datanotmaster.Count; i++)
            {
                DataRow datanotms = dtCustomNoMaster.NewRow();
                datanotms["CustomerID"] = datanotmaster[i].CustomerID;
                datanotms["CustomerName"] = datanotmaster[i].CustomerName;
                datanotms["WorkNo"] = datanotmaster[i].WorkNo;
                datanotms["LedNumber"] = datanotmaster[i].LedNumber;
                dtCustomNoMaster.Rows.Add(datanotms);
            }
            if (dtCustomNoMaster.Rows.Count > 0)
            {
                string filexlspath = selectSavefile("DataExecuteNotMaster.xlsx");
                if (string.IsNullOrEmpty(filexlspath)) return;
                excelDataService xlsserv = new excelDataService();
                xlsserv.ExportToExcelAll(dtCustomNoMaster, filexlspath);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filexlspath), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("ไม่พบข้อมูลผิดปกติ", "ตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void doExportCPSDataDuplicateID()
        {
            DataTable dtFestNoMaster = dtService.doCreateFestDataTemplate();
            List<DataCPSPerson> datanotmaster = sqlitesrv.doGetFestDuplicateIDAll();
            for (int i = 0; i < datanotmaster.Count; i++)
            {
                DataRow datanotms = dtFestNoMaster.NewRow();
                datanotms["CustomerID"] = datanotmaster[i].CustomerID;
                datanotms["CustomerName"] = datanotmaster[i].CustomerName;
                datanotms["WorkNo"] = datanotmaster[i].WorkNo;
                datanotms["LedNumber"] = datanotmaster[i].LedNumber;
                dtFestNoMaster.Rows.Add(datanotms);
            }
            if (dtFestNoMaster.Rows.Count > 0)
            {
                string filexlspath = selectSavefile("DataCPSdataDuplicate.xlsx");
                if (string.IsNullOrEmpty(filexlspath)) return;
                excelDataService xlsserv = new excelDataService();
                xlsserv.ExportToExcelAll(dtFestNoMaster, filexlspath);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filexlspath), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("ไม่พบข้อมูลผิดปกติ", "ตรวจสอบ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void doSelectedFileMaster(TextBox txtpath)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel Files|*.xls;*.xlsx;";
            if (openFile.ShowDialog(this) == DialogResult.OK)
            {
                txtpath.Text = openFile.FileName;
            }
            else
            {
                txtpath.Text = string.Empty;
            }
        }
        private Dictionary<string, string>? doGetHeaderFileExcel(TextBox txtpath)
        {
            Dictionary<string, string>? dicHeader = null;
            if (File.Exists(txtpath.Text))
            {
                Cursor.Current = Cursors.WaitCursor;
                dicHeader = excelService.doGetColumnHDFromExcel(txtpath.Text, true);
                Cursor.Current = Cursors.Default;
            }
            else
            {
                MessageBox.Show(string.Format("{0} \r\n ไม่มีอยู่จริง", txtpath.Text), "File Not Exist", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return dicHeader;
        }
        private Dictionary<string, string>? doSelectFileExcel(TextBox txtpath, bool ismaster)
        {
            Dictionary<string, string>? dicHeader = null;
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel Files|*.xls;*.xlsx;";
            if (openFile.ShowDialog(this) == DialogResult.OK)
            {
                txtpath.Text = openFile.FileName;
                Cursor.Current = Cursors.WaitCursor;
                dicHeader = excelService.doGetColumnHDFromExcel(txtpath.Text, ismaster);
                Cursor.Current = Cursors.Default;
            }
            else
            {
                txtpath.Text = string.Empty;
            }
            return dicHeader;
        }
        private string selectSavefile(string filename_)
        {
            string path = string.Empty;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Save a Excel File";
            saveFileDialog.FileName = filename_;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = saveFileDialog.FileName;
            }
            return path;
        }
        #endregion
        #region Event Control
        private void btn_brows_file_Click(object sender, EventArgs e)
        {
            doSelectedFileMaster(txt_path_excel);


        }
        private void btn_header_get_Click(object sender, EventArgs e)
        {
            Dictionary<string, string>? dicHD = doGetHeaderFileExcel(txt_path_excel);
            if (dicHD != null) doAddHeaderDataFileMap(dicHD);
        }

        private void btn_loaddata_excel_Click(object sender, EventArgs e)
        {
            string pathfile = txt_path_excel.Text;
            if (!string.IsNullOrEmpty(pathfile))
            {
                if (File.Exists(pathfile))
                {
                    Cursor = Cursors.WaitCursor;
                    DataTable dtraw = excelService.excelToDataTable(pathfile);
                    doSortDataTable(ref dtraw, "CustomerID ASC, CaseID ASC");
                    sqlitesrv.doInsertDataCPSMaster(dtraw, CPSMasterCtrl, chk_clear_master.Checked, this.progressBarMaster);
                    excelService.ExportExcelResult(sqlitesrv.getResultData(), Path.GetDirectoryName(pathfile) ?? "", "Master");
                    Cursor = Cursors.Default;

                    MessageBox.Show("นำเข้าข้อมูลเรียบร้อย", "นำเข้าข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(string.Format("{0} \r\n ไม่มีอยู่จริง", pathfile), "File Not Exist", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private void doSortDataTable(ref DataTable datasoure, string sort_str)
        {
            DataView view = datasoure.DefaultView;
            view.Sort = sort_str;

            datasoure = view.ToTable();
        }
        private void btn_fest_save_Click(object sender, EventArgs e)
        {
            doSaveDataMapFest();
        }
        private void btn_fest_custom_save_Click(object sender, EventArgs e)
        {
            doSaveDataMapFestCustom();
        }
        private void btn_festpath_Click(object sender, EventArgs e)
        {
            Dictionary<string, string>? dicHD = doSelectFileExcel(txt_pathName_fest, false);
            if (dicHD != null) doAddHeaderDataFestMap(dicHD);
        }
        private void btn_festcustom_path_Click(object sender, EventArgs e)
        {
            Dictionary<string, string>? dicHD = doSelectFileExcel(txt_festcustom_path, false);
            if (dicHD != null) doAddHeaderDataCustomMap(dicHD);
        }
        private void btn_fest_process_Click(object sender, EventArgs e)
        {
            doLoadDataExcelFestData();//doUpdateCPSMasterWithFestData();
        }
        private void btn_fest_custom_process_Click(object sender, EventArgs e)
        {
            doSaveDataExcelFestCustom();
        }
        private void btn_savefaestto_db_Click(object sender, EventArgs e)
        {
            doInsetCPSFestData();
        }
        private void btn_check_datadup_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doExportDataDuplicateID();
            Cursor = Cursors.Default;
        }
        private void btn_export_template_Click(object sender, EventArgs e)
        {
            string folderpath = browFolder();
            string filename = "";
            string[] captionth = new string[] { "CaseID", "สถานะบัตร", "หมายเลขบัตร", "ยอดพิพากษา", "ต้นเงิน", "ชำระหลังพิพากษา", "ภาระหนี้ปัจจุบัน", "LastPayDate", "ชื่อ-สกุล (ใหม่)", "เบอร์ติดต่อ(ลูกค้า)", "ID", "Legal Status", "คดีดำ", "คดีแดง", "วันพิพากษา", "ชื่อศาล", "หมายเหตุบังคับคดี", "วันที่ยึดอายัติ(ล่าสุด)", "Collector", "Team", "เบอร์ติดต่อ" };
            string[] captioneng = new string[] { "CaseID", "CardStatus", "CardNo", "JudgmentAmnt", "PrincipleAmnt", "PayAfterJudgAmt", "DeptAmnt", "LastPayDate", "CustomerName", "CustomerTel", "CustomerID", "LegalStatus", "BlackNo", "RedNo", "JudgeDate", "CourtName", "LegalExecRemark", "LegalExecDate", "CollectorName", "CollectorTeam", "CollectorTel", };

            if (!string.IsNullOrEmpty(folderpath))
            {
                DataTable datatatemplate = dtService.doCreateDataCPSMasterTemplate();
                filename = Path.Combine(folderpath, "template_master.xlsx");
                excelService.ExportToExcel(datatatemplate, captionth, captioneng, filename);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filename), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void btn_fest_template_Click(object sender, EventArgs e)
        {
            string folderpath = browFolder();
            string filename = "";
            string[] captionth = new string[] { "เลขที่ชุดงาน", "ลำดับกรม", "ID", "ชื่อ-สกุล", "หมายเหตุบังคับคดี" };
            string[] captioneng = new string[] { "WorkNo", "LedNumber", "CustomerID", "CustomerName", "LegalExecRemark" };
            if (!string.IsNullOrEmpty(folderpath))
            {
                DataTable datatatemplate = dtService.doCreateFestDataTemplate();
                filename = Path.Combine(folderpath, "template_cpsdata.xlsx");
                excelService.ExportToExcel(datatatemplate, captionth, captioneng, filename);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filename), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void btn_fest_csutomtemplate_Click(object sender, EventArgs e)
        {
            string folderpath = browFolder();
            string filename = "";
            string[] captionth = new string[] {"CaseID","ID","ชื่อ-สกุล","เลขที่ชุดงาน","ลำดับกรม","หมายเลขบัตร1","ปิดงวดเดียวบัตร1","ผ่อน 6 เดือนบัตร1","งวดละ(6เดือน บัตร1)","ผ่อน 12 เดือนบัตร1","งวดละ(12เดือน บัตร1)","ผ่อน 24 เดือนบัตร1","งวดละ(24เดือน บัตร1)","หมายเลขบัตร2","ปิดงวดเดียวบัตร2","ผ่อน 6 เดือนบัตร2","งวดละ(6เดือน บัตร2)","ผ่อน 12 เดือนบัตร2","งวดละงวดละ(12เดือน บัตร2)","ผ่อน 24 เดือนบัตร2","งวดละงวดละ(24เดือน บัตร2)"
                                               ,"หมายเลขบัตร3","ปิดงวดเดียวบัตร3","ผ่อน 6 เดือนบัตร3","งวดละ(6เดือน บัตร3)","ผ่อน 12 เดือนบัตร3","งวดละ(12เดือน บัตร3)","ผ่อน 24 เดือนบัตร3","งวดละ(24เดือน บัตร3)","หมายเลขบัตร4","ปิดงวดเดียวบัตร4","ผ่อน 6 เดือนบัตร4","งวดละ(6เดือน บัตร4)","ผ่อน 12 เดือนบัตร4","งวดละ(12เดือน บัตร4)","ผ่อน 24 เดือนบัตร4","งวดละ(24เดือน บัตร4)","หมายเลขบัตร5","ปิดงวดเดียวบัตร5"
                                               ,"ผ่อน 6 เดือนบัตร5","งวดละ(6เดือน บัตร5)","ผ่อน 12 เดือนบัตร5","งวดละ(12เดือน บัตร5)","ผ่อน 24 เดือนบัตร5","งวดละ(24เดือน บัตร5)","หมายเลขบัตร6","ปิดงวดเดียวบัตร6","ผ่อน 6 เดือนบัตร6","งวดละ(6เดือน บัตร6)","ผ่อน 12 เดือนบัตร6","งวดละ(12เดือน บัตร6)","ผ่อน 24 เดือนบัตร6","งวดละ(24เดือน บัตร6)","หมายเหตุบังคับคดี" };

            string[] captioneng = new string[] {"CaseID","CustomerID","CustomerName","WorkNo","LedNumber","CardNo1","AccCloseAmnt1","AccClose6Amnt1","Installment6Amnt1","AccClose12Amnt1","Installment12Amnt1","AccClose24Amnt1","Installment24Amnt1","CardNo2","AccCloseAmnt2","AccClose6Amnt2","Installment6Amnt2","AccClose12Amnt2","Installment12Amnt2"
                                                ,"AccClose24Amnt2","Installment24Amnt2","CardNo3","AccCloseAmnt3","AccClose6Amnt3","Installment6Amnt3","AccClose12Amnt3","Installment12Amnt3","AccClose24Amnt3","Installment24Amnt3","CardNo4","AccCloseAmnt4","AccClose6Amnt4","Installment6Amnt4","AccClose12Amnt4","Installment12Amnt4","AccClose24Amnt4","Installment24Amnt4"
                                                ,"CardNo5","AccCloseAmnt5","AccClose6Amnt5","Installment6Amnt5","AccClose12Amnt5","Installment12Amnt5","AccClose24Amnt5","Installment24Amnt5","CardNo6","AccCloseAmnt6","AccClose6Amnt6","Installment6Amnt6","AccClose12Amnt6","Installment12Amnt6","AccClose24Amnt6","Installment24Amnt6", "LegalExecRemark"};

            if (!string.IsNullOrEmpty(folderpath))
            {
                DataTable datatatemplate = dtService.doCreateFestCustomTemplate();
                filename = Path.Combine(folderpath, "template_execute_data.xlsx");
                excelService.ExportToExcel(datatatemplate, captionth, captioneng, filename);
                MessageBox.Show(string.Format("บันทึกเรียบร้อย \r\n {0}", filename), "บันทึก", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void btn_save_map_Click(object sender, EventArgs e)
        {
            doSaveDataMapMaster();
        }
        #endregion
        #region Default Map
        private void doLoadDefaultDataMasterCPS()
        {
            for (int i = 0; i < CPSMasterCtrl.Length; i++)
            {
                string ctrlname = CPSMasterCtrl[i].Name;
                CPSMasterCtrl[i].Text = ctrlname.Replace("cmb_", "");
            }
        }
        private void doLoadDefaultCPSCustom()
        {
            for (int i = 0; i < CPSCustomCtrl.Length; i++)
            {
                string ctrlname = CPSCustomCtrl[i].Name;
                CPSCustomCtrl[i].Text = ctrlname.Replace("cmb_custom_", "");
            }
        }

        #endregion

        private void btn_load_default_Click(object sender, EventArgs e)
        {
            doLoadDefaultDataMasterCPS();
        }


        private void btn_custom_loaddefault_Click(object sender, EventArgs e)
        {
            doLoadDefaultCPSCustom();
        }
        private void btn_checkdata_custom_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doExportDataCustomDuplicateID();
            Cursor = Cursors.Default;
        }
        private void btn_check_festnomater_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doExportDataFestNoMaster();
            Cursor = Cursors.Default;
        }

        private void btn_check_customnomater_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doExportDataCustomNoMaster();
            Cursor = Cursors.Default;
        }

        private void btn_docheck_cpsdatadup_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            doExportCPSDataDuplicateID();
            Cursor = Cursors.Default;
        }

        private void btn_fest_custom_save_Click_1(object sender, EventArgs e)
        {
            doSaveDataMapFestCustom();
        }
    }
}
