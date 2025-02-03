using CPSAppData.Services;
using CPSAppData.UI.BaseForm;
using CPSAppData.UI.Report;
using CPSAppData.UI.Setting;
using System.Data;

namespace CPSAppData
{
    public partial class frmMainApp : frmBaseWindow
    {
        excelDataService excelserv = new excelDataService();
        public frmMainApp()
        {
            InitializeComponent();
        }

        private void btn_setting_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Control)
            {
                this.Hide();
                frmAdminData frmadmin = new frmAdminData();
                frmadmin.ShowDialog();
                this.Show();
            }
            else
            {
                //this.Hide();
                //frmSettingData frmsetting = new frmSettingData();
                //frmsetting.ShowDialog();
                //this.Show();

                this.Hide();
                frmImportMasterData frmsetting = new frmImportMasterData();
                frmsetting.ShowDialog();
                this.Show();
            }
        }

        private void btn_report_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmReportManage frmreportManage = new frmReportManage();
            frmreportManage.ShowDialog();
            this.Show();
        }

        private void btn_user_report_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmUserReport frmuserreoirt = new frmUserReport();
            frmuserreoirt.ShowDialog();
            this.Show();
        }
    }
}
