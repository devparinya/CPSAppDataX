using CPSAppData.Services;
using CPSAppData.UI.BaseForm;
using QueueAppManager.Service;
using System.Text.RegularExpressions;

namespace CPSAppData.UI.Setting
{
    public partial class frmAdminData : frmBaseWindow
    {
        securityService securityServ = new securityService();
        public frmAdminData()
        {
            InitializeComponent();
        }

        private void btn_decrypt_Click(object sender, EventArgs e)
        {
            txt_show.Text = string.Empty;
            string textdata =  txt_data.Text.ToString();
            if (IsValidBase64(textdata))
            {
                txt_show.Text = securityServ.DecryptString(textdata);
            }
        }

        private void btn_encrypt_Click(object sender, EventArgs e)
        {
            txt_show.Text = string.Empty;
            string textdata = txt_data.Text.ToString();
            txt_show.Text = securityServ.EncryptString(textdata);
        }

        private void btn_clear_Click(object sender, EventArgs e)
        {
            txt_data.Clear();
            txt_show.Clear();
        }
        public static bool IsValidBase64(string input) 
        { 
            if (string.IsNullOrEmpty(input) || input.Length % 4 != 0) 
            { 
                return false; 
            } 
            string pattern = @"^[a-zA-Z0-9\+/]*={0,2}$"; 
            Regex regex = new Regex(pattern); 
            return regex.IsMatch(input); 
        }
    }
}
