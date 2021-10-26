using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace cs_form_mail_cdo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void sendButton_Click(object sender, EventArgs e)
        {
            // CDO
            dynamic cdo = Activator.CreateInstance(Type.GetTypeFromProgID("CDO.Message"));

            var user = account.Text;
            var from = fromAddress.Text;
            var pass = password.Text;

            // ***************************
            // 自分のアドレスと宛先
            // ***************************
            cdo.From = from;
            cdo.To = toAddress.Text;

            // ***************************
            // 件名と本文
            // ***************************
            cdo.Subject = subject.Text;
            cdo.Textbody = textBody.Text;

            // ***************************
            // 設定
            // ***************************
            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/sendusing"] = 2;
            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpserver"] = server.Text;
            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpserverport"] = Int32.Parse(portNo.Text);

            // ポートが 465 の場合は true
            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpusessl"] = true;

            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"] = 1;
            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/sendusername"] = user;
            cdo.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/sendpassword"] = pass;

            // ***************************
            // 設定
            // ***************************
            cdo.Configuration.Fields.Update();

            // ***************************
            // 送信
            // ***************************
            try
            {
                cdo.Send();
                MessageBox.Show("メールを送信しました");
            }
            catch (Exception ex)
            {
                MessageBox.Show("cdo.Send() でエラーが発生しました");
            }

            // 解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(cdo);
        }
    }
}
