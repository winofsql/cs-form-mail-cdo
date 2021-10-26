
namespace cs_form_mail_cdo
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.server = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.portNo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.toAddress = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.subject = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.fromAddress = new System.Windows.Forms.TextBox();
            this.account = new System.Windows.Forms.TextBox();
            this.password = new System.Windows.Forms.TextBox();
            this.textBody = new System.Windows.Forms.TextBox();
            this.sendButton = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(48, 37);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 16);
            this.label1.TabIndex = 9;
            this.label1.Text = "サーバー";
            // 
            // server
            // 
            this.server.Location = new System.Drawing.Point(142, 33);
            this.server.Margin = new System.Windows.Forms.Padding(4);
            this.server.Name = "server";
            this.server.Size = new System.Drawing.Size(324, 23);
            this.server.TabIndex = 0;
            this.server.Text = "smtp.lolipop.jp";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(48, 73);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 16);
            this.label2.TabIndex = 10;
            this.label2.Text = "ポート";
            // 
            // portNo
            // 
            this.portNo.Location = new System.Drawing.Point(142, 73);
            this.portNo.Margin = new System.Windows.Forms.Padding(4);
            this.portNo.Name = "portNo";
            this.portNo.Size = new System.Drawing.Size(324, 23);
            this.portNo.TabIndex = 1;
            this.portNo.Text = "465";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(48, 114);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 16);
            this.label3.TabIndex = 11;
            this.label3.Text = "宛先";
            // 
            // toAddress
            // 
            this.toAddress.Location = new System.Drawing.Point(142, 111);
            this.toAddress.Margin = new System.Windows.Forms.Padding(4);
            this.toAddress.Name = "toAddress";
            this.toAddress.Size = new System.Drawing.Size(324, 23);
            this.toAddress.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(48, 155);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 16);
            this.label4.TabIndex = 12;
            this.label4.Text = "件名";
            // 
            // subject
            // 
            this.subject.Location = new System.Drawing.Point(142, 152);
            this.subject.Margin = new System.Windows.Forms.Padding(4);
            this.subject.Name = "subject";
            this.subject.Size = new System.Drawing.Size(675, 23);
            this.subject.TabIndex = 3;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(48, 197);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 16);
            this.label5.TabIndex = 13;
            this.label5.Text = "差出人";
            // 
            // fromAddress
            // 
            this.fromAddress.Location = new System.Drawing.Point(142, 194);
            this.fromAddress.Margin = new System.Windows.Forms.Padding(4);
            this.fromAddress.Name = "fromAddress";
            this.fromAddress.Size = new System.Drawing.Size(324, 23);
            this.fromAddress.TabIndex = 4;
            // 
            // account
            // 
            this.account.Location = new System.Drawing.Point(587, 34);
            this.account.Margin = new System.Windows.Forms.Padding(4);
            this.account.Name = "account";
            this.account.Size = new System.Drawing.Size(280, 23);
            this.account.TabIndex = 5;
            // 
            // password
            // 
            this.password.Location = new System.Drawing.Point(587, 73);
            this.password.Margin = new System.Windows.Forms.Padding(4);
            this.password.Name = "password";
            this.password.PasswordChar = '*';
            this.password.Size = new System.Drawing.Size(280, 23);
            this.password.TabIndex = 6;
            // 
            // textBody
            // 
            this.textBody.Location = new System.Drawing.Point(142, 239);
            this.textBody.Margin = new System.Windows.Forms.Padding(4);
            this.textBody.Multiline = true;
            this.textBody.Name = "textBody";
            this.textBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBody.Size = new System.Drawing.Size(725, 222);
            this.textBody.TabIndex = 7;
            // 
            // sendButton
            // 
            this.sendButton.Location = new System.Drawing.Point(143, 482);
            this.sendButton.Name = "sendButton";
            this.sendButton.Size = new System.Drawing.Size(133, 30);
            this.sendButton.TabIndex = 8;
            this.sendButton.Text = "送信";
            this.sendButton.UseVisualStyleBackColor = true;
            this.sendButton.Click += new System.EventHandler(this.sendButton_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(498, 37);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(66, 16);
            this.label6.TabIndex = 14;
            this.label6.Text = "アカウント";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(498, 76);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 16);
            this.label7.TabIndex = 15;
            this.label7.Text = "パスワード";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(895, 541);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.sendButton);
            this.Controls.Add(this.textBody);
            this.Controls.Add(this.password);
            this.Controls.Add(this.account);
            this.Controls.Add(this.fromAddress);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.subject);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.toAddress);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.portNo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.server);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox server;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox portNo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox toAddress;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox subject;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox fromAddress;
        private System.Windows.Forms.TextBox account;
        private System.Windows.Forms.TextBox password;
        private System.Windows.Forms.TextBox textBody;
        private System.Windows.Forms.Button sendButton;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
    }
}

