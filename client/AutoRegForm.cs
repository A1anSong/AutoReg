using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using MailKit;
using MailKit.Net.Imap;

namespace AutoReg;

public partial class AutoRegForm : Form
{
    public AutoRegForm()
    {
        InitializeComponent();
    }

    private const int WM_COPYDATA = 0x004A;

    public struct COPYDATASTRUCT
    {
        public IntPtr dwData;
        public int cData;

        [MarshalAs(UnmanagedType.LPStr)] public string lpData;
    }

    protected override void DefWndProc(ref Message m)
    {
        switch (m.Msg)
        {
            case WM_COPYDATA:
                var cds = new COPYDATASTRUCT();
                var t = cds.GetType();
                cds = (COPYDATASTRUCT)m.GetLParam(t);
                textBox3.AppendText(cds.lpData);
                break;
            default:
                base.DefWndProc(ref m);
                break;
        }
    }

    private void Apply911Proxy()
    {
        textBox3.Clear();
        var rd = new Random();
        var port = rd.Next(4000, 4100);
        textBox3.AppendText($"Port: {port}");
        textBox3.AppendText(Environment.NewLine);
        var p = Process.Start(@"D:\911\ProxyTool\AutoProxyTool.exe",
            $"-changeproxy/GB -proxyport={port} -hwnd={Handle}");
    }

    private bool checkingEmail;

    private void CheckImapEmail()
    {
        textBox4.Clear();

        checkingEmail = true;
        button1.Enabled = false;
        button2.Enabled = true;

        var client = new ImapClient();
        var emailInfo = emailInput.Text.Split("----");
        client.Connect("outlook.office365.com", 993, true);
        client.Authenticate(emailInfo[0], emailInfo[1]);
        var inbox = client.Inbox;
        while (checkingEmail)
        {
            inbox.Open(FolderAccess.ReadOnly);
            for (var i = 0; i < inbox.Count; i++)
            {
                var message = inbox.GetMessage(i);
                if (message.From.Mailboxes.ToList()[0].Address == "noreply@messaging.squareup.com")
                {
                    if (message.Subject == "Please confirm your email address to accept payments" &&
                        !textBox4.Text.Contains("账户验证地址: "))
                    {
                        var regex = new Regex(@"https://squareup.com/verify\?\S*");
                        textBox4.AppendText("账户验证地址: " + Environment.NewLine +
                                            regex.Match(message.TextBody).Value);
                        textBox4.AppendText(Environment.NewLine);
                    }

                    if (message.Subject.StartsWith("Bank Account Verification in Progress"))
                    {
                        var regex = new Regex(@"\S*\s\d+");
                        textBox4.AppendText("银行验证日期: " + Environment.NewLine +
                                            regex.Match(message.Subject).Value);
                        textBox4.AppendText(Environment.NewLine);
                        StopReg();
                    }
                }
            }

            Thread.Sleep(10000);
        }

        client.Disconnect(true);
    }

    private static string GenerateTelNo()
    {
        var telNo = "74";
        var rd = new Random();
        for (var i = 0; i < 8; i++)
        {
            telNo += rd.Next(0, 10);
        }

        return telNo;
    }

    private void button1_Click(object sender, EventArgs e)
    {
        if (emailInput.Text == "")
        {
            return;
        }

        textBox1.Text = @"电话：" + Environment.NewLine + GenerateTelNo();

        emailInput.ReadOnly = true;
        Apply911Proxy();
        Task.Run(CheckImapEmail);
    }

    private void StopReg()
    {
        emailInput.ReadOnly = false;
        checkingEmail = false;
        button1.Enabled = true;
        button2.Enabled = false;
    }

    private void button2_Click(object sender, EventArgs e)
    {
        StopReg();
    }

    private void CheckAccountsStatus()
    {
        textBox4.Clear();
        var emails = textBox2.Text.Split(Environment.NewLine);
        foreach (var email in emails)
        {
            var status = "未验证？？？";
            var isChecked = false;
            var emailInfo = email.Split("----");
            var client = new ImapClient();
            client.Connect("outlook.office365.com", 993, true);
            client.Authenticate(emailInfo[0], emailInfo[1]);
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadOnly);
            for (var i = 0; i < inbox.Count; i++)
            {
                var message = inbox.GetMessage(i);
                if (message.From.Mailboxes.ToList()[0].Address == "noreply@messaging.squareup.com")
                {
                    if (message.Subject == "Your bank account has been verified successfully")
                    {
                        status = "已验证。。。";
                    }
                }

                if (message.From.Mailboxes.ToList()[0].Address == "noreply@help-messaging.squareup.com")
                {
                    if (message.Subject.StartsWith("Urgent: "))
                    {
                        status = "挂了！！！";
                        break;
                    }
                }
            }

            textBox4.AppendText($"{emailInfo[0]}: {status}");
            textBox4.AppendText(Environment.NewLine);

            client.Disconnect(true);
        }

        button3.Enabled = true;
        textBox2.ReadOnly = false;
    }

    private void button3_Click(object sender, EventArgs e)
    {
        button3.Enabled = false;
        textBox2.ReadOnly = true;

        Task.Run(CheckAccountsStatus);
    }
}