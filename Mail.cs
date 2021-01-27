using System.Net.Mail;

namespace RoboReestrService
{
    class Mail
    {
        public string SMTPServer;
        public string from;
        public string to;
        public string port;
        public string login;
        public string password;

        public Mail(string smtp, string strfrom, string strto, string strport, string strlogin, string strpassword) 
        {
            SMTPServer = smtp;
            from = strfrom;
            to = strto;
            port = strport;
            login = strlogin;
            password = strpassword;
        }
        public void SendMail(string fileNameUPRID, string fileNameWallet, string strPath)
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient(SMTPServer);

            //email
            mail.From = new MailAddress(from);
            mail.To.Add(to);
            mail.Subject = "Реестры УПРИД + Кошельки";
            mail.Body += "Реестры во вложении.";
            mail.Attachments.Add(new Attachment(strPath + @"\Reestrs\Reestr UPRID " + fileNameUPRID + ".xlsx"));
            mail.Attachments.Add(new Attachment(strPath + @"\Reestrs\Reestr Wallet " + fileNameWallet + ".xlsx"));

            SmtpServer.Port = int.Parse(port);
            SmtpServer.Credentials = new System.Net.NetworkCredential(login, password);

            SmtpServer.EnableSsl = false;
            SmtpServer.Send(mail);
            Logger.Log.Info("Файлы " + fileNameUPRID + " и " + fileNameWallet + " отправлены");
        }
    }
}
