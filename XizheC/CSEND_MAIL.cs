using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Web.Mail;
namespace XizheC
{
    public class CSEND_MAIL
    {
        public CSEND_MAIL()
        {

        }
        private void SEND_MAL()
        {
            /*string date_yy = DateTime.Now.Date.ToString("yyyy");
            string date_mm = DateTime.Now.Date.ToString("MM");
            string date_dd = DateTime.Now.Date.ToString("dd");
            string date = date_yy + date_mm + date_dd;
            string file = @"D:\" + date + ".xls";
            MailAddress from = new MailAddress("uay022@126.com");
            MailMessage message = new MailMessage();
            message.From = from;
            string[] a = new string[] { "uay022@126.com" };
            for (int i = 0; i < a.Length; i++)
            {
                message.To.Add(a[i]);
            }
            message.Subject = "Using the new SMTP client.";
            message.Body = @"
此类型的任何公共静态（Visual Basic 中的 Shared）成员都是线程安全的，但不保证所有实例成员都是线程安全的。
此类型的任何公共静态（Visual Basic 中的 Shared）成员都是线程安全的，
但不保证所有实例成员都是线程安全的。
此类型的任何公共静态（Visual Basic 中的 Shared）成员都是线程安全的，但不保证所有实例成员都是线程安全的。
";
            // Create  the file attachment for this e-mail message.
            Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);
            // Add time stamp information for the file.
            ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = System.IO.File.GetCreationTime(file);
            disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
            disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
            // Add the file attachment to this e-mail message.
            message.Attachments.Add(data);
            //Send the message.
            SmtpClient client = new SmtpClient("mail.126.com");
            // Add credentials if the SMTP server requires them.
            client.Credentials = CredentialCache.DefaultNetworkCredentials;
            try
            {
                client.Send(message);
            }
            catch (Exception)
            {

            }*/

        }
        public void SetToMail(string MailTo, string Subject, string Mailaddress, string MailBody)
        {
            MailMessage objMailMessage = new MailMessage();
            objMailMessage.From = Mailaddress;//源邮件地址 
            objMailMessage.To = MailTo;//目的邮件地址，也就是发给我哈 
            objMailMessage.Subject = Subject;//发送邮件的标题 
            objMailMessage.Body = MailBody;//发送邮件的内容 
            SmtpMail.SmtpServer = "smtp3.126.com";
            //开始发送邮件 
            SmtpMail.Send(objMailMessage);
        }
    }
    
}
