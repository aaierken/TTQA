using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GlobalLibrary
{
    class GlobalMethods
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toAddress"></param>
        /// <param name="fromAddress"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="attachmentPath"></param>
        /// <param name="userName"></param>
        /// <param name="userPassword"></param>
        /// <returns></returns>
        public bool SendEmail(string toAddress, string fromAddress, string subject, string body, string attachmentPath, string userName, string userPassword)
        {
            bool bRtn = true;

            try
            {
                System.Net.Mail.MailAddress SendFrom = new System.Net.Mail.MailAddress(fromAddress);
                System.Net.Mail.MailAddress SendTo = new System.Net.Mail.MailAddress(toAddress);
                System.Net.Mail.MailMessage MyMessage = new System.Net.Mail.MailMessage(SendFrom, SendTo);
                MyMessage.Subject = subject;
                MyMessage.Body = body;
                System.Net.Mail.Attachment attachFile = new System.Net.Mail.Attachment(attachmentPath);
                MyMessage.Attachments.Add(attachFile);
                System.Net.Mail.SmtpClient emailClient = new System.Net.Mail.SmtpClient("smtp.gmail.com");
                emailClient.Credentials = new System.Net.NetworkCredential(userName, userPassword);
                emailClient.Port = 587;
                emailClient.EnableSsl = true;
                emailClient.Send(MyMessage);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }




        public bool SendEmail(string toAddress, string subject, string body, string attachmentPath)
        {
            bool bRtn = true;
            string fromAddress = "Automation@Musabay.com";
            string userName = GlobalVariables.SendEmailUserName;
            string userPassword = GlobalVariables.SendEmailUserPassword;
            string smtpClient = "smtp.gmail.com";
            //string attachmentpath = attachmentPath;

            try
            {
                System.Net.Mail.MailAddress SendFrom = new System.Net.Mail.MailAddress(fromAddress);
                System.Net.Mail.MailAddress SendTo = new System.Net.Mail.MailAddress(toAddress);
                System.Net.Mail.MailMessage MyMessage = new System.Net.Mail.MailMessage(SendFrom, SendTo);
                MyMessage.Subject = subject;
                MyMessage.Body = body;
                System.Net.Mail.Attachment attachFile = new System.Net.Mail.Attachment(attachmentPath);
                MyMessage.Attachments.Add(attachFile);
                System.Net.Mail.SmtpClient emailClient = new System.Net.Mail.SmtpClient(smtpClient);
                emailClient.Credentials = new System.Net.NetworkCredential(userName, userPassword);
                emailClient.Port = 587;
                emailClient.EnableSsl = true;
                emailClient.Send(MyMessage);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                bRtn = false;
            }

            return bRtn;
        }
    }


}