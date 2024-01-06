 using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using Newtonsoft.Json;

namespace Mail
{
public class Notify
    {
        private static string senderEmailAddress = "anudeepreddyspear@gmail.com";
        private static string senderPassword = "wumdwzlyqtcrfang";

        public static bool Credentials(string add,string pass)
        {
            if(validEmail(add)){
                senderEmailAddress = add;
                senderPassword = pass;
                return true;
            }
            else
            {
                Console.WriteLine("Invalid Email Address");
                return false;
            }
        }
        public static void SendMail(string jsonString)
        {
            dynamic data = ParseAndValidateJson(jsonString);
            if (data == null)
            {
                Console.WriteLine("Invalid JSON format.");
                return;
            }
            List<string> toList = GetEmailListFromJson(data.to);
            List<string> ccList = GetEmailListFromJson(data.cc);
            List<string> bccList = GetEmailListFromJson(data.bcc);
            string subject = data.subject;
            string body = data.body;
            string signature = data.signature;
            MailMessage message = new MailMessage();
            message.From = new MailAddress(senderEmailAddress);

            foreach (string toEmail in toList)
            {
                if (validEmail(toEmail))
                {
                    message.To.Add(toEmail);
                }
                else
                {
                    Console.WriteLine(toEmail+" is invalid Recipients Email");
                }
            }

            foreach (string ccEmail in ccList)
            {
                if (validEmail(ccEmail))
                {
                    message.To.Add(ccEmail);
                }
                else
                {
                    Console.WriteLine(ccEmail + " is invalid Recipients Email");
                }
            }

            foreach (string bccEmail in bccList)
            {
                if (validEmail(bccEmail))
                {
                    message.To.Add(bccEmail);
                }
                else
                {
                    Console.WriteLine(bccEmail + " is invalid Recipients Email");
                }
            }

            message.Subject = subject;
            message.Body = body + Environment.NewLine + signature;
            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.EnableSsl = true;
            smtpClient.Credentials = new NetworkCredential(senderEmailAddress, senderPassword);

            try
            {
                smtpClient.Send(message);
                Console.WriteLine("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to send email. Error: " + ex.Message);
            }
        }

        private static dynamic ParseAndValidateJson(string jsonString)
        {
            try
            {
                dynamic data = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonString);
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error parsing JSON: " + ex.Message);
                return null;
            }
        }
        private static List<string> GetEmailListFromJson(dynamic emailArray)
        {
            List<string> emailList = new List<string>();
            if (emailArray != null && emailArray is Newtonsoft.Json.Linq.JArray)
            {
                foreach (string email in emailArray)
                {
                        emailList.Add(email);
                }
            }
            foreach (var k in emailList)
            {
                Console.WriteLine(k);
            }
            return emailList;
        }
         static bool validEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}