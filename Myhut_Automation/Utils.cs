using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using NLog;


namespace Myhut_Automation
{
    class Utils
    {
        private static NLog.Logger logger = NLog.LogManager.GetLogger("MyHut");

        public IWebElement FindAulaLink(IWebElement aulasContainer, Aula aula)
        {
            Log(NLog.LogLevel.Debug, "");

            IWebElement linkAula = null;

            try
            {
                // enumera os links dentro do container
                IEnumerator<IWebElement> e = aulasContainer.FindElements(By.TagName("a")).GetEnumerator();

                if (e == null)
                    return linkAula;

                while (e.MoveNext())
                {
                    string innerHtml = e.Current.Text;
                    // procura todos os elementos da aula no texto do link
                    if (innerHtml.Contains(aula.Nome) && innerHtml.Contains(aula.Hora)
                        && innerHtml.Contains(aula.Estudio) && innerHtml.Contains(aula.Duracao))
                    {
                        linkAula = e.Current;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log(NLog.LogLevel.Error, "ERROR finding links in aulasContainer");
                throw (ex);
            }
            return linkAula;
        }

        public void ClickWhenReady(IWebDriver driver, By locator, TimeSpan timeout)
        {
            Log(NLog.LogLevel.Debug, "");

            IWebElement element = null;
            WebDriverWait wait = new WebDriverWait(driver, timeout);
            element = wait.Until(ExpectedConditions.ElementToBeClickable(locator));
            element.Click();
        }

        public void Log(NLog.LogLevel logLevel, object message)
        {
            // frame 1, true for source info
            StackFrame frame = new StackFrame(1, true);
            var method = frame.GetMethod();

            //Console.WriteLine("{0} {1} {2}", method.Name, logLevel.ToString().ToUpper(), message,ToString());
            logger.Log(logLevel, "{0} {1} {2}", method.Name, logLevel.ToString().ToUpper(), message, ToString());
        }

        public void sendEMailThroughSMTP(SmtpCredenciais smtpCredenciais, 
            string emailTo, string subject, string body)
        {
            if (smtpCredenciais == null || string.IsNullOrEmpty(smtpCredenciais.SmtpAddress) || string.IsNullOrEmpty(smtpCredenciais.SmtpEmail) || string.IsNullOrEmpty(smtpCredenciais.SmtpPassword)
                || smtpCredenciais.SmtpPortNumber == 0)
                return;

            Log(NLog.LogLevel.Debug, string.Format("{0} {1}", smtpCredenciais.ToString(), body));
            
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(smtpCredenciais.SmtpEmail);
                    mail.To.Add(emailTo);
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.IsBodyHtml = false;

                    using (SmtpClient smtp = new SmtpClient(smtpCredenciais.SmtpAddress, smtpCredenciais.SmtpPortNumber))
                    {
                        smtp.Credentials = new NetworkCredential(smtpCredenciais.SmtpEmail, smtpCredenciais.SmtpPassword);
                        smtp.EnableSsl = smtpCredenciais.SmtpEnableSSL;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception ex)
            {
                Log(NLog.LogLevel.Error, "ERROR sending SMTP Email " + ex.Message);
            }

        }

    }
}