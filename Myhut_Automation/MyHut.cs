using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Diagnostics;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace Myhut_Automation
{
    class SmtpCredenciais 
    {
        public string SmtpAddress { get; set; }
        public int SmtpPortNumber { get; set; }
        public bool SmtpEnableSSL { get; set; }
        public string SmtpEmail { get; set; }
        public string SmtpPassword { get; set; }

        public SmtpCredenciais(string smtpAddress, int smtpPortNumber, bool smtpEnableSSL,
            string smtpEmail, string smtpPassword)
        {
            SmtpAddress = smtpAddress;
            SmtpPortNumber = smtpPortNumber;
            SmtpEnableSSL = smtpEnableSSL;
            SmtpEmail = smtpEmail;
            SmtpPassword = smtpPassword;
        }

        public override string ToString()
        {
            return string.Format("SmtpAddress= '{0}'; PortNumber= '{1}'; EnableSSL= '{2}'; Email= '{3}'; Password= '******'",
                SmtpAddress, SmtpPortNumber, SmtpEnableSSL, SmtpEmail, SmtpPassword);
        }

    }

    class Credenciais
    {
        public string ConfigFile { get; set; }
        public string Directory  { get { return Path.GetDirectoryName(Path.GetFullPath(ConfigFile)); } }

        public string Atleta { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }

        public Credenciais(string atleta, string email, string password, string configFile)
        {
            Atleta = atleta;
            Email = email;
            Password = password;
            ConfigFile = configFile;
        }

    }

    class Aula
    {
        public string Nome { get; set; }
        public string Hora { get; set; }
        public string Estudio { get; set; }
        public string Duracao { get; set; }
        public DateTime HoraAula
        {
            get
            {
                string todayString = DateTime.Today.ToString("yyyy-MM-dd");
                string datetimeString = string.Format("{0} {1}", todayString, Hora);
                return Convert.ToDateTime(datetimeString);
            }
        }
        public DateTime HoraInscricao
        {
            get
            {
                if (string.Compare(Hora, "10:00") > 0)
                    return HoraAula.AddHours(-10);
                else
                    return DateTime.Today;

            }
        }

        public Aula(string nome, string hora, string estudio, string duracao)
        {
            Nome = nome;
            Hora = hora;
            Estudio = "ESTÚDIO " + estudio;
            Duracao = duracao;
        }

        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3}", this.Hora, this.Nome, this.Estudio, this.Duracao);
        }
    }

    class MyHut
    {

        private Utils utils = new Utils();

        private void WaitForPageLoad(IWebDriver driver, int maxWaitTimeInSeconds) 
        {
            string state = string.Empty;
            try {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

                //Checks every 500 ms whether predicate returns true if returns exit otherwise keep trying till it returns ture
                wait.Until(d => {

                    try {
                        state = ((IJavaScriptExecutor) driver).ExecuteScript(@"return document.readyState").ToString();
                    } catch (InvalidOperationException) {
                        //Ignore
                    } catch (NoSuchWindowException) {
                        //when popup is closed, switch to last windows
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                    }
                    //In IE7 there are chances we may get state as loaded instead of complete
                    return (state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase));

                });
            } catch (TimeoutException) {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            } catch (NullReferenceException) {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            } catch (WebDriverException) {
                if (driver.WindowHandles.Count == 1) {
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                }
                state = ((IJavaScriptExecutor) driver).ExecuteScript(@"return document.readyState").ToString();
                if (!(state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase)))
                    throw;
            }
        }

        // Devolve as credenciais de acesso ao MyHut
        public Credenciais LoadCredenciais(string filepath)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load(filepath);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR loading configuration file! " + ex.Message);
                throw (ex);
            }

            try
            {
                XmlNode credNode = xDoc.SelectSingleNode("/MyHut/Credenciais");
                string atleta = credNode.SelectSingleNode("Atleta").InnerText;
                string login = credNode.SelectSingleNode("Email").InnerText;
                string password = credNode.SelectSingleNode("Password").InnerText;
                return new Credenciais(atleta, login, password, filepath);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR loading Credenciais from configuration file! "  +ex.Message);
                throw (ex);
            }
        }

        public SmtpCredenciais LoadSmtpCredenciais(string filepath)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load(filepath);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Warn, "Error loading configuration file! " + ex.Message);
                return null;
            }

            try
            {
                XmlNode credNode = xDoc.SelectSingleNode("/MyHut/Definicoes/SmtpCredenciais");
                string smtpAddress = credNode.SelectSingleNode("SmtpAddress").InnerText;
                int portNumber = int.Parse(credNode.SelectSingleNode("PortNumber").InnerText);
                bool enableSSL = bool.Parse(credNode.SelectSingleNode("EnableSSL").InnerText);
                string email = credNode.SelectSingleNode("Email").InnerText;
                string password = credNode.SelectSingleNode("Password").InnerText;
                return new SmtpCredenciais(smtpAddress, portNumber, enableSSL, email, password);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR loading SmtpCredenciais from configuration file! " + ex.Message);
                return null;
            }
        }

        // Devolve a lista de aulas a inscrever no dia de hoje, ordenada por hora
        public List<Aula> LoadInscricaoAulas(string filepath, DateTime day)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            List<Aula> listAula = new List<Aula>();

            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load(filepath);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR loading configuration file! " + ex.Message);
                throw (ex);
            }

            string weekday = DateTime.Today.DayOfWeek.ToString();
            try
            {
                XmlNodeList aulaNodes= xDoc.SelectNodes(string.Format("/MyHut/MapaAulas/DiaSemana[@id='{0}' and not(@inativo)]/Aula", weekday));
                foreach (XmlNode aulaNode in aulaNodes)
                {
                    XmlAttribute inativo = aulaNode.Attributes["inativo"];
                    if (inativo == null)
                    {
                        string nome = aulaNode.SelectSingleNode("Nome").InnerText;
                        string hora = aulaNode.SelectSingleNode("Hora").InnerText;
                        string estudio = aulaNode.SelectSingleNode("Estudio").InnerText;
                        string duracao = aulaNode.SelectSingleNode("Duracao").InnerText;

                        Aula aula = new Aula(nome, hora, estudio, duracao);

                        listAula.Add(aula);

                        utils.Log(NLog.LogLevel.Debug, string.Format("Aula carregada: {0}", aula.ToString()));
                    }
                }

            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR loading Aulas from configuration file! " + ex.Message);
                throw (ex);
            }

            return listAula;
        }

        public IWebDriver StartBrowser()
        {
            IWebDriver driver= new FirefoxDriver();
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));

            driver.Navigate().GoToUrl("https://myhut.pt");
            return driver;
        }

        public void Login(IWebDriver driver, string login, string password)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            WebDriverWait waitLogin = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement loginField = waitLogin.Until<IWebElement>(d => d.FindElement(By.Id("myhut-login-email")));
            loginField.Clear();
            loginField.SendKeys(login);

            WebDriverWait waitPassword = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement passwordField = waitPassword.Until<IWebElement>(d => d.FindElement(By.Id("myhut-login-password")));
            passwordField.Clear();
            passwordField.SendKeys(password);

            driver.FindElement(By.Id("b-login-form")).Click();
        }

        public void Logout(IWebDriver driver)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            driver.Navigate().GoToUrl("https://myhut.pt/myhut/logout/");
        }

        // adormece até à hora indicada, fazendo logout e login novamente
        public void ScheduleNextTask(IWebDriver driver, DateTime hora, Credenciais cred)
        {
            utils.Log(NLog.LogLevel.Debug, hora.ToString("yyyy-MM-dd HH:mm"));

            if (driver != null)
            {
                Logout(driver);
                driver.Quit();
            }

            MyHutScheduler schedule = new MyHutScheduler(cred.Atleta, cred.ConfigFile, cred.Directory);
            schedule.SetMyHutTask(hora);
        }

        // Reserva a aula
        public void MapaDeAulas(IWebDriver driver)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                IWebElement reservar = wait.Until<IWebElement>(d => d.FindElement(By.LinkText("Reservar Aulas")));

                reservar.Click();
                WaitForPageLoad(driver, 10);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR getting Mapa de Aulas! " + ex.Message);
                throw (ex);
            }
        }

        // Refresca a página corrente - usado no mapa de aulas para refrescar as aulas reservadas
        public void RefreshCurrentPage(IWebDriver driver)
        {
            utils.Log(NLog.LogLevel.Debug, "");
            try
            {
                driver.Navigate().Refresh();
                WaitForPageLoad(driver, 10);
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR refreshing página Mapa de Aulas! " + ex.Message);
                throw (ex);
            }
        }

        // Refresca a disponibilidade das aulas no mapa de aulas. Este refresh NÃO refresca as aulas reservadas!
        public void RefreshDisponilidadeAulas(IWebDriver driver)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                IWebElement refresh = wait.Until<IWebElement>(d => d.FindElement(By.Id("b-refresh")));

                refresh.Click(); // pode dar erro de Unexpected error. Element is not clickable at point. Other element would receive the click <div>
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR refreshing Disponibilidades - continuing! " + ex.Message);
            }
        }

        // Verifica se a aula já está reservada
        public bool IsAulaReservada(IWebDriver driver, Aula aula)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                IWebElement aulasReservadas = wait.Until<IWebElement>(d => d.FindElement(By.Id("aulas-reservadas-holder")));

                string innerHtml = aulasReservadas.Text;
                // procura todos os elementos da aula no container
                if (innerHtml.Contains(aula.Nome) && innerHtml.Contains(aula.Hora)
                    && innerHtml.Contains(aula.Estudio) && innerHtml.Contains(aula.Duracao))
                {
                    utils.Log(NLog.LogLevel.Debug, string.Format("Aula está reservada: {0}", aula.ToString()));
                    return true;
                }
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, "ERROR getting Aulas Reservadas status! " + ex.Message);
                throw (ex);
            }
            return false;
        }

        // Reserva a aula
        public void ReservarAula(IWebDriver driver, Aula aula)
        {
            utils.Log(NLog.LogLevel.Debug, "");

            IWebElement aulasContainer = null;
            IWebElement aulaLink = null;

            try
            {
                // procura os dados da aula num mesmo elemento de link <a/> com multiplos <div/>
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                aulasContainer = wait.Until<IWebElement>(d => driver.FindElement(By.Id("aulas-holder")));
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Debug, "ERROR getting Aulas Container! " + ex.Message);
                throw (ex);
            }

            try
            {
                aulaLink = utils.FindAulaLink(aulasContainer, aula);

                if (aulaLink != null)
                {
                    // abre a aula para reservar
                    aulaLink.Click();

                    // reserva a aula
                    utils.ClickWhenReady(driver, By.LinkText("RESERVAR AULA"), TimeSpan.FromSeconds(15));

                    // fecha a aula
                    aulaLink.Click();
                }
            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Warn, string.Format("Aula não disponível para reservar: {0}", aula.ToString()) + ex.Message);
            }
        }
    }
}
