using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;


namespace Myhut_Automation
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.Write("Especifique o ficheiro XML com os dados das aulas a inscrever.");
                Console.ReadKey();
                return;
            }
           
            Utils utils= new Utils();

            MyHut myhut = new MyHut();

            Credenciais credenciais = myhut.LoadCredenciais(args[0]);
            SmtpCredenciais smtpCredenciais = myhut.LoadSmtpCredenciais(args[0]);

            // carrega as aulas do corrente dia
            List<Aula> listaAulas = myhut.LoadInscricaoAulas(args[0], DateTime.Today);

            if (listaAulas == null || listaAulas.Count== 0)
            {   // se não há aulas para hoje, schedule para o dia seguinte
                utils.sendEMailThroughSMTP(smtpCredenciais, credenciais.Email, "MyHut",
                    string.Format("Não há aulas agendadas para hoje {0}.", DateTime.Today.ToString("yyyy-MM-dd")));
                myhut.ScheduleNextTask(null, DateTime.Today.AddDays(1), credenciais);
                return;
            }

            int nErrors = 0;
            IWebDriver driver= null;
            // Executa até alguma condição se verificar
            while (nErrors< 5)
            {
                try
                {
                    // Navega para o site, faz login e obtem o mapa de aulas
                    driver = myhut.StartBrowser();
                    myhut.Login(driver, credenciais.Email, credenciais.Password);
                    myhut.MapaDeAulas(driver);

                    while (listaAulas.Count> 0)
                    {
                        DateTime horaPrimeiraInscricao = listaAulas[0].HoraInscricao;
                        if (DateTime.Compare(horaPrimeiraInscricao, DateTime.Now) > 0)
                        {   // schedule para a hora de inscrição na primeira aula
                            utils.sendEMailThroughSMTP(smtpCredenciais, credenciais.Email, "MyHut",
                                string.Format("Inscrição da aula {0} agendada para hoje às {1}", listaAulas[0].ToString(), horaPrimeiraInscricao.ToString("HH:mm")));
                            myhut.ScheduleNextTask(driver, horaPrimeiraInscricao, credenciais);
                            return;
                        }

                        // Refresca a página corrente - usado no mapa de aulas para refrescar as aulas reservadas
                        myhut.RefreshCurrentPage(driver);

                        int nAulas = listaAulas.Count;

                        for (int i = 0; i < nAulas; i++)
                        {
                            Aula aula = listaAulas[i];

                            myhut.RefreshDisponilidadeAulas(driver);

                            // Se a aula ainda não está reservada e ainda não passou da hora da aula - 1 hora
                            if (!myhut.IsAulaReservada(driver, aula) && DateTime.Compare(aula.HoraAula, DateTime.Now.AddHours(-1)) > 0)
                                // se não está, tenta reservá-la
                                myhut.ReservarAula(driver, aula);
                            else
                            {
                                // remove a aula da lista
                                listaAulas.Remove(aula);
                                nAulas--;
                                i--;
                            }
                        }
                    }
                    // a lista fica com Count== 0 quando as aulas estão todas reservadas/ passadas
                    // schedule para o dia seguinte
                    utils.sendEMailThroughSMTP(smtpCredenciais, credenciais.Email, "MyHut",
                        string.Format("Todas as aulas marcadas para hoje {0}.", DateTime.Today.ToString("yyyy-MM-dd")));
                    myhut.ScheduleNextTask(driver, DateTime.Today.AddDays(1), credenciais);
                    return;
                }
                catch (Exception ex)
                {
                    nErrors++;
                    utils.Log(NLog.LogLevel.Error, string.Format("Tentativa {0}: Aplicação dessincronizada com o site MyHut. {1}", nErrors, ex.Message));
                    if (driver != null) try { driver.Quit(); } catch { } // para remover o driver dos serviços
                    Thread.Sleep(15 * 1000); // espera 15 segundos
                }
            }
            utils.sendEMailThroughSMTP(smtpCredenciais, credenciais.Email, "MyHut",
                string.Format("Aplicação MyHut_Automation interrompida no dia {0}.", DateTime.Today.ToString("yyyy-MM-dd")));
            myhut.ScheduleNextTask(null, DateTime.Now.AddHours(1), credenciais);
        }
    }
}
