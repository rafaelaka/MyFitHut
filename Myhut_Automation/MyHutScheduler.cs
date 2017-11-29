using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32.TaskScheduler;


namespace Myhut_Automation
{
    class MyHutScheduler
    {
        public string ConfigFile { get; set; }
        public string Directory { get; set; }
        public string Atleta { get; set; }
        public string TaskName { get { return "MyHut_" + Atleta; } }
        public string TaskFolder { get { return "MyHut"; } }

        private Utils utils = new Utils();

        public MyHutScheduler (string atleta, string configFile, string directory)
        {
            Atleta = atleta;
            ConfigFile = configFile;
            Directory = directory;
        }

        public void DeleteTriggers(TaskDefinition def)
        {
            utils.Log(NLog.LogLevel.Debug, TaskName);

            if (def == null || def.Triggers == null || def.Triggers.Count== 0) return;

            for (int i= def.Triggers.Count-1; i>= 0; i--)
                def.Triggers.Remove(def.Triggers[i]);
        }

        public void DeleteActions(TaskDefinition def)
        {
            utils.Log(NLog.LogLevel.Debug, TaskName);

            if (def == null || def.Actions == null || def.Actions.Count== 0) return;

            for (int i = def.Actions.Count - 1; i >= 0; i--)
                def.Actions.Remove(def.Actions[i]);
        }

        public Task GetTask()
        {
            utils.Log(NLog.LogLevel.Debug, TaskName);

            Task task = null;
            TaskService ts = new TaskService();
            TaskFolder tf = ts.GetFolder(TaskFolder);
            if (tf != null)
                task = ts.GetFolder(TaskFolder).GetTasks().Where(a => a.Name.ToLower() == TaskName.ToLower()).FirstOrDefault();

            return task;
        }

        public void SetMyHutTask(DateTime startDate)
        {
            utils.Log(NLog.LogLevel.Debug, string.Format("Schedule next task {0} at {1}", TaskName, startDate.ToString("yyyy-MM-dd HH:mm")));

            TaskService ts = new TaskService();
            TaskFolder tf = ts.GetFolder(TaskFolder);
            if (tf == null)
                tf = ts.RootFolder.CreateFolder(TaskFolder);

            Task task= GetTask();

            TaskDefinition def= null;
            if (task == null)
            {
                utils.Log(NLog.LogLevel.Warn, string.Format("Crie uma task com o nome '{0}' na pasta '{1}' no Task Scheduler.", TaskName, TaskFolder));
                def = ts.NewTask();
            }
            else
                def = task.Definition;

            try
            {
                DeleteTriggers(def);
                DeleteActions(def);

                ExecAction action = new ExecAction();
                action.Path = Assembly.GetExecutingAssembly().Location;
                action.Arguments = ConfigFile;
                action.WorkingDirectory = Directory;
                def.Actions.Add(action);

                TimeTrigger trigger = new TimeTrigger(startDate);
                trigger.Enabled = true;
                def.Triggers.Add(trigger);

                def.Principal.LogonType = TaskLogonType.ServiceAccount;
                def.Principal.UserId = "SYSTEM";
                def.Principal.RunLevel = TaskRunLevel.LUA;

                def.Settings.DisallowStartIfOnBatteries= true;
                def.Settings.StopIfGoingOnBatteries= true;
                def.Settings.AllowHardTerminate= true;
                def.Settings.StartWhenAvailable= false;
                def.Settings.RunOnlyIfNetworkAvailable= false;
                def.Settings.IdleSettings.StopOnIdleEnd= true;
                def.Settings.IdleSettings.RestartOnIdle= false;
                def.Settings.AllowDemandStart= true;
                def.Settings.Enabled= true;
                def.Settings.Hidden= false;
                def.Settings.RunOnlyIfIdle= false;
                def.Settings.WakeToRun= true;

                // Este ação requere direitos de Administrador. Executar a aplicação com "Run as Administrator".
                tf.RegisterTaskDefinition(TaskName, def);

            }
            catch (Exception ex)
            {
                utils.Log(NLog.LogLevel.Error, string.Format("ERROR scheduling next MyHut Task {0} - {1}", def.XmlText, ex.Message));
                throw (ex);
            }
        }

    }
}
