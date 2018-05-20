using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Security.Principal;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Win32.TaskScheduler;


namespace TaskScheduler
{
    public partial class MainWindow : Window
    {
        ObservableCollection<string> taskList = new ObservableCollection<string>();
        TriggerEditDialog dlg = new TriggerEditDialog(null, false, AvailableTriggers.AllTriggers);
        public MainWindow()
        {
            InitializeComponent();
            LoadTasks();
            Timer timer1 = new Timer();
            timer1.Tick += new EventHandler(TriggerBoxText);
            timer1.Interval = 500;
            timer1.Start();
        }
        public void CreateTask(string FilePath, string desc, string TaskName)
        {
            using (TaskService ts = new TaskService())
            {
                TaskDefinition td = ts.NewTask();
                td.RegistrationInfo.Description = desc;
                td.Triggers.Add(dlg.Trigger);
                td.Actions.Add(new ExecAction(FilePath));
                ts.RootFolder.RegisterTaskDefinition(TaskName, td);
            };
            LoadTasks();
        }
        public void LoadTasks()
        {
            taskList = new ObservableCollection<string>();
            using (TaskService ts = new TaskService())
            {
                TaskFolder taskFolder = ts.RootFolder;
                foreach (Task task in taskFolder.Tasks)
                {
                    taskList.Add(task.Name + "- Next run at " + task.NextRunTime.ToString());
                }
            }
            TaskListView.ItemsSource = taskList;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new System.Windows.Forms.OpenFileDialog();
            var result = fileDialog.ShowDialog();
            switch (result)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var file = fileDialog.FileName;
                    FilePathInput.Text = file;
                    FilePathInput.ToolTip = file;
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    FilePathInput.Text = null;
                    FilePathInput.ToolTip = null;
                    break;
            }
        }

        private void RefreshList(object sender, RoutedEventArgs e)
        {
            LoadTasks();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string error = "";
            bool errorCheck = false;
            if (string.IsNullOrEmpty(FilePathInput.Text))
            {
                error += "No file chosen.";
                errorCheck = true;
            }
            if (string.IsNullOrEmpty(TaskNameInput.Text))
            {
                error += "Missing task name.";
                errorCheck = true;
            }
            if (errorCheck)
            {
                ErrorText.Content = error;
            }
            else
            {
                ErrorText.Content = error;
                CreateTask(FilePathInput.Text, DescriptionInput.Text, TaskNameInput.Text);
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            dlg = new TriggerEditDialog(null, false, AvailableTriggers.AllTriggers);
            dlg.Show();
        }
        private void TriggerBoxText(object sender, EventArgs e)
        {
             TriggerName.Text = dlg.Trigger.ToString();
        }
    }
}
