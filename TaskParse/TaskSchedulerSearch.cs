//using Microsoft.Win32.TaskScheduler;
using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
//using System.Threading.Tasks;
using System.Windows.Forms;
using TaskScheduler;

namespace TaskParse
{
    public partial class TaskSchedulerSearch : Form
    {
        private List<Regex> ReleventActions;
        private string _searchTerm;
        private Regex WindowsFileMatch;
        public TaskSchedulerSearch()
        {
            ReleventActions = new List<Regex>();
            InitializeComponent();

            //R file search in tasks
            ReleventActions.Add(new Regex(@"[a-zA-Z]:[\\\/](?:[a-zA-Z0-9]+[\\\/])*([a-zA-Z0-9]+\.R)"));

            //VBS file search in tasks
            ReleventActions.Add(new Regex(@"[a-zA-Z]:[\\\/](?:[a-zA-Z0-9]+[\\\/])*([a-zA-Z0-9]+\.vbs)"));


        }

        private async void button1_Click(object sender, EventArgs e)
        {
            this._searchTerm = textBox1.Text;
            using (TaskService ts = new TaskService())
            {
                
                foreach (Task task in ts.AllTasks)
                {
                    //this.listBox1.Items.Add($"Parsing {task.Name}");
                    CheckTask(task.Definition);
                    
                    this.Invalidate();
                }
            }
                
        }

     private void CheckTask(TaskDefinition TD)
        {
            ActionCollection actions = TD.Actions;
            
            foreach(Microsoft.Win32.TaskScheduler.Action action in actions)
            {
                foreach(Regex releventAction in ReleventActions)
                {
                    MatchCollection matches = releventAction.Matches(action.ToString());
                    if (matches.Count > 0)
                        ParseActionFile(matches.First());
                }
                //this.listBox1.Items.Add(action.ToString());
                
            }
        }
        private void ParseActionFile(Match matchedActionStr)
        {
            try
            {
                //Parse text file for matched text
                string[] parsedLines = File.ReadAllLines(matchedActionStr.Value);
                foreach(string line in parsedLines)
                {
                    if (line.Contains(this._searchTerm))
                    {
                        this.listBox1.Items.Add(matchedActionStr.Value);
                    }
                }

            }
            catch (IOException)
            {
                //Try to parse SAS Project

            }
        }
    }

}
