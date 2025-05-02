using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO;

namespace PSReport
{
    /// <summary>
    /// Interaction logic for ScriptDialog.xaml
    /// </summary>
    public partial class ScriptDialog : Window
    {
        public Script SelectedScript { get; set; }

        // constructor used for adding new Script
        public ScriptDialog()
        {
            InitializeComponent();
            Script_Enabled.IsChecked = true;
        }

        // constructor used for editing Script
        public ScriptDialog(Script selectedScript)
        {
            SelectedScript = selectedScript;

            InitializeComponent();

            Script_Title.Text = SelectedScript.Title;
            Script_Path.Text = SelectedScript.Path;
            Script_Enabled.IsChecked = SelectedScript.Enabled;

        }

        private void OK_Btn_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedScript == null)
            {
                Add_Script();
            }
            else
            {
                Edit_Script();
            }

            this.Close();
        }

        private void Cancel_Btn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Browse_Btn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PowerShell script (*.ps1)|*.ps1|All files (*.*)|*.*";

            if(String.IsNullOrEmpty(Script_Path.Text))
            { 
                openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory + "scripts";
            }
            else
            {
                openFileDialog.InitialDirectory = Path.GetFullPath(Script_Path.Text);
            }
            

            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true)
            {
                Script_Path.Text = ConvertTo_RelativePath(openFileDialog.FileName);
            }
        }

        // To Add new Server
        private void Add_Script()
        {
            MainWindow mainWindow = (MainWindow)this.Owner;
            Script NewScript = new Script(Script_Title.Text, Script_Path.Text, (Boolean)Script_Enabled.IsChecked);

            mainWindow.Config.Scripts.Add(NewScript);
        }

        // To edit selected server
        private void Edit_Script()
        {
            MainWindow mainWindow = (MainWindow)this.Owner;

            SelectedScript.Title = Script_Title.Text;
            SelectedScript.Path = Script_Path.Text;
            SelectedScript.Enabled = (Boolean)Script_Enabled.IsChecked;

            mainWindow.Script_List.Items.Refresh();
        }

        public string ConvertTo_RelativePath(string Path)
        {
            if (Path.Contains(AppDomain.CurrentDomain.BaseDirectory))
            {
                string RelativePath = Path.Replace(AppDomain.CurrentDomain.BaseDirectory, @".\");
                return RelativePath;
            }
            else
            {
                return Path;
            }
        }
    }
}
