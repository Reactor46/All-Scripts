using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Collections.Specialized;
using Microsoft.Win32;
using System.IO;
using System.Xml.Linq;

namespace PSReport
{

    public partial class MainWindow : Window
    {
        public string XMLPath { get; set; }
        public Configuration Config { get; set; }


        public MainWindow()
        {
            // Set Path to config.xml and load it to object
            XMLPath = AppDomain.CurrentDomain.BaseDirectory + @"config\config.xml";

            try
            {
                Config = Configuration.Load(XMLPath);
            }
            catch  (Exception e) when (e is FileNotFoundException || e is InvalidOperationException)
            {
                MessageBoxResult result = MessageBox.Show("Warning: " + e.Message + " Program will create new 'config.xml' file.", "Warning", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                if (result == MessageBoxResult.OK)
                {
                    create_XMLConfig();
                    Config = Configuration.Load(XMLPath);
                }
                else
                {
                    throw e;
                }
            }

            InitializeComponent();

            // Load XML form file

            load_Configuration();

        }

        public void create_XMLConfig()
        {
            string XMLString = @"<?xml version=""1.0"" encoding=""utf-8""?>
                                <Configuration xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                                  <General>
                                    <FileName />
                                    <ReportName />
                                  </General>
                                  <Email Enabled=""false"">
                                    <SMTPServer />
                                    <Port>0</Port>
                                    <From />
                                    <To />
                                    <Cc />
                                    <Bcc />
                                    <Subject />
                                    <Body />
                                  </Email>
                                  <Scripts />
                                  <Servers />
                                  <Other Enabled=""false"">
                                    <Script Path="""" />
                                  </Other>
                                </Configuration>";

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(XMLString);
            doc.Save(XMLPath);
        }

        // Load various configurations
        public void load_Configuration()
        {
            // Load General
            GeneralProperties.DataContext = this.Config.General;

            // Load Scripts
            Script_List.ItemsSource = this.Config.Scripts;

            // load email
            MailProperties.DataContext = this.Config.Email;

            // load servers
            Server_List.ItemsSource = this.Config.Servers;

            // load other
            OtherProperties.DataContext = this.Config.Other;
        }

        // Server DataGrid button actions
        public void New_Server(object sender, RoutedEventArgs e)
        {

            ServerDialog NewServer = new ServerDialog();
            NewServer.Owner = this;
            NewServer.Top = this.Top + 150;
            NewServer.Left = this.Left + 200;
            NewServer.Title = "New server";

            NewServer.ShowDialog();
        }

        public void Edit_Server(object sender, RoutedEventArgs e)
        {

            ServerDialog EditServer = new ServerDialog((Server)Server_List.SelectedItem);
            EditServer.Owner = this;
            EditServer.Top = this.Top + 150;
            EditServer.Left = this.Left + 200;
            EditServer.Title = "Edit server";

            EditServer.ShowDialog();
        }

        public void Delete_Server(object sender, RoutedEventArgs e)
        {
            if(Server_List.SelectedItem != null)
            {
                Server selectedServer = (Server)Server_List.SelectedItem;
                this.Config.Servers.Remove(selectedServer);
            }
        }

        public void MoveUp_Server(object sender, RoutedEventArgs e)
        {
            if(Server_List.SelectedItem != null )
            {
                Server selectedServer = (Server)Server_List.SelectedItem;
                Int32 currentIndex = this.Config.Servers.IndexOf(selectedServer);
                if(currentIndex != 0 )
                {
                    this.Config.Servers.Remove(selectedServer);
                    this.Config.Servers.Insert(currentIndex - 1, selectedServer);
                    Server_List.SelectedIndex = currentIndex - 1;
                }
                    
            }   
        }

        public void MoveDown_Server(object sender, RoutedEventArgs e)
        {
            if (Server_List.SelectedItem != null)
            {
                Server selectedServer = (Server)Server_List.SelectedItem;
                Int32 currentIndex = this.Config.Servers.IndexOf(selectedServer);
                if (currentIndex != Server_List.Items.Count - 1)
                {
                    this.Config.Servers.Remove(selectedServer);
                    this.Config.Servers.Insert(currentIndex + 1, selectedServer);
                    Server_List.SelectedIndex = currentIndex + 1;
                }

            }
        }

        private void Server_List_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if(Server_List.SelectedItem != null)
            {
                Edit_Server(sender, e);
            }

        }

        // Script DataGrid button actions
        public void New_Script(object sender, RoutedEventArgs e)
        {

            ScriptDialog NewScript = new ScriptDialog();
            NewScript.Owner = this;
            NewScript.Top = this.Top + 150;
            NewScript.Left = this.Left + 200;
            NewScript.Title = "New script";

            NewScript.ShowDialog();
        }

        public void Edit_Script(object sender, RoutedEventArgs e)
        {
            Script selectedScript = (Script)Script_List.SelectedItem;
            ScriptDialog EditScript = new ScriptDialog(selectedScript);
            EditScript.Owner = this;
            EditScript.Top = this.Top + 150;
            EditScript.Left = this.Left + 200;
            EditScript.Title = "Edit script";

            EditScript.ShowDialog();
        }

        public void Delete_Script(object sender, RoutedEventArgs e)
        {
            if (Script_List.SelectedItem != null)
            {
                Script selectedScript = (Script)Script_List.SelectedItem;
                this.Config.Scripts.Remove(selectedScript);
            }
        }

        public void MoveUp_Script(object sender, RoutedEventArgs e)
        {
            if (Script_List.SelectedItem != null)
            {
                Script selectedScript = (Script)Script_List.SelectedItem;
                Int32 currentIndex = this.Config.Scripts.IndexOf(selectedScript);
                if (currentIndex != 0)
                {
                    this.Config.Scripts.Remove(selectedScript);
                    this.Config.Scripts.Insert(currentIndex - 1, selectedScript);
                    Script_List.SelectedIndex = currentIndex - 1;
                }

            }
        }

        public void MoveDown_Script(object sender, RoutedEventArgs e)
        {
            if (Script_List.SelectedItem != null)
            {
                Script selectedScript = (Script)Script_List.SelectedItem;
                Int32 currentIndex = this.Config.Scripts.IndexOf(selectedScript);
                if (currentIndex != Script_List.Items.Count - 1)
                {
                    this.Config.Scripts.Remove(selectedScript);
                    this.Config.Scripts.Insert(currentIndex + 1, selectedScript);
                    Script_List.SelectedIndex = currentIndex + 1;
                }

            }
        }

        private void Script_List_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (Script_List.SelectedItem != null)
            {
                Edit_Script(sender, e);
            }  
        }

        // refreshes Datagrids after update/remove form orginal collections
        private void Refresh_DataGrid(object sender, DataTransferEventArgs e)
        {
            DataGrid passedDataGrid = (DataGrid)sender;

            passedDataGrid.Items.Refresh();
        }

        // Broswe button in Other Tab
        private void Browse_Btn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PowerShell script (*.ps1)|*.ps1|All files (*.*)|*.*";

            openFileDialog.RestoreDirectory = false;

            if (String.IsNullOrEmpty(Other_Script_Path.Text))
            {
                openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory + "scripts";
            }
            else
            {
                openFileDialog.InitialDirectory = Other_Script_Path.Text;
            }


            Nullable<bool> result = openFileDialog.ShowDialog();

            if (result == true)
            {
                Other_Script_Path.Text = ConvertTo_RelativePath(openFileDialog.FileName);
            }
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

        // main buttons actions
        private void Main_Cancel_Btn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Main_OK_Btn_Click(object sender, RoutedEventArgs e)
        {
            Save();
            this.Close();
        }

        private void Main_Apply_Btn_Click(object sender, RoutedEventArgs e)
        {
            Save();
        }

        public void Save()
        {
            this.Config.Save(this.XMLPath);
        }
    }
}
