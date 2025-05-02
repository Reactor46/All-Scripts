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
using System.Windows.Shapes;

namespace PSReport
{
    /// <summary>
    /// Interaction logic for ServerDialog.xaml
    /// </summary>
    public partial class ServerDialog : Window
    {
        public Server SelectedServer { get; set; }

        // constructor used for adding new Server
        public ServerDialog()
        {
            InitializeComponent();  
        }

        // constructor used for editing Server
        public ServerDialog(Server selecedServer)
        {
            SelectedServer = selecedServer;

            InitializeComponent();

            // Pupulate Text and Type textbox
            Server_Name.Text = SelectedServer.Name;
            Server_Type.Text = SelectedServer.Type;

        }

        private void OK_Btn_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedServer == null)
            {
                Add_Server();
            }
            else
            {
                Edit_Server();
            }

            this.Close();
        }

        private void Cancel_Btn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // To Add new Server
        private void Add_Server()
        {
            MainWindow mainWindow = (MainWindow)this.Owner;
            Server NewServer = new Server(Server_Name.Text, Server_Type.Text);
            mainWindow.Config.Servers.Add(NewServer);
        }

        // To edit selected server
        private void Edit_Server()
        {

            SelectedServer.Name = Server_Name.Text;
            SelectedServer.Type = Server_Type.Text;

            MainWindow mainWindow = (MainWindow)this.Owner;
            mainWindow.Server_List.Items.Refresh();
        }

    }
}
